param(
  [Parameter(Mandatory=$true)][string]$Vmrun,
  [Parameter(Mandatory=$true)][string]$Vmx,
  [Parameter(Mandatory=$true)][string]$GuestUser,
  [Parameter(Mandatory=$true)][string]$GuestPassword,
  [Parameter(Mandatory=$true)][string]$GuestWorkDir,
  [Parameter(Mandatory=$true)][string]$HostSkillDir,
  [Parameter(Mandatory=$true)][string]$HostPlanJson,
  [Parameter(Mandatory=$true)][string]$HostEvidenceDir,
  [string]$GuestPython = "python.exe",
  [string]$TaskName = "CodexWindowsVmCheck",
  [int]$PollSeconds = 5,
  [int]$TimeoutSeconds = 900,
  [switch]$KeepTmp
)

$ErrorActionPreference = "Stop"

function Invoke-Vmrun {
  param([Parameter(ValueFromRemainingArguments=$true)][string[]]$Args)
  & $Vmrun -T ws @Args
  if ($LASTEXITCODE -ne 0) {
    throw "vmrun failed: $($Args -join ' ')"
  }
}

function Invoke-Guest {
  param([string]$Program, [string[]]$Args = @())
  Invoke-Vmrun -gu $GuestUser -gp $GuestPassword runProgramInGuest $Vmx $Program @Args
}

function Test-GuestFile {
  param([string]$Path)
  & $Vmrun -T ws -gu $GuestUser -gp $GuestPassword fileExistsInGuest $Vmx $Path | Out-Null
  return $LASTEXITCODE -eq 0
}

function Remove-HostTmp {
  param([string]$EvidenceDir, [string]$TmpDir)
  if (-not (Test-Path -LiteralPath $TmpDir)) {
    return
  }

  $evidenceRoot = (Resolve-Path -LiteralPath $EvidenceDir).Path
  $tmpRoot = (Resolve-Path -LiteralPath $TmpDir).Path
  if (-not $tmpRoot.StartsWith($evidenceRoot, [System.StringComparison]::OrdinalIgnoreCase)) {
    throw "Refusing to remove tmp outside evidence dir: $tmpRoot"
  }

  Remove-Item -LiteralPath $tmpRoot -Recurse -Force
}

New-Item -ItemType Directory -Force -Path $HostEvidenceDir | Out-Null
$hostTmp = Join-Path $HostEvidenceDir "tmp"
Remove-HostTmp -EvidenceDir $HostEvidenceDir -TmpDir $hostTmp
New-Item -ItemType Directory -Force -Path $hostTmp | Out-Null

Invoke-Guest "cmd.exe" @("/c", "mkdir `"$GuestWorkDir`" 2>nul")

$guestScripts = Join-Path $GuestWorkDir "scripts"
Invoke-Guest "cmd.exe" @("/c", "mkdir `"$guestScripts`" 2>nul")

foreach ($name in @("guest_preflight.py", "guest_gui_runner.py", "requirements-guest.txt", "guest_setup_pyautogui.ps1")) {
  Invoke-Vmrun -gu $GuestUser -gp $GuestPassword copyFileFromHostToGuest $Vmx (Join-Path $HostSkillDir "scripts\$name") (Join-Path $guestScripts $name)
}

Invoke-Vmrun -gu $GuestUser -gp $GuestPassword copyFileFromHostToGuest $Vmx $HostPlanJson (Join-Path $GuestWorkDir "plan.json")

# Plain runProgramInGuest is acceptable for preflight import checks, but GUI
# evidence capture should be launched from an interactive scheduled task.
Invoke-Guest $GuestPython @((Join-Path $guestScripts "guest_preflight.py"))

$guestEvidence = Join-Path $GuestWorkDir "evidence"
$runner = Join-Path $guestScripts "guest_gui_runner.py"
$resultJson = Join-Path $GuestWorkDir "runner_result.json"
$keepTmpArg = if ($KeepTmp) { " --keep-tmp" } else { "" }
$taskCommand = "`"$GuestPython`" `"$runner`" --plan `"$GuestWorkDir\plan.json`" --out-dir `"$guestEvidence`" --result-json `"$resultJson`"$keepTmpArg"

Invoke-Guest "schtasks.exe" @("/Create", "/F", "/TN", $TaskName, "/SC", "ONCE", "/ST", "23:59", "/RL", "HIGHEST", "/TR", $taskCommand)
Invoke-Guest "schtasks.exe" @("/Run", "/TN", $TaskName)

$deadline = (Get-Date).AddSeconds($TimeoutSeconds)
while ((Get-Date) -lt $deadline) {
  if (Test-GuestFile $resultJson) {
    break
  }
  Start-Sleep -Seconds $PollSeconds
}

if (-not (Test-GuestFile $resultJson)) {
  throw "Timed out waiting for guest GUI task result: $resultJson"
}

$hostResultJson = Join-Path $hostTmp "runner_result.json"
Invoke-Vmrun -gu $GuestUser -gp $GuestPassword copyFileFromGuestToHost $Vmx $resultJson $hostResultJson
$result = Get-Content -LiteralPath $hostResultJson -Raw | ConvertFrom-Json

foreach ($entry in $result.results) {
  if ($null -eq $entry.evidence) {
    continue
  }
  foreach ($name in @($entry.evidence)) {
    if ([string]::IsNullOrWhiteSpace([string]$name)) {
      continue
    }
    $fileName = [System.IO.Path]::GetFileName([string]$name)
    if ($fileName -notmatch '^row\d{2}_.+\.png$') {
      throw "Refusing to copy non-final evidence file to evidence root: $fileName"
    }
    Invoke-Vmrun -gu $GuestUser -gp $GuestPassword copyFileFromGuestToHost $Vmx (Join-Path $guestEvidence $fileName) (Join-Path $HostEvidenceDir $fileName)
  }
}

Invoke-Guest "schtasks.exe" @("/Delete", "/F", "/TN", $TaskName)

$failedRows = @($result.results | Where-Object { $_.status -ne "captured" } | ForEach-Object { "row$($_.row): $($_.error)" })

if (-not $KeepTmp) {
  Invoke-Guest "cmd.exe" @("/c", "rmdir /s /q `"$guestEvidence\tmp`" 2>nul")
  Remove-HostTmp -EvidenceDir $HostEvidenceDir -TmpDir $hostTmp
}

if ($failedRows.Count -gt 0) {
  throw "Guest GUI task completed with failed rows: $($failedRows -join '; ')"
}

Write-Host "Copied final row screenshots to $HostEvidenceDir"
