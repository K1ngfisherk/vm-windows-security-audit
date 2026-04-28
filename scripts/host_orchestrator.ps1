param(
  [Parameter(Mandatory=$true)][string]$Vmrun,
  [Parameter(Mandatory=$true)][string]$Vmx,
  [Parameter(Mandatory=$true)][string]$GuestUser,
  [Parameter(Mandatory=$true)][string]$GuestPassword,
  [string]$InteractiveGuestUser = "",
  [Parameter(Mandatory=$true)][string]$GuestWorkDir,
  [Parameter(Mandatory=$true)][string]$HostSkillDir,
  [Parameter(Mandatory=$true)][string]$HostPlanJson,
  [Parameter(Mandatory=$true)][string]$HostEvidenceDir,
  [string]$GuestPython = "python.exe",
  [string]$TaskName = "CodexWindowsVmCheck",
  [int]$PollSeconds = 5,
  [int]$TimeoutSeconds = 900,
  [switch]$SkipInteractiveDesktopCheck,
  [switch]$KeepTmp
)

$ErrorActionPreference = "Stop"

if ($PollSeconds -lt 1) {
  throw "PollSeconds must be at least 1"
}
if ($TimeoutSeconds -lt 1) {
  throw "TimeoutSeconds must be at least 1"
}
if (-not (Test-Path -LiteralPath $HostSkillDir -PathType Container)) {
  throw "HostSkillDir does not exist or is not a directory: $HostSkillDir"
}
if (-not (Test-Path -LiteralPath $HostPlanJson -PathType Leaf)) {
  throw "HostPlanJson does not exist or is not a file: $HostPlanJson"
}
$hostPlanFullPath = [System.IO.Path]::GetFullPath($HostPlanJson)
$hostEvidenceFullPath = [System.IO.Path]::GetFullPath($HostEvidenceDir)
$hostTmpFullPath = [System.IO.Path]::GetFullPath((Join-Path $hostEvidenceFullPath "tmp"))
$hostPlanParentPath = [System.IO.Path]::GetFullPath([System.IO.Path]::GetDirectoryName($hostPlanFullPath))
if (-not $hostPlanParentPath.Equals($hostTmpFullPath, [System.StringComparison]::OrdinalIgnoreCase)) {
  throw "HostPlanJson must be staged under the current evidence tmp directory so it is destroyed with runtime files: $hostTmpFullPath"
}
$hostPlanText = Get-Content -LiteralPath $hostPlanFullPath -Raw
$RunAsGuestUser = if ([string]::IsNullOrWhiteSpace($InteractiveGuestUser)) { $GuestUser } else { $InteractiveGuestUser }
$script:GuestTempCmdScripts = [System.Collections.Generic.List[string]]::new()
$script:InteractiveGuestTasks = [System.Collections.Generic.List[object]]::new()

function Invoke-Vmrun {
  param([Parameter(ValueFromRemainingArguments=$true)][string[]]$Args)
  & $Vmrun -T ws @Args
  if ($LASTEXITCODE -ne 0) {
    throw "vmrun failed ($LASTEXITCODE): $($Args -join ' ')"
  }
}

function Invoke-VmrunOutput {
  param([Parameter(ValueFromRemainingArguments=$true)][string[]]$Args)
  $output = @(& $Vmrun -T ws @Args 2>&1 | ForEach-Object { [string]$_ })
  if ($LASTEXITCODE -ne 0) {
    throw "vmrun failed ($LASTEXITCODE): $($Args -join ' ')`n$($output -join [Environment]::NewLine)"
  }
  return $output
}

function Test-GuestFile {
  param([string]$Path)
  & $Vmrun -T ws -gu $GuestUser -gp $GuestPassword fileExistsInGuest $Vmx $Path | Out-Null
  return $LASTEXITCODE -eq 0
}

function Wait-GuestFile {
  param([string]$Path, [int]$Timeout)
  $deadline = (Get-Date).AddSeconds($Timeout)
  while ((Get-Date) -lt $deadline) {
    if (Test-GuestFile $Path) {
      return $true
    }
    Start-Sleep -Seconds $PollSeconds
  }
  return $false
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

function Get-SafeName {
  param([string]$Name)
  return ($Name -replace '[^A-Za-z0-9_.-]', '_')
}

function Remove-GuestFileIfExists {
  param([string]$Path)
  if ([string]::IsNullOrWhiteSpace($Path)) {
    return
  }
  try {
    if (Test-GuestFile $Path) {
      & $Vmrun -T ws -gu $GuestUser -gp $GuestPassword deleteFileInGuest $Vmx $Path | Out-Null
      if ($LASTEXITCODE -ne 0) {
        Write-Warning "Could not remove guest temp file '$Path' (vmrun exit $LASTEXITCODE)"
      }
    }
  } catch {
    Write-Warning "Could not remove guest temp file '$Path': $($_.Exception.Message)"
  }
}

function ConvertTo-BatchValue {
  param([string]$Value, [string]$Label = "value")
  if ($null -eq $Value) {
    return ""
  }
  if ($Value -match '[\r\n]') {
    throw "$Label contains a newline and cannot be embedded in a guest .cmd script"
  }
  if ($Value.Contains('"')) {
    throw "$Label contains a double quote and cannot be safely embedded in a guest .cmd script"
  }
  return $Value.Replace('^', '^^').Replace('%', '%%')
}

function Quote-BatchArg {
  param([string]$Value, [string]$Label = "value")
  return '"' + (ConvertTo-BatchValue -Value $Value -Label $Label) + '"'
}

function Copy-GuestCmdToTemp {
  param([string]$Name, [string[]]$Lines)
  $safeName = Get-SafeName "$TaskName`_$Name.cmd"
  $hostScript = Join-Path $hostTmp $safeName
  $guestScript = "C:\Windows\Temp\$safeName"
  [System.IO.File]::WriteAllLines($hostScript, $Lines, [System.Text.Encoding]::Default)
  try {
    Invoke-Vmrun -gu $GuestUser -gp $GuestPassword copyFileFromHostToGuest $Vmx $hostScript $guestScript
    [void]$script:GuestTempCmdScripts.Add($guestScript)
  } finally {
    Remove-Item -LiteralPath $hostScript -Force -ErrorAction SilentlyContinue
  }
  return $guestScript
}

function Invoke-GuestCmdScript {
  param([string]$Name, [string[]]$Lines, [switch]$AllowFail)
  $guestScript = Copy-GuestCmdToTemp -Name $Name -Lines $Lines
  $guestExitCode = 0
  try {
    & $Vmrun -T ws -gu $GuestUser -gp $GuestPassword runProgramInGuest $Vmx "C:\Windows\System32\cmd.exe" "/d /c $guestScript"
    $guestExitCode = $LASTEXITCODE
  } finally {
    Remove-GuestFileIfExists $guestScript
    [void]$script:GuestTempCmdScripts.Remove($guestScript)
  }
  if ($guestExitCode -ne 0 -and -not $AllowFail) {
    throw "guest batch failed ($guestExitCode): $guestScript"
  }
}

function Start-InteractiveGuestCmdScript {
  param([string]$Name, [string[]]$Lines)
  $safeTask = Get-SafeName "$TaskName`_$Name"
  $payload = Copy-GuestCmdToTemp -Name "$Name`_payload" -Lines $Lines
  $createLog = "C:\Windows\Temp\$safeTask.create.log"
  $runLog = "C:\Windows\Temp\$safeTask.run.log"
  $taskArg = Quote-BatchArg -Value $safeTask -Label "TaskName"
  $userArg = Quote-BatchArg -Value $RunAsGuestUser -Label "InteractiveGuestUser"
  $passwordArg = Quote-BatchArg -Value $GuestPassword -Label "GuestPassword"
  $payloadArg = Quote-BatchArg -Value $payload -Label "scheduled task payload"
  $createLogArg = Quote-BatchArg -Value $createLog -Label "scheduled task create log"
  $runLogArg = Quote-BatchArg -Value $runLog -Label "scheduled task run log"
  $taskArtifact = [pscustomobject]@{
    TaskName = $safeTask
    Payload = $payload
    CreateLog = $createLog
    RunLog = $runLog
  }
  try {
    Invoke-GuestCmdScript -Name "$Name`_controller" -Lines @(
      "@echo off",
      "setlocal DisableDelayedExpansion",
      "schtasks /Delete /F /TN $taskArg >nul 2>&1",
      "schtasks /Create /F /TN $taskArg /SC ONCE /ST 23:59 /RL HIGHEST /IT /RU $userArg /RP $passwordArg /TR $payloadArg > $createLogArg 2>&1",
      "if errorlevel 1 exit /b %errorlevel%",
      "schtasks /Run /TN $taskArg > $runLogArg 2>&1",
      "exit /b %errorlevel%"
    )
  } catch {
    $copiedDiagnostics = @()
    if (Copy-GuestFileToHostIfExists -GuestPath $createLog -HostPath (Join-Path $hostTmp "$safeTask.create.log")) {
      $copiedDiagnostics += "$safeTask.create.log"
    }
    if (Copy-GuestFileToHostIfExists -GuestPath $runLog -HostPath (Join-Path $hostTmp "$safeTask.run.log")) {
      $copiedDiagnostics += "$safeTask.run.log"
    }
    $diagnosticHint = ""
    if ($copiedDiagnostics.Count -gt 0) {
      $diagnosticHint = " Copied scheduler diagnostics to host tmp: $($copiedDiagnostics -join ', ')"
    }
    Remove-GuestFileIfExists $payload
    [void]$script:GuestTempCmdScripts.Remove($payload)
    throw "Failed to start interactive scheduled task '$safeTask'. $($_.Exception.Message)$diagnosticHint"
  }
  [void]$script:InteractiveGuestTasks.Add($taskArtifact)
  return $taskArtifact
}

function Remove-InteractiveGuestTask {
  param([object]$Task)
  if ($null -eq $Task) {
    return
  }
  try {
    $taskName = [string]$Task.TaskName
    if (-not [string]::IsNullOrWhiteSpace($taskName)) {
      $taskArg = Quote-BatchArg -Value $taskName -Label "scheduled task cleanup name"
      Invoke-GuestCmdScript -Name "delete_$taskName" -AllowFail -Lines @(
        "@echo off",
        "setlocal DisableDelayedExpansion",
        "schtasks /Delete /F /TN $taskArg >nul 2>&1",
        "exit /b 0"
      )
    }
  } catch {
    Write-Warning "Could not delete interactive guest task '$($Task.TaskName)': $($_.Exception.Message)"
  }
  foreach ($path in @($Task.Payload, $Task.CreateLog, $Task.RunLog)) {
    Remove-GuestFileIfExists ([string]$path)
    [void]$script:GuestTempCmdScripts.Remove([string]$path)
  }
  [void]$script:InteractiveGuestTasks.Remove($Task)
}

function Remove-TrackedGuestArtifacts {
  foreach ($task in @($script:InteractiveGuestTasks)) {
    Remove-InteractiveGuestTask -Task $task
  }
  foreach ($guestScript in @($script:GuestTempCmdScripts)) {
    Remove-GuestFileIfExists $guestScript
    [void]$script:GuestTempCmdScripts.Remove($guestScript)
  }
}

function Copy-GuestFileToHostIfExists {
  param([string]$GuestPath, [string]$HostPath)
  try {
    if (Test-GuestFile $GuestPath) {
      Invoke-Vmrun -gu $GuestUser -gp $GuestPassword copyFileFromGuestToHost $Vmx $GuestPath $HostPath
      return $true
    }
  } catch {
    Write-Warning "Could not copy guest diagnostic file '$GuestPath': $($_.Exception.Message)"
  }
  return $false
}

function Copy-GuestTmpDiagnostics {
  param([string]$GuestTmpPath, [string]$HostTmpPath)
  New-Item -ItemType Directory -Force -Path $HostTmpPath | Out-Null
  try {
    $entries = Invoke-VmrunOutput -gu $GuestUser -gp $GuestPassword listDirectoryInGuest $Vmx $GuestTmpPath
  } catch {
    Write-Warning "Could not list guest tmp diagnostics '$GuestTmpPath': $($_.Exception.Message)"
    return
  }

  foreach ($entry in $entries) {
    $name = [string]$entry
    if ([string]::IsNullOrWhiteSpace($name) -or $name -match '^Directory list:' -or $name -in @('.', '..')) {
      continue
    }
    $fileName = [System.IO.Path]::GetFileName($name)
    if ([string]::IsNullOrWhiteSpace($fileName) -or $fileName -ne $name) {
      continue
    }
    if ($fileName -notmatch '\.(png|json|log|txt|cmd)$') {
      continue
    }
    try {
      Invoke-Vmrun -gu $GuestUser -gp $GuestPassword copyFileFromGuestToHost $Vmx (Join-Path $GuestTmpPath $fileName) (Join-Path $HostTmpPath $fileName)
    } catch {
      Write-Warning "Could not copy guest tmp diagnostic '$fileName': $($_.Exception.Message)"
    }
  }
}

function Assert-InteractiveDesktopReady {
  param([string]$HostSnapshotPath)
  $processes = Invoke-VmrunOutput -gu $GuestUser -gp $GuestPassword listProcessesInGuest $Vmx
  [System.IO.File]::WriteAllLines($HostSnapshotPath, $processes, [System.Text.Encoding]::UTF8)
  $text = $processes -join "`n"
  $hasExplorer = $text -match '(?i)\bexplorer(\.exe)?\b'
  $hasShellHost = $text -match '(?i)\bsihost(\.exe)?\b'
  $hasLogonUi = $text -match '(?i)\blogonui(\.exe)?\b'

  if ($hasLogonUi -and -not $hasExplorer) {
    throw "Guest appears to be at the lock/logon screen (LogonUI.exe present, Explorer.exe absent). Log in to the VM desktop before running GUI evidence."
  }
  if (-not ($hasExplorer -and $hasShellHost)) {
    throw "Guest interactive desktop is not ready. Expected Explorer.exe and sihost.exe in listProcessesInGuest output. Snapshot: $HostSnapshotPath"
  }
  if ([string]::IsNullOrWhiteSpace($InteractiveGuestUser) -and $GuestUser -match '(?i)^(\.\\)?administrator$') {
    Write-Warning "Interactive scheduled tasks may require the fully qualified logged-in user, for example WIN-HOST\Administrator. Pass -InteractiveGuestUser if /IT task launch fails."
  }
}

New-Item -ItemType Directory -Force -Path $HostEvidenceDir | Out-Null
$hostTmp = Join-Path $HostEvidenceDir "tmp"
Remove-HostTmp -EvidenceDir $HostEvidenceDir -TmpDir $hostTmp
New-Item -ItemType Directory -Force -Path $hostTmp | Out-Null
$hostTmpPlanJson = Join-Path $hostTmp "plan.json"
[System.IO.File]::WriteAllText($hostTmpPlanJson, $hostPlanText, [System.Text.UTF8Encoding]::new($false))

if (-not $SkipInteractiveDesktopCheck) {
  Assert-InteractiveDesktopReady -HostSnapshotPath (Join-Path $hostTmp "guest_processes_before_gui.txt")
}

$guestScripts = Join-Path $GuestWorkDir "scripts"
$guestEvidence = Join-Path $GuestWorkDir "evidence"
$guestTmp = Join-Path $guestEvidence "tmp"

# Keep vmrun command lines simple: stage every guest-side command as a .cmd
# file, copy it into the guest, then run only "cmd.exe /d /c <script path>".
$guestWorkDirArg = Quote-BatchArg -Value $GuestWorkDir -Label "GuestWorkDir"
$guestScriptsArg = Quote-BatchArg -Value $guestScripts -Label "guest scripts directory"
$guestEvidenceArg = Quote-BatchArg -Value $guestEvidence -Label "guest evidence directory"
$guestTmpArg = Quote-BatchArg -Value $guestTmp -Label "guest tmp directory"
$guestEvidenceRowsArg = Quote-BatchArg -Value (Join-Path $guestEvidence "row*.png") -Label "guest stale row evidence glob"
Invoke-GuestCmdScript -Name "mkdirs" -Lines @(
  "@echo off",
  "setlocal DisableDelayedExpansion",
  "mkdir $guestWorkDirArg 2>nul",
  "mkdir $guestScriptsArg 2>nul",
  "mkdir $guestEvidenceArg 2>nul",
  "if exist $guestTmpArg rmdir /s /q $guestTmpArg",
  "if exist $guestTmpArg exit /b 20",
  "mkdir $guestTmpArg 2>nul",
  "del /q $guestEvidenceRowsArg 2>nul",
  "exit /b 0"
)

foreach ($name in @("guest_preflight.py", "guest_gui_runner.py", "requirements-guest.txt", "guest_setup_pyautogui.ps1")) {
  $hostScriptPath = Join-Path $HostSkillDir "scripts\$name"
  if (-not (Test-Path -LiteralPath $hostScriptPath -PathType Leaf)) {
    throw "Required local guest script is missing: $hostScriptPath"
  }
  Invoke-Vmrun -gu $GuestUser -gp $GuestPassword copyFileFromHostToGuest $Vmx $hostScriptPath (Join-Path $guestScripts $name)
}

$guestPreflight = Join-Path $guestScripts "guest_preflight.py"
$guestPythonArg = Quote-BatchArg -Value $GuestPython -Label "GuestPython"
$guestPreflightArg = Quote-BatchArg -Value $guestPreflight -Label "guest preflight script"
$preflightJson = Join-Path $guestTmp "preflight.json"
$preflightTmpJson = "$preflightJson.tmp"
$preflightJsonArg = Quote-BatchArg -Value $preflightJson -Label "guest preflight JSON"
$preflightTmpJsonArg = Quote-BatchArg -Value $preflightTmpJson -Label "guest preflight tmp JSON"
$preflightTask = Start-InteractiveGuestCmdScript -Name "preflight" -Lines @(
  "@echo off",
  "setlocal DisableDelayedExpansion",
  "$guestPythonArg $guestPreflightArg > $preflightTmpJsonArg 2>&1",
  "set PRECHECK_EXIT=%errorlevel%",
  "move /Y $preflightTmpJsonArg $preflightJsonArg >nul 2>&1",
  "exit /b %PRECHECK_EXIT%"
)
try {
  if (-not (Wait-GuestFile -Path $preflightJson -Timeout 120)) {
    $copiedDiagnostics = @()
    if (Copy-GuestFileToHostIfExists -GuestPath $preflightTask.CreateLog -HostPath (Join-Path $hostTmp "$($preflightTask.TaskName).create.log")) {
      $copiedDiagnostics += "$($preflightTask.TaskName).create.log"
    }
    if (Copy-GuestFileToHostIfExists -GuestPath $preflightTask.RunLog -HostPath (Join-Path $hostTmp "$($preflightTask.TaskName).run.log")) {
      $copiedDiagnostics += "$($preflightTask.TaskName).run.log"
    }
    if (Copy-GuestFileToHostIfExists -GuestPath $preflightTmpJson -HostPath (Join-Path $hostTmp "preflight.json.tmp")) {
      $copiedDiagnostics += "preflight.json.tmp"
    }
    $diagnosticHint = ""
    if ($copiedDiagnostics.Count -gt 0) {
      $diagnosticHint = " Copied diagnostics to host tmp: $($copiedDiagnostics -join ', ')"
    }
    throw "Timed out waiting for interactive guest preflight: $preflightJson.$diagnosticHint"
  }
  $hostPreflightJson = Join-Path $hostTmp "preflight.json"
  Invoke-Vmrun -gu $GuestUser -gp $GuestPassword copyFileFromGuestToHost $Vmx $preflightJson $hostPreflightJson
  $preflightText = Get-Content -LiteralPath $hostPreflightJson -Raw
  try {
    $preflight = $preflightText | ConvertFrom-Json
  } catch {
    throw "Interactive guest preflight did not produce JSON. Python/PyAutoGUI may be missing or failed to start. Raw output: $preflightText"
  }
  if (-not $preflight.pyautogui -or -not $preflight.screenshot -or -not $preflight.image_usable) {
    throw "Interactive guest preflight failed or did not capture the real desktop. Details: $preflightText"
  }
  if ([string]::IsNullOrWhiteSpace([string]$preflight.sessionname)) {
    Write-Warning "Interactive guest preflight captured a usable screenshot, but SESSIONNAME was empty. Continuing because desktop capture is functional."
  }
} finally {
  Remove-InteractiveGuestTask -Task $preflightTask
}

$guestPlanJson = Join-Path $guestTmp "plan.json"
Invoke-Vmrun -gu $GuestUser -gp $GuestPassword copyFileFromHostToGuest $Vmx $hostTmpPlanJson $guestPlanJson

$runner = Join-Path $guestScripts "guest_gui_runner.py"
$resultJson = Join-Path $guestTmp "runner_result.json"
$runnerStdout = Join-Path $guestTmp "runner_stdout.txt"
$guestRunnerLog = Join-Path $guestTmp "guest_gui_runner.log"
$runnerArg = Quote-BatchArg -Value $runner -Label "guest runner script"
$guestPlanJsonArg = Quote-BatchArg -Value $guestPlanJson -Label "guest plan JSON"
$resultJsonArg = Quote-BatchArg -Value $resultJson -Label "guest result JSON"
$runnerStdoutArg = Quote-BatchArg -Value $runnerStdout -Label "guest runner stdout"
$runnerTask = Start-InteractiveGuestCmdScript -Name "runner" -Lines @(
  "@echo off",
  "setlocal DisableDelayedExpansion",
  "$guestPythonArg $runnerArg --plan $guestPlanJsonArg --out-dir $guestEvidenceArg --result-json $resultJsonArg --keep-tmp > $runnerStdoutArg 2>&1",
  "exit /b %errorlevel%"
)

try {
  if (-not (Wait-GuestFile -Path $resultJson -Timeout $TimeoutSeconds)) {
    $copiedDiagnostics = @()
    if (Copy-GuestFileToHostIfExists -GuestPath $runnerStdout -HostPath (Join-Path $hostTmp "runner_stdout.txt")) {
      $copiedDiagnostics += "runner_stdout.txt"
    }
    if (Copy-GuestFileToHostIfExists -GuestPath $guestRunnerLog -HostPath (Join-Path $hostTmp "guest_gui_runner.log")) {
      $copiedDiagnostics += "guest_gui_runner.log"
    }
    $diagnosticHint = ""
    if ($copiedDiagnostics.Count -gt 0) {
      $diagnosticHint = " Copied diagnostics to host tmp: $($copiedDiagnostics -join ', ')"
    }
    throw "Timed out waiting for guest GUI task result: $resultJson.$diagnosticHint"
  }

  $hostResultJson = Join-Path $hostTmp "runner_result.json"
  Invoke-Vmrun -gu $GuestUser -gp $GuestPassword copyFileFromGuestToHost $Vmx $resultJson $hostResultJson
  $result = Get-Content -LiteralPath $hostResultJson -Raw | ConvertFrom-Json
} finally {
  Remove-InteractiveGuestTask -Task $runnerTask
}

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
Remove-TrackedGuestArtifacts

$failedRows = @($result.results | Where-Object { $_.status -ne "captured" } | ForEach-Object { "row$($_.row): $($_.error)" })

if ($KeepTmp -or $failedRows.Count -gt 0) {
  Copy-GuestTmpDiagnostics -GuestTmpPath $guestTmp -HostTmpPath $hostTmp
}

if (-not $KeepTmp -and $failedRows.Count -eq 0) {
  Invoke-GuestCmdScript -Name "cleanup_tmp" -AllowFail -Lines @(
    "@echo off",
    "setlocal DisableDelayedExpansion",
    "rmdir /s /q $guestTmpArg 2>nul",
    "exit /b 0"
  )
  Remove-HostTmp -EvidenceDir $HostEvidenceDir -TmpDir $hostTmp
}

if ($failedRows.Count -gt 0) {
  throw "Guest GUI task completed with failed rows: $($failedRows -join '; ')"
}

Write-Host "Copied final row screenshots to $HostEvidenceDir"
