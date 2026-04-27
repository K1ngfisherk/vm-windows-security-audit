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
  [string]$TaskName = "CodexWindowsVmCheck"
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

New-Item -ItemType Directory -Force -Path $HostEvidenceDir | Out-Null

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
$taskCommand = "`"$GuestPython`" `"$runner`" --plan `"$GuestWorkDir\plan.json`" --out-dir `"$guestEvidence`" --result-json `"$resultJson`""

Invoke-Guest "schtasks.exe" @("/Create", "/F", "/TN", $TaskName, "/SC", "ONCE", "/ST", "23:59", "/RL", "HIGHEST", "/TR", $taskCommand)
Invoke-Guest "schtasks.exe" @("/Run", "/TN", $TaskName)

Write-Host "Started guest GUI task $TaskName. Wait for completion, then copy evidence with vmrun copyFileFromGuestToHost."
