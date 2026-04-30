param(
  [Parameter(Mandatory=$true)][string]$Vmrun,
  [Parameter(Mandatory=$true)][string]$Vmx,
  [Parameter(Mandatory=$true)][string]$GuestUser,
  [Parameter(Mandatory=$true)][string]$GuestPassword,
  [string]$InteractiveGuestUser = "",
  [string]$InteractiveGuestPassword = "",
  [string]$GuestWorkDir = "",
  [Parameter(Mandatory=$true)][string]$HostSkillDir,
  [Parameter(Mandatory=$true)][string]$HostPlanJson,
  [Parameter(Mandatory=$true)][string]$HostEvidenceDir,
  [string]$GuestPython = "python.exe",
  [string]$GuestPythonEnvRoot = "",
  [string]$GuestPythonInstallDir = "",
  [string]$TaskName = "CodexWindowsVmCheck",
  [int]$PollSeconds = 5,
  [int]$TimeoutSeconds = 900,
  [switch]$SkipInteractiveDesktopCheck,
  [switch]$SkipGuestPythonSetup,
  [switch]$KeepGuestPythonEnv,
  [switch]$KeepTmp,
  [switch]$DeferHostTmpCleanup,
  [switch]$Screenshots
)

$ErrorActionPreference = "Stop"
$ProgressPreference = "SilentlyContinue"
[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new($false)
$OutputEncoding = [System.Text.UTF8Encoding]::new($false)

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
$hostPlanText = Get-Content -LiteralPath $hostPlanFullPath -Raw -Encoding UTF8
$RunAsGuestUser = $InteractiveGuestUser.Trim()
$RunAsGuestPassword = if ([string]::IsNullOrWhiteSpace($InteractiveGuestPassword)) { $GuestPassword } else { $InteractiveGuestPassword }
if ($Screenshots) {
  if ([string]::IsNullOrWhiteSpace($RunAsGuestUser)) {
    throw "Windows GUI screenshot collection requires -InteractiveGuestUser with the actual logged-in Windows account, for example '.\Administrator' or 'HOSTNAME\Administrator'. Ask the user for the Windows Server login account before running vmrun GUI evidence."
  }
  if ($RunAsGuestUser -match '(?i)^(guest|\.\\guest|[^\\]+\\guest)$') {
    throw "Refusing to run Windows GUI evidence as guest account '$RunAsGuestUser'. Ask the user for an administrator or the actual logged-in Windows Server account and pass it as -InteractiveGuestUser."
  }
  if ([string]::IsNullOrWhiteSpace($RunAsGuestPassword)) {
    throw "Windows GUI screenshot collection requires a password for -InteractiveGuestUser. Ask the user for the Windows Server account password before running vmrun GUI evidence."
  }
}
$script:GuestTempCmdScripts = [System.Collections.Generic.List[string]]::new()
$script:InteractiveGuestTasks = [System.Collections.Generic.List[object]]::new()
$script:GuestCommandTmp = ""

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

function Join-Codepoints {
  param([int[]]$Codepoints)
  return -join ($Codepoints | ForEach-Object { [char]$_ })
}

function Get-GuestAccountName {
  param([string]$Account)
  $text = ([string]$Account).Trim()
  if ([string]::IsNullOrWhiteSpace($text)) {
    return ""
  }
  if ($text.Contains("\")) {
    $text = $text.Substring($text.LastIndexOf("\") + 1)
  }
  if ($text.Contains("@")) {
    $text = $text.Substring(0, $text.IndexOf("@"))
  }
  if ($text -eq ".") {
    return ""
  }
  return ($text -replace '[\\/:*?"<>|]', '_').Trim()
}

function Resolve-GuestUserWorkDir {
  param([string]$RequestedWorkDir, [string]$Account)
  $profileName = Get-GuestAccountName -Account $Account
  if ([string]::IsNullOrWhiteSpace($profileName)) {
    throw "Could not infer guest profile name from account '$Account'. Pass -GuestWorkDir under C:\Users\<username> explicitly."
  }
  $profileRoot = "C:\Users\$profileName"
  $defaultLeaf = "CodexVmAudit"

  if ([string]::IsNullOrWhiteSpace($RequestedWorkDir)) {
    return (Join-Path $profileRoot $defaultLeaf)
  }
  if (-not [System.IO.Path]::IsPathRooted($RequestedWorkDir)) {
    return (Join-Path $profileRoot $RequestedWorkDir)
  }
  if ($RequestedWorkDir -match '(?i)^C:\\Users\\[^\\]+(\\|$)') {
    return $RequestedWorkDir
  }

  $leaf = Split-Path -Leaf $RequestedWorkDir
  if ([string]::IsNullOrWhiteSpace($leaf)) {
    $leaf = $defaultLeaf
  }
  Write-Warning "GuestWorkDir '$RequestedWorkDir' is outside the guest user profile. Using '$profileRoot\$leaf' instead."
  return (Join-Path $profileRoot $leaf)
}

$AdminInterviewFinding = Join-Codepoints @(0x6D89,0x53CA,0x8BBF,0x8C08,0x7BA1,0x7406,0x5458,0xFF0C,0x672A,0x8FDB,0x884C,0x64CD,0x4F5C)
$AdminInterviewResult = Join-Codepoints @(0x672A,0x68C0,0x67E5)

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
  if ([string]::IsNullOrWhiteSpace($script:GuestCommandTmp)) {
    throw "Guest command tmp directory was not initialized"
  }
  $safeName = Get-SafeName "$TaskName`_$Name.cmd"
  $hostScript = Join-Path $hostTmp $safeName
  $guestScript = Join-Path $script:GuestCommandTmp $safeName
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
  $guestScriptArg = Quote-BatchArg -Value $guestScript -Label "guest script path"
  $guestExitCode = 0
  try {
    & $Vmrun -T ws -gu $GuestUser -gp $GuestPassword runProgramInGuest $Vmx "C:\Windows\System32\cmd.exe" "/d /c $guestScriptArg"
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
  $createLog = Join-Path $script:GuestCommandTmp "$safeTask.create.log"
  $runLog = Join-Path $script:GuestCommandTmp "$safeTask.run.log"
  $taskArg = Quote-BatchArg -Value $safeTask -Label "TaskName"
  $userArg = Quote-BatchArg -Value $RunAsGuestUser -Label "InteractiveGuestUser"
  $passwordArg = Quote-BatchArg -Value $RunAsGuestPassword -Label "InteractiveGuestPassword"
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

function Write-DeferredGuestCleanupManifest {
  param(
    [string]$HostTmpPath,
    [string]$GuestWorkPath,
    [string]$GuestEvidencePath,
    [string]$GuestTmpPath,
    [string]$GuestScriptsPath,
    [string]$GuestOfflinePath,
    [string]$GuestCommandTmpPath,
    [string]$GuestPythonEnvPath,
    [string]$GuestPythonInstallPath,
    [string]$GuestPythonManifestPath
  )
  New-Item -ItemType Directory -Force -Path $HostTmpPath | Out-Null
  $manifest = [ordered]@{
    cleanup_deferred_at = (Get-Date).ToString("o")
    reason = "Host report output requested delayed tmp cleanup; guest runtime was left intact so follow-up collection can still write under evidence\tmp."
    guest_work_dir = $GuestWorkPath
    guest_evidence_dir = $GuestEvidencePath
    guest_tmp_dir = $GuestTmpPath
    guest_scripts_dir = $GuestScriptsPath
    guest_offline_dir = $GuestOfflinePath
    guest_command_tmp_dir = $GuestCommandTmpPath
    guest_python_env_root = $GuestPythonEnvPath
    guest_python_install_dir = $GuestPythonInstallPath
    guest_python_manifest = $GuestPythonManifestPath
  }
  $manifest | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath (Join-Path $HostTmpPath "deferred_guest_cleanup.json") -Encoding UTF8
}

function Copy-HostDirectoryToGuest {
  param([string]$HostDir, [string]$GuestDir, [string]$Name)
  if (-not (Test-Path -LiteralPath $HostDir -PathType Container)) {
    return $false
  }

  $hostRoot = (Resolve-Path -LiteralPath $HostDir).Path.TrimEnd('\')
  $guestDirArg = Quote-BatchArg -Value $GuestDir -Label "$Name guest directory"
  Invoke-GuestCmdScript -Name "mkdir_$Name" -Lines @(
    "@echo off",
    "setlocal DisableDelayedExpansion",
    "mkdir $guestDirArg 2>nul",
    "exit /b 0"
  )

  foreach ($dir in Get-ChildItem -LiteralPath $hostRoot -Directory -Recurse -Force) {
    $relative = $dir.FullName.Substring($hostRoot.Length).TrimStart('\')
    $guestSubdir = Join-Path $GuestDir $relative
    $guestSubdirArg = Quote-BatchArg -Value $guestSubdir -Label "$Name guest subdirectory"
    Invoke-GuestCmdScript -Name "mkdir_$Name`_$(Get-SafeName $relative)" -Lines @(
      "@echo off",
      "setlocal DisableDelayedExpansion",
      "mkdir $guestSubdirArg 2>nul",
      "exit /b 0"
    )
  }

  foreach ($file in Get-ChildItem -LiteralPath $hostRoot -File -Recurse -Force) {
    $relative = $file.FullName.Substring($hostRoot.Length).TrimStart('\')
    $guestFile = Join-Path $GuestDir $relative
    Invoke-Vmrun -gu $GuestUser -gp $GuestPassword copyFileFromHostToGuest $Vmx $file.FullName $guestFile
  }
  return $true
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
  if (-not $hasExplorer) {
    throw "Guest interactive desktop is not ready. Expected Explorer.exe in listProcessesInGuest output. Snapshot: $HostSnapshotPath"
  }
  if (-not $hasShellHost) {
    Write-Warning "sihost.exe was not present in listProcessesInGuest output. Continuing because Windows Server 2012-class desktops can be ready with Explorer.exe only. Snapshot: $HostSnapshotPath"
  }
  if ($RunAsGuestUser -match '(?i)^(\.\\)?administrator$') {
    Write-Warning "Interactive scheduled tasks may require the fully qualified logged-in user, for example WIN-HOST\Administrator."
  }
}

New-Item -ItemType Directory -Force -Path $HostEvidenceDir | Out-Null
$hostTmp = Join-Path $HostEvidenceDir "tmp"
Remove-HostTmp -EvidenceDir $HostEvidenceDir -TmpDir $hostTmp
New-Item -ItemType Directory -Force -Path $hostTmp | Out-Null
$hostTmpPlanJson = Join-Path $hostTmp "plan.json"
[System.IO.File]::WriteAllText($hostTmpPlanJson, $hostPlanText, [System.Text.UTF8Encoding]::new($false))

if (-not $Screenshots) {
  $plan = $hostPlanText | ConvertFrom-Json
  $results = @()
  foreach ($item in @($plan.items)) {
    if ($item.gui_action -eq "skip_admin_interview" -or $item.check_id -eq "admin_interview") {
      $results += [pscustomobject]@{
        row = $item.row
        status = "skipped"
        skip_reason = "administrator_interview"
        action = $item.gui_action
        lane = $item.lane
        finding = if ($item.finding) { $item.finding } else { $AdminInterviewFinding }
        result = if ($item.result) { $item.result } else { $AdminInterviewResult }
        evidence = @()
      }
    } else {
      $results += [pscustomobject]@{
        row = $item.row
        status = "not_run"
        action = $item.gui_action
        lane = $item.lane
        evidence = @()
        note = "screenshots_disabled"
      }
    }
  }
  $result = [pscustomobject]@{
    screenshots = $false
    results = $results
    log = $null
    tmp_dir = $hostTmp
  }
  $result | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath (Join-Path $hostTmp "runner_result.json") -Encoding UTF8
  if (-not $KeepTmp -and -not $DeferHostTmpCleanup) {
    Remove-HostTmp -EvidenceDir $HostEvidenceDir -TmpDir $hostTmp
  }
  $skippedCount = @($results | Where-Object { $_.status -eq "skipped" }).Count
  $notRunCount = @($results | Where-Object { $_.status -eq "not_run" }).Count
  Write-Host "Screenshot collection disabled by default. Rows not run: $notRunCount; skipped administrator-interview rows: $skippedCount. No guest GUI work was run."
  if ($DeferHostTmpCleanup) {
    Write-Host "Host tmp preserved for workbook output: $hostTmp"
  }
  exit 0
}

$guestWorkAccount = if (-not [string]::IsNullOrWhiteSpace($RunAsGuestUser)) { $RunAsGuestUser } else { $GuestUser }
$GuestWorkDir = Resolve-GuestUserWorkDir -RequestedWorkDir $GuestWorkDir -Account $guestWorkAccount
$cmdRunLeaf = "{0}_{1}" -f (Get-SafeName $TaskName), (Get-Date -Format "yyyyMMddHHmmss")
$script:GuestCommandTmp = Join-Path (Join-Path $GuestWorkDir "cmdtmp") $cmdRunLeaf
Write-Host "Guest runtime work directory: $GuestWorkDir"

if (-not $SkipInteractiveDesktopCheck) {
  Assert-InteractiveDesktopReady -HostSnapshotPath (Join-Path $hostTmp "guest_processes_before_gui.txt")
}

$guestScripts = Join-Path $GuestWorkDir "scripts"
$guestEvidence = Join-Path $GuestWorkDir "evidence"
$guestTmp = Join-Path $guestEvidence "tmp"
$guestOffline = Join-Path $GuestWorkDir "offline"
if ([string]::IsNullOrWhiteSpace($GuestPythonEnvRoot)) {
  $GuestPythonEnvRoot = Join-Path $GuestWorkDir "pyenv"
}
if ([string]::IsNullOrWhiteSpace($GuestPythonInstallDir)) {
  $GuestPythonInstallDir = Join-Path $GuestWorkDir "python39"
}
$guestPythonManifest = Join-Path $guestTmp "python_env.json"

# Keep vmrun command lines simple: stage every guest-side command as a .cmd
# file, copy it into the guest, then run only "cmd.exe /d /c <script path>".
$guestWorkDirArg = Quote-BatchArg -Value $GuestWorkDir -Label "GuestWorkDir"
$guestScriptsArg = Quote-BatchArg -Value $guestScripts -Label "guest scripts directory"
$guestEvidenceArg = Quote-BatchArg -Value $guestEvidence -Label "guest evidence directory"
$guestTmpArg = Quote-BatchArg -Value $guestTmp -Label "guest tmp directory"
$guestOfflineArg = Quote-BatchArg -Value $guestOffline -Label "guest offline directory"
$guestCommandTmpArg = Quote-BatchArg -Value $script:GuestCommandTmp -Label "guest command tmp directory"
$guestEvidenceRowsArg = Quote-BatchArg -Value (Join-Path $guestEvidence "row*.png") -Label "guest stale row evidence glob"
& $Vmrun -T ws -gu $GuestUser -gp $GuestPassword runProgramInGuest $Vmx "C:\Windows\System32\cmd.exe" "/d /c mkdir $guestCommandTmpArg 2>nul"
if ($LASTEXITCODE -ne 0) {
  throw "Could not create guest command tmp directory under the user profile: $($script:GuestCommandTmp) (vmrun exit $LASTEXITCODE)"
}
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

$hostOfflineDir = Join-Path $HostSkillDir "offline"
$offlineStaged = $false
if (-not $SkipGuestPythonSetup) {
  $offlineStaged = Copy-HostDirectoryToGuest -HostDir $hostOfflineDir -GuestDir $guestOffline -Name "offline"
  $guestSetup = Join-Path $guestScripts "guest_setup_pyautogui.ps1"
  $guestSetupLog = Join-Path $guestTmp "python_setup.log"
  $guestSetupArg = Quote-BatchArg -Value $guestSetup -Label "guest setup script"
  $guestSetupLogArg = Quote-BatchArg -Value $guestSetupLog -Label "guest setup log"
  $guestPythonInputArg = Quote-BatchArg -Value $GuestPython -Label "GuestPython"
  $guestPythonEnvRootArg = Quote-BatchArg -Value $GuestPythonEnvRoot -Label "GuestPythonEnvRoot"
  $guestPythonInstallDirArg = Quote-BatchArg -Value $GuestPythonInstallDir -Label "GuestPythonInstallDir"
  $guestPythonManifestArg = Quote-BatchArg -Value $guestPythonManifest -Label "guest python manifest"
  if ($offlineStaged) {
    $guestRequirements = Join-Path $guestOffline "requirements-guest-py39.txt"
    $guestWheelhouse = Join-Path $guestOffline "wheelhouse"
    $guestPythonInstaller = Join-Path $guestOffline "python-3.9.13-amd64.exe"
  } else {
    $guestRequirements = Join-Path $guestScripts "requirements-guest.txt"
    $guestWheelhouse = ""
    $guestPythonInstaller = ""
  }
  $guestRequirementsArg = Quote-BatchArg -Value $guestRequirements -Label "guest requirements"
  $setupCommand = "powershell -NoProfile -OutputFormat Text -ExecutionPolicy Bypass -File $guestSetupArg -Python $guestPythonInputArg -Requirements $guestRequirementsArg -EnvRoot $guestPythonEnvRootArg -InstallDir $guestPythonInstallDirArg -Manifest $guestPythonManifestArg"
  if ($offlineStaged) {
    $guestWheelhouseArg = Quote-BatchArg -Value $guestWheelhouse -Label "guest wheelhouse"
    $guestPythonInstallerArg = Quote-BatchArg -Value $guestPythonInstaller -Label "guest python installer"
    $setupCommand = "$setupCommand -Wheelhouse $guestWheelhouseArg -PythonInstaller $guestPythonInstallerArg -NoWinget"
  }
  try {
    Invoke-GuestCmdScript -Name "setup_python_env" -Lines @(
      "@echo off",
      "setlocal DisableDelayedExpansion",
      "chcp 65001 >nul",
      "$setupCommand > $guestSetupLogArg 2>&1",
      "exit /b %errorlevel%"
    )
  } catch {
    $copiedDiagnostics = @()
    if (Copy-GuestFileToHostIfExists -GuestPath $guestSetupLog -HostPath (Join-Path $hostTmp "python_setup.log")) {
      $copiedDiagnostics += "python_setup.log"
    }
    $diagnosticHint = ""
    if ($copiedDiagnostics.Count -gt 0) {
      $diagnosticHint = " Copied diagnostics to host tmp: $($copiedDiagnostics -join ', ')"
    }
    throw "Failed to prepare guest Python environment. $($_.Exception.Message)$diagnosticHint"
  }

  $hostPythonManifest = Join-Path $hostTmp "python_env.json"
  Invoke-Vmrun -gu $GuestUser -gp $GuestPassword copyFileFromGuestToHost $Vmx $guestPythonManifest $hostPythonManifest
  $pythonManifest = Get-Content -LiteralPath $hostPythonManifest -Raw -Encoding UTF8 | ConvertFrom-Json
  $GuestPython = [string]$pythonManifest.python_exe
  if ([string]::IsNullOrWhiteSpace($GuestPython)) {
    throw "Guest Python environment manifest did not contain python_exe: $hostPythonManifest"
  }
  Write-Host "Using managed guest Python: $GuestPython"
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
  "chcp 65001 >nul",
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
  $preflightText = Get-Content -LiteralPath $hostPreflightJson -Raw -Encoding UTF8
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
  "chcp 65001 >nul",
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
  $result = Get-Content -LiteralPath $hostResultJson -Raw -Encoding UTF8 | ConvertFrom-Json
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

$failedRows = @($result.results | Where-Object { $_.status -notin @("captured", "skipped") } | ForEach-Object { "row$($_.row): $($_.error)" })

if ($KeepTmp -or $DeferHostTmpCleanup -or $failedRows.Count -gt 0) {
  Copy-GuestTmpDiagnostics -GuestTmpPath $guestTmp -HostTmpPath $hostTmp
}

if (-not $KeepTmp -and $failedRows.Count -eq 0) {
  if ($DeferHostTmpCleanup) {
    Write-DeferredGuestCleanupManifest `
      -HostTmpPath $hostTmp `
      -GuestWorkPath $GuestWorkDir `
      -GuestEvidencePath $guestEvidence `
      -GuestTmpPath $guestTmp `
      -GuestScriptsPath $guestScripts `
      -GuestOfflinePath $guestOffline `
      -GuestCommandTmpPath $script:GuestCommandTmp `
      -GuestPythonEnvPath $GuestPythonEnvRoot `
      -GuestPythonInstallPath $GuestPythonInstallDir `
      -GuestPythonManifestPath $guestPythonManifest
    Write-Host "Guest runtime tmp preserved for deferred cleanup: $guestTmp"
    Write-Host "Host tmp preserved for workbook output: $hostTmp"
  } else {
    if (-not $SkipGuestPythonSetup -and -not $KeepGuestPythonEnv) {
      $guestSetup = Join-Path $guestScripts "guest_setup_pyautogui.ps1"
      $guestSetupArg = Quote-BatchArg -Value $guestSetup -Label "guest setup cleanup script"
      $guestPythonEnvRootArg = Quote-BatchArg -Value $GuestPythonEnvRoot -Label "GuestPythonEnvRoot"
      $guestPythonInstallDirArg = Quote-BatchArg -Value $GuestPythonInstallDir -Label "GuestPythonInstallDir"
      $guestPythonManifestArg = Quote-BatchArg -Value $guestPythonManifest -Label "guest python manifest"
      Invoke-GuestCmdScript -Name "cleanup_python_env" -AllowFail -Lines @(
        "@echo off",
        "setlocal DisableDelayedExpansion",
      "chcp 65001 >nul",
      "powershell -NoProfile -OutputFormat Text -ExecutionPolicy Bypass -File $guestSetupArg -Cleanup -EnvRoot $guestPythonEnvRootArg -InstallDir $guestPythonInstallDirArg -Manifest $guestPythonManifestArg",
        "exit /b 0"
      )
    }
    Invoke-GuestCmdScript -Name "cleanup_runtime" -AllowFail -Lines @(
      "@echo off",
      "setlocal DisableDelayedExpansion",
      "rmdir /s /q $guestTmpArg 2>nul",
      "rmdir /s /q $guestScriptsArg 2>nul",
      "rmdir /s /q $guestOfflineArg 2>nul",
      "rmdir /s /q $guestEvidenceArg 2>nul",
      "rmdir /s /q $guestCommandTmpArg 2>nul",
      "rmdir $guestWorkDirArg 2>nul",
      "exit /b 0"
    )
    Remove-HostTmp -EvidenceDir $HostEvidenceDir -TmpDir $hostTmp
  }
}

if ($failedRows.Count -gt 0) {
  throw "Guest GUI task completed with failed rows: $($failedRows -join '; ')"
}

Write-Host "Copied final row screenshots to $HostEvidenceDir"
