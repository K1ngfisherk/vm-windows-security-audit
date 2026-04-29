param(
  [string]$HostName = "",
  [Parameter(Mandatory=$true)][string]$User,
  [string]$Password = "",
  [string]$KeyFile = "",
  [int]$Port = 22,
  [Parameter(Mandatory=$true)][string]$CommandsJson,
  [string]$OutputRoot = (Join-Path (Get-Location) 'linux-command-output'),
  [string]$TaskLabel = "",
  [string]$Vmrun = "",
  [string]$Vmx = "",
  [string]$GuestUser = "",
  [string]$GuestPassword = "",
  [switch]$AllowVmrunFallback,
  [switch]$ValidateOnly
)

$ErrorActionPreference = "Stop"
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$sshScript = Join-Path $scriptDir 'run-ssh-commands.ps1'
$vmrunScript = Join-Path $scriptDir 'run-guest-commands.ps1'

function Invoke-Script {
  param([string]$Path, [object[]]$ScriptArgs)
  $output = @(& powershell -NoProfile -ExecutionPolicy Bypass -File $Path @ScriptArgs 2>&1 | ForEach-Object { [string]$_ })
  $code = $LASTEXITCODE
  if ($code -ne 0) {
    throw "Script failed ($code): $Path`n$($output -join [Environment]::NewLine)"
  }
  foreach ($line in $output) {
    if (-not [string]::IsNullOrWhiteSpace($line)) {
      Write-Host $line
    }
  }
}

$sshArgs = @(
  '-HostName', $HostName,
  '-User', $User,
  '-Port', [string]$Port,
  '-CommandsJson', $CommandsJson,
  '-OutputRoot', $OutputRoot,
  '-TaskLabel', $TaskLabel
)
if (-not [string]::IsNullOrWhiteSpace($Password)) { $sshArgs += @('-Password', $Password) }
if (-not [string]::IsNullOrWhiteSpace($KeyFile)) { $sshArgs += @('-KeyFile', $KeyFile) }

$guestUserToUse = if ([string]::IsNullOrWhiteSpace($GuestUser)) { $User } else { $GuestUser }
$guestPasswordToUse = if ([string]::IsNullOrWhiteSpace($GuestPassword)) { $Password } else { $GuestPassword }
$vmrunArgs = @(
  '-Vmrun', $Vmrun,
  '-Vmx', $Vmx,
  '-GuestUser', $guestUserToUse,
  '-GuestPassword', $guestPasswordToUse,
  '-CommandsJson', $CommandsJson,
  '-OutputRoot', $OutputRoot,
  '-GuestOs', 'linux',
  '-TaskLabel', $TaskLabel
)

if ($ValidateOnly) {
  if (-not [string]::IsNullOrWhiteSpace($HostName)) {
    Invoke-Script -Path $sshScript -ScriptArgs ($sshArgs + @('-ValidateOnly'))
  }
  if ($AllowVmrunFallback -and -not [string]::IsNullOrWhiteSpace($Vmrun) -and -not [string]::IsNullOrWhiteSpace($Vmx)) {
    Invoke-Script -Path $vmrunScript -ScriptArgs ($vmrunArgs + @('-ValidateOnly'))
  }
  if ([string]::IsNullOrWhiteSpace($HostName) -and -not $AllowVmrunFallback) {
    throw "Linux SSH connection info is required first: HostName/IP, User, and Password or KeyFile. Use vmrun fallback only after the user says SSH is unavailable or cannot be provided."
  }
  exit 0
}

if ([string]::IsNullOrWhiteSpace($HostName)) {
  if (-not $AllowVmrunFallback) {
    throw "Linux SSH connection info is required first: HostName/IP, User, and Password or KeyFile. Ask the customer for SSH IP and account credentials before using vmrun."
  }
} else {
  try {
    Write-Host "Trying Linux SSH command path first: $User@$HostName`:$Port"
    Invoke-Script -Path $sshScript -ScriptArgs $sshArgs
    exit 0
  } catch {
    $firstLine = ([string]$_.Exception.Message -split "\r?\n")[0]
    if (-not $AllowVmrunFallback) {
      throw "SSH command path failed and vmrun fallback is disabled. Ask the customer for working SSH IP/account credentials, or rerun with -AllowVmrunFallback only after the user says SSH is unavailable or cannot be provided. $firstLine"
    }
    Write-Warning "SSH command path failed; user allowed vmrun fallback. $firstLine"
  }
}

if (-not $AllowVmrunFallback) {
  throw "vmrun fallback is disabled. Use it only after the user says SSH is unavailable or cannot be provided."
}

if (
  [string]::IsNullOrWhiteSpace($Vmrun) -or
  [string]::IsNullOrWhiteSpace($Vmx) -or
  [string]::IsNullOrWhiteSpace($guestUserToUse) -or
  [string]::IsNullOrWhiteSpace($guestPasswordToUse)
) {
  throw "SSH failed or was unavailable, and vmrun fallback inputs are incomplete."
}

Write-Host "Trying Linux vmrun direct command fallback: $Vmx"
Invoke-Script -Path $vmrunScript -ScriptArgs $vmrunArgs
