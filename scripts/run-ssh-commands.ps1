param(
  [Parameter(Mandatory=$true)][string]$HostName,
  [Parameter(Mandatory=$true)][string]$User,
  [string]$Password = "",
  [string]$KeyFile = "",
  [int]$Port = 22,
  [Parameter(Mandatory=$true)][string]$CommandsJson,
  [string]$OutputRoot = (Join-Path (Get-Location) 'ssh-command-output'),
  [string]$TaskLabel = "",
  [switch]$ValidateOnly
)

$ErrorActionPreference = "Stop"
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path

function Assert-CommandExists([string]$Name) {
  if (-not (Get-Command $Name -ErrorAction SilentlyContinue)) {
    throw "Required executable not found: $Name"
  }
}

function ConvertTo-SafeFilePart([string]$Text) {
  if ([string]::IsNullOrWhiteSpace($Text)) { return 'command' }
  return (($Text -replace '[\\/:*?"<>|\s]+','_') -replace '[^A-Za-z0-9_.-]','_').Trim('_')
}

function Load-CommandList([string]$Path) {
  if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { throw "CommandsJson not found: $Path" }
  $items = Get-Content -LiteralPath $Path -Raw -Encoding UTF8 | ConvertFrom-Json
  if ($null -eq $items) { throw 'CommandsJson is empty or invalid.' }
  if ($items -isnot [System.Array]) { $items = @($items) }
  return @($items)
}

function Quote-ShSingle([string]$Value) {
  return "'" + ($Value -replace "'", "'\''") + "'"
}

function Test-TcpPort([string]$TargetHost, [int]$TargetPort, [int]$TimeoutMs = 3000) {
  $client = [System.Net.Sockets.TcpClient]::new()
  try {
    $async = $client.BeginConnect($TargetHost, $TargetPort, $null, $null)
    if (-not $async.AsyncWaitHandle.WaitOne($TimeoutMs)) {
      return $false
    }
    $client.EndConnect($async)
    return $true
  } catch {
    return $false
  } finally {
    $client.Close()
  }
}

function Invoke-ParamikoSshCommands {
  $helper = Join-Path $scriptDir 'run-ssh-commands.py'
  if (-not (Test-Path -LiteralPath $helper -PathType Leaf)) {
    throw "Required SSH password helper is missing: $helper"
  }
  $python = Get-Command python -ErrorAction SilentlyContinue
  if (-not $python) {
    throw "Password SSH requires Python with paramiko, or use key/agent SSH."
  }
  $args = @(
    $helper,
    '--host-name', $HostName,
    '--user', $User,
    '--password', $Password,
    '--port', [string]$Port,
    '--commands-json', $CommandsJson,
    '--output-root', $OutputRoot,
    '--task-label', $TaskLabel
  )
  if (-not [string]::IsNullOrWhiteSpace($KeyFile)) { $args += @('--key-file', $KeyFile) }
  if ($ValidateOnly) { $args += @('--validate-only') }
  $oldErrorActionPreference = $ErrorActionPreference
  $ErrorActionPreference = "Continue"
  try {
    $output = @(& $python.Source @args 2>&1 | ForEach-Object { [string]$_ })
    $code = $LASTEXITCODE
  } finally {
    $ErrorActionPreference = $oldErrorActionPreference
  }
  if ($code -ne 0) {
    throw "SSH password command path failed ($code): $($output -join [Environment]::NewLine)"
  }
  foreach ($line in $output) {
    if (-not [string]::IsNullOrWhiteSpace($line)) { Write-Host $line }
  }
  exit 0
}

function Invoke-OpenSshCommand([string]$RemoteCommand) {
  $target = "$User@$HostName"
  $args = @(
    '-p', [string]$Port,
    '-o', 'BatchMode=yes',
    '-o', 'StrictHostKeyChecking=no',
    '-o', 'ConnectTimeout=8'
  )
  if (-not [string]::IsNullOrWhiteSpace($KeyFile)) {
    $args += @('-i', $KeyFile)
  }
  $args += @($target, $RemoteCommand)
  $output = @(& ssh.exe @args 2>&1 | ForEach-Object { [string]$_ })
  return [pscustomobject]@{ ExitCode = $LASTEXITCODE; Output = $output }
}

if (-not [string]::IsNullOrWhiteSpace($Password) -and [string]::IsNullOrWhiteSpace($KeyFile)) {
  Invoke-ParamikoSshCommands
}

Assert-CommandExists ssh.exe
$commands = Load-CommandList $CommandsJson

if ($ValidateOnly) {
  Write-Host "Validation OK. Commands loaded: $($commands.Count)"
  Write-Host "OpenSSH: $((Get-Command ssh.exe).Source)"
  Write-Host "SSH target: $User@$HostName`:$Port"
  exit 0
}

if (-not (Test-TcpPort -TargetHost $HostName -TargetPort $Port)) {
  throw "SSH port is not reachable: $HostName`:$Port"
}

$timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$label = ConvertTo-SafeFilePart $TaskLabel
if ([string]::IsNullOrWhiteSpace($label)) {
  $label = ConvertTo-SafeFilePart $HostName
}
$outputDir = Join-Path $OutputRoot (('{0}_{1}' -f $label, $timestamp))
New-Item -ItemType Directory -Force -Path $outputDir | Out-Null
$manifest = [System.Collections.Generic.List[object]]::new()

for ($i = 0; $i -lt $commands.Count; $i++) {
  $cmd = $commands[$i]
  $rawId = [string]$cmd.id
  if ([string]::IsNullOrWhiteSpace($rawId)) {
    $rawId = 'cmd{0:D2}' -f ($i + 1)
  }
  $id = ConvertTo-SafeFilePart $rawId
  $name = if ($cmd.name) { [string]$cmd.name } else { $id }
  $command = [string]$cmd.command
  if ([string]::IsNullOrWhiteSpace($command)) {
    throw "Command item '$id' has no command text."
  }
  $fileName = '{0:D2}_{1}.txt' -f ($i + 1), $id
  $outputPath = Join-Path $outputDir $fileName
  $wrapped = "printf '%s\n' '$ command'; " + $command
  $remote = "sh -lc " + (Quote-ShSingle $wrapped)
  $result = Invoke-OpenSshCommand -RemoteCommand $remote
  if ($result.ExitCode -eq 255) {
    throw "SSH command transport failed for '$id': $($result.Output -join [Environment]::NewLine)"
  }
  [System.IO.File]::WriteAllLines($outputPath, $result.Output, [System.Text.UTF8Encoding]::new($false))
  $manifest.Add([pscustomobject]@{
    id = $id
    name = $name
    command = $command
    transport = "ssh"
    output = $outputPath
    exitCode = $result.ExitCode
    screenshot = $null
  })
}

$manifestPath = Join-Path $outputDir 'manifest.json'
$manifest | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $manifestPath -Encoding UTF8
Write-Host "Output: $outputDir"
Write-Host "Manifest: $manifestPath"
Write-Host "Commands: $($commands.Count)"
