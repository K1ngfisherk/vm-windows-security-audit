param(
  [Parameter(Mandatory=$true)][string]$Vmrun,
  [Parameter(Mandatory=$true)][string]$Vmx,
  [Parameter(Mandatory=$true)][string]$GuestUser,
  [Parameter(Mandatory=$true)][string]$GuestPassword,
  [Parameter(Mandatory=$true)][string]$CommandsJson,
  [string]$OutputRoot = (Join-Path (Get-Location) 'guest-command-output'),
  [ValidateSet('auto','windows','linux')][string]$GuestOs = 'auto',
  [string]$TaskLabel = '',
  [switch]$ValidateOnly
)

$ErrorActionPreference = "Stop"

function Assert-CommandExists([string]$PathOrName) {
  if ($PathOrName -match '[\\/]') {
    if (-not (Test-Path -LiteralPath $PathOrName -PathType Leaf)) {
      throw "Command not found: $PathOrName"
    }
    return
  }
  if (-not (Get-Command $PathOrName -ErrorAction SilentlyContinue)) {
    throw "Command not found: $PathOrName"
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

function Invoke-Vmrun {
  param([Parameter(ValueFromRemainingArguments=$true)][string[]]$Args)
  $output = @(& $Vmrun -T ws @Args 2>&1 | ForEach-Object { [string]$_ })
  $code = $LASTEXITCODE
  return [pscustomobject]@{ ExitCode = $code; Output = $output }
}

function Assert-VmrunOk {
  param([object]$Result, [string]$Action)
  if ($Result.ExitCode -ne 0) {
    throw "vmrun failed during ${Action} ($($Result.ExitCode)): $($Result.Output -join [Environment]::NewLine)"
  }
}

function Detect-GuestOs([string]$VmxPath, [string]$Requested) {
  if ($Requested -ne 'auto') { return $Requested }
  $text = Get-Content -LiteralPath $VmxPath -Raw -ErrorAction SilentlyContinue
  if ($text -match '(?i)guestos\s*=\s*".*windows') { return 'windows' }
  return 'linux'
}

function Write-Utf8NoBom([string]$Path, [string[]]$Lines, [string]$LineEnding = "`r`n") {
  [System.IO.File]::WriteAllText($Path, (($Lines -join $LineEnding) + $LineEnding), [System.Text.UTF8Encoding]::new($false))
}

function New-WindowsScript {
  param([string]$Path, [string]$OutPath, [string]$Command)
  Write-Utf8NoBom -Path $Path -Lines @(
    '@echo off',
    'setlocal EnableExtensions DisableDelayedExpansion',
    ('call :run > "{0}" 2>&1' -f $OutPath),
    'exit /b %errorlevel%',
    ':run',
    ('echo $ {0}' -f $Command),
    $Command,
    'exit /b %errorlevel%'
  )
}

function New-LinuxScript {
  param([string]$Path, [string]$OutPath, [string]$Command)
  Write-Utf8NoBom -Path $Path -LineEnding "`n" -Lines @(
    '#!/bin/sh',
    ('OUT={0}' -f (Quote-ShSingle $OutPath)),
    '{',
    "  printf '%s\n' '$ command'",
    ('  {0}' -f $Command),
    '} > "$OUT" 2>&1',
    'exit $?'
  )
}

function Quote-ShSingle([string]$Value) {
  return "'" + ($Value -replace "'", "'\''") + "'"
}

function Remove-GuestFileIfExists {
  param([string]$Path)
  if ([string]::IsNullOrWhiteSpace($Path)) { return }
  $result = Invoke-Vmrun -gu $GuestUser -gp $GuestPassword deleteFileInGuest $Vmx $Path
  if ($result.ExitCode -ne 0) {
    Write-Warning "Could not remove guest file '$Path': $($result.Output -join ' ')"
  }
}

Assert-CommandExists $Vmrun
$commands = Load-CommandList $CommandsJson
$guestOsResolved = Detect-GuestOs -VmxPath $Vmx -Requested $GuestOs

if ($ValidateOnly) {
  Write-Host "Validation OK. Commands loaded: $($commands.Count)"
  Write-Host "Guest OS mode: $guestOsResolved"
  exit 0
}

$timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$vmName = ConvertTo-SafeFilePart ([System.IO.Path]::GetFileNameWithoutExtension($Vmx))
$label = ConvertTo-SafeFilePart $TaskLabel
if ([string]::IsNullOrWhiteSpace($label)) { $label = $vmName }
$outputDir = Join-Path $OutputRoot ('{0}安全检查证据' -f $label)
New-Item -ItemType Directory -Force -Path $outputDir | Out-Null
$hostTmp = Join-Path $outputDir 'tmp'
New-Item -ItemType Directory -Force -Path $hostTmp | Out-Null

$guestTmp = if ($guestOsResolved -eq 'windows') { 'C:\Windows\Temp' } else { '/tmp' }
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
  $prefix = '{0:D2}_{1}' -f ($i + 1), $id
  $hostScript = Join-Path $hostTmp "$prefix.$(if ($guestOsResolved -eq 'windows') { 'cmd' } else { 'sh' })"
  $hostOutput = Join-Path $outputDir "$prefix.txt"
  $guestScript = if ($guestOsResolved -eq 'windows') { Join-Path $guestTmp "$prefix.cmd" } else { "$guestTmp/$prefix.sh" }
  $guestOutput = if ($guestOsResolved -eq 'windows') { Join-Path $guestTmp "$prefix.txt" } else { "$guestTmp/$prefix.txt" }

  if ($guestOsResolved -eq 'windows') {
    New-WindowsScript -Path $hostScript -OutPath $guestOutput -Command $command
  } else {
    New-LinuxScript -Path $hostScript -OutPath $guestOutput -Command $command
  }

  Assert-VmrunOk (Invoke-Vmrun -gu $GuestUser -gp $GuestPassword copyFileFromHostToGuest $Vmx $hostScript $guestScript) "copy command script $id"
  $program = if ($guestOsResolved -eq 'windows') { 'C:\Windows\System32\cmd.exe' } else { '/bin/sh' }
  $arguments = if ($guestOsResolved -eq 'windows') { "/d /c `"$guestScript`"" } else { $guestScript }
  $runResult = Invoke-Vmrun -gu $GuestUser -gp $GuestPassword runProgramInGuest $Vmx $program $arguments
  $copyResult = Invoke-Vmrun -gu $GuestUser -gp $GuestPassword copyFileFromGuestToHost $Vmx $guestOutput $hostOutput
  Remove-GuestFileIfExists $guestScript
  Remove-GuestFileIfExists $guestOutput
  if ($copyResult.ExitCode -ne 0 -or -not (Test-Path -LiteralPath $hostOutput -PathType Leaf)) {
    throw "Could not copy command output for '$id': $($copyResult.Output -join [Environment]::NewLine)"
  }

  $manifest.Add([pscustomobject]@{
    id = $id
    name = $name
    command = $command
    output = $hostOutput
    exitCode = $runResult.ExitCode
    vmrunOutput = $runResult.Output
    screenshot = $null
  })
}

$manifestPath = Join-Path $outputDir 'manifest.json'
$manifest | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $manifestPath -Encoding UTF8
Remove-Item -LiteralPath $hostTmp -Recurse -Force -ErrorAction SilentlyContinue
Write-Host "Output: $outputDir"
Write-Host "Manifest: $manifestPath"
Write-Host "Commands: $($commands.Count)"
