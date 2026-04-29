param(
  [Parameter(Mandatory=$true)][string]$HostName,
  [Parameter(Mandatory=$true)][string]$User,
  [string]$Password,
  [Parameter(Mandatory=$true)][string]$CommandsJson,
  [string]$OutputRoot = (Join-Path (Get-Location) 'ssh-evidence'),
  [int]$DefaultWaitSeconds = 2,
  [int]$LoginWaitSeconds = 3,
  [int]$TerminalCols = 180,
  [int]$TerminalRows = 60,
  [switch]$ValidateOnly
)

$ErrorActionPreference = 'Stop'

function Assert-CommandExists([string]$Name) {
  if (-not (Get-Command $Name -ErrorAction SilentlyContinue)) {
    throw "Required command not found: $Name"
  }
}

function ConvertTo-SafeFilePart([string]$Text) {
  if ([string]::IsNullOrWhiteSpace($Text)) { return 'command' }
  return (($Text -replace '[\\/:*?"<>|\s]+','_') -replace '[^A-Za-z0-9_.-]','_').Trim('_')
}

function Load-CommandList([string]$Path) {
  if (-not (Test-Path -LiteralPath $Path)) { throw "CommandsJson not found: $Path" }
  $items = Get-Content -LiteralPath $Path -Raw -Encoding UTF8 | ConvertFrom-Json
  if ($null -eq $items) { throw 'CommandsJson is empty or invalid.' }
  if ($items -isnot [System.Array]) { $items = @($items) }
  $index = 0
  foreach ($item in $items) {
    $index++
    if (-not $item.command) { throw "Command item $index is missing 'command'." }
    if (-not $item.id) { $item | Add-Member -NotePropertyName id -NotePropertyValue ('cmd{0:D2}' -f $index) }
    if (-not $item.name) { $item | Add-Member -NotePropertyName name -NotePropertyValue $item.id }
    if (-not $item.waitSeconds) { $item | Add-Member -NotePropertyName waitSeconds -NotePropertyValue $DefaultWaitSeconds }
  }
  return @($items)
}

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type @"
using System;
using System.Runtime.InteropServices;
public class SshEvidenceWin32 {
  [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr hWnd);
  [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
  [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);
  [DllImport("user32.dll")] public static extern bool PrintWindow(IntPtr hwnd, IntPtr hdcBlt, uint nFlags);
}
public struct RECT { public int Left; public int Top; public int Right; public int Bottom; }
"@

function Get-TerminalByTitle([string]$Title, [int]$TimeoutSeconds = 25) {
  $deadline = (Get-Date).AddSeconds($TimeoutSeconds)
  do {
    $proc = Get-Process -ErrorAction SilentlyContinue |
      Where-Object { $_.ProcessName -eq 'WindowsTerminal' -and $_.MainWindowTitle -eq $Title } |
      Select-Object -First 1
    if ($proc) { return $proc }
    Start-Sleep -Milliseconds 300
  } while ((Get-Date) -lt $deadline)
  throw "Windows Terminal window not found: $Title"
}

function Invoke-TerminalPaste([IntPtr]$Hwnd, [string]$Text) {
  [SshEvidenceWin32]::SetForegroundWindow($Hwnd) | Out-Null
  Start-Sleep -Milliseconds 250
  Set-Clipboard $Text
  $shell = New-Object -ComObject WScript.Shell
  $shell.SendKeys('^v')
  Start-Sleep -Milliseconds 200
  $shell.SendKeys('{ENTER}')
}

function Save-WindowScreenshot([IntPtr]$Hwnd, [string]$Path) {
  $rect = New-Object RECT
  [SshEvidenceWin32]::GetWindowRect($Hwnd, [ref]$rect) | Out-Null
  $width = $rect.Right - $rect.Left
  $height = $rect.Bottom - $rect.Top
  if ($width -le 0 -or $height -le 0) { throw "Bad terminal window size: ${width}x${height}" }
  $bitmap = New-Object System.Drawing.Bitmap($width, $height)
  $graphics = [System.Drawing.Graphics]::FromImage($bitmap)
  $hdc = $graphics.GetHdc()
  [SshEvidenceWin32]::PrintWindow($Hwnd, $hdc, 2) | Out-Null
  $graphics.ReleaseHdc($hdc)
  $graphics.Dispose()
  $bitmap.Save($Path, [System.Drawing.Imaging.ImageFormat]::Png)
  $bitmap.Dispose()
}

Assert-CommandExists wt.exe
Assert-CommandExists ssh.exe
$commands = Load-CommandList $CommandsJson

if ($ValidateOnly) {
  Write-Host "Validation OK. Commands loaded: $($commands.Count)"
  Write-Host "Windows Terminal: $((Get-Command wt.exe).Source)"
  Write-Host "OpenSSH: $((Get-Command ssh.exe).Source)"
  return
}

$timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$outputDir = Join-Path $OutputRoot (('{0}_{1}' -f (ConvertTo-SafeFilePart $HostName), $timestamp))
New-Item -ItemType Directory -Force -Path $outputDir | Out-Null

$title = 'ssh_evidence_' + (ConvertTo-SafeFilePart $HostName) + '_' + $timestamp
$sshTarget = "$User@$HostName"
Start-Process wt.exe -ArgumentList @('new-tab','--title',$title,'ssh','-tt','-o','StrictHostKeyChecking=no',$sshTarget)
$terminal = Get-TerminalByTitle $title 25
$hwnd = $terminal.MainWindowHandle
[SshEvidenceWin32]::ShowWindow($hwnd, 3) | Out-Null
[SshEvidenceWin32]::SetForegroundWindow($hwnd) | Out-Null
Start-Sleep -Seconds $LoginWaitSeconds

if ($Password) {
  Invoke-TerminalPaste $hwnd $Password
  Start-Sleep -Seconds 2
}

Invoke-TerminalPaste $hwnd "export LANG=C; stty cols $TerminalCols rows $TerminalRows; clear; echo CONNECTED_TO_$HostName; hostname; whoami"
Start-Sleep -Seconds 1

$manifest = [System.Collections.Generic.List[object]]::new()
$index = 0
foreach ($item in $commands) {
  $index++
  $id = ConvertTo-SafeFilePart ([string]$item.id)
  $name = [string]$item.name
  $wait = $DefaultWaitSeconds
  if (($item.PSObject.Properties.Name -contains 'waitSeconds') -and $null -ne $item.waitSeconds -and -not [string]::IsNullOrWhiteSpace([string]$item.waitSeconds)) {
    $wait = [int]$item.waitSeconds
  }
  if ($wait -lt 0) { $wait = $DefaultWaitSeconds }
  $fileName = ('{0:D2}_{1}.png' -f $index, $id)
  $screenshotPath = Join-Path $outputDir $fileName
  Invoke-TerminalPaste $hwnd ([string]$item.command)
  Start-Sleep -Seconds $wait
  Save-WindowScreenshot $hwnd $screenshotPath
  $manifest.Add([pscustomobject]@{
    index = $index
    id = $id
    name = $name
    command = [string]$item.command
    waitSeconds = $wait
    screenshot = $screenshotPath
    capturedAt = (Get-Date).ToString('s')
  })
}

$manifestPath = Join-Path $outputDir 'manifest.json'
$manifest | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $manifestPath -Encoding UTF8
Invoke-TerminalPaste $hwnd 'exit'

Write-Host "Evidence folder: $outputDir"
Write-Host "Manifest: $manifestPath"
Write-Host "Screenshots: $($commands.Count)"
