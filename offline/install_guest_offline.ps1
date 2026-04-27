param(
  [string]$PythonInstaller = "",
  [string]$Wheelhouse = "",
  [string]$Requirements = ""
)

$ErrorActionPreference = "Stop"

$offlineDir = $PSScriptRoot
$repoRoot = Split-Path -Parent $offlineDir
$setup = Join-Path $repoRoot "scripts\guest_setup_pyautogui.ps1"

if (-not (Test-Path -LiteralPath $setup)) {
  throw "Cannot find scripts\guest_setup_pyautogui.ps1. Copy the full repository directory into the guest, not only the offline folder."
}

if (-not $PythonInstaller) {
  $PythonInstaller = Join-Path $offlineDir "python-3.9.13-amd64.exe"
}
if (-not $Wheelhouse) {
  $Wheelhouse = Join-Path $offlineDir "wheelhouse"
}
if (-not $Requirements) {
  $Requirements = Join-Path $offlineDir "requirements-guest-py39.txt"
}

& $setup `
  -PythonInstaller $PythonInstaller `
  -Wheelhouse $Wheelhouse `
  -Requirements $Requirements `
  -NoWinget
