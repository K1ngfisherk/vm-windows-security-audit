param(
  [string]$Python = "",
  [string]$Requirements = ".\requirements-guest.txt",
  [string]$Wheelhouse = "",
  [string]$PythonInstaller = "",
  [string]$PythonDownloadUrl = "",
  [switch]$NoWinget
)

$ErrorActionPreference = "Stop"

function Resolve-RequirementsPath {
  param([string]$PathValue)
  if (Test-Path -LiteralPath $PathValue) {
    return (Resolve-Path -LiteralPath $PathValue).Path
  }
  $fromScript = Join-Path $PSScriptRoot $PathValue
  if (Test-Path -LiteralPath $fromScript) {
    return (Resolve-Path -LiteralPath $fromScript).Path
  }
  throw "requirements file not found: $PathValue"
}

function Test-PythonExe {
  param([string]$Exe)
  try {
    $output = & $Exe -c "import sys; print(sys.executable)" 2>$null
    if ($LASTEXITCODE -eq 0 -and $output) {
      return $true
    }
  } catch {
    return $false
  }
  return $false
}

function Resolve-PythonExe {
  if ($Python) {
    if (Test-PythonExe $Python) {
      return $Python
    }
    $cmd = Get-Command $Python -ErrorAction SilentlyContinue
    if ($cmd -and (Test-PythonExe $cmd.Source)) {
      return $cmd.Source
    }
  }

  foreach ($name in @("python.exe", "python3.exe")) {
    $cmd = Get-Command $name -ErrorAction SilentlyContinue
    if ($cmd -and (Test-PythonExe $cmd.Source)) {
      return $cmd.Source
    }
  }

  $commonRoots = @(
    "$env:LOCALAPPDATA\Programs\Python",
    "$env:ProgramFiles\Python*",
    "${env:ProgramFiles(x86)}\Python*",
    "C:\Python*"
  )
  foreach ($root in $commonRoots) {
    $candidates = Get-ChildItem -Path $root -Filter python.exe -Recurse -ErrorAction SilentlyContinue |
      Sort-Object FullName -Descending
    foreach ($candidate in $candidates) {
      if (Test-PythonExe $candidate.FullName) {
        return $candidate.FullName
      }
    }
  }

  $py = Get-Command "py.exe" -ErrorAction SilentlyContinue
  if ($py) {
    try {
      $resolved = & $py.Source -3 -c "import sys; print(sys.executable)" 2>$null
      if ($LASTEXITCODE -eq 0 -and $resolved -and (Test-PythonExe $resolved.Trim())) {
        return $resolved.Trim()
      }
    } catch {
      # Continue to installer paths below.
    }
  }

  return $null
}

function Install-Python {
  if ($PythonInstaller -and (Test-Path -LiteralPath $PythonInstaller)) {
    Write-Host "Installing Python from local installer: $PythonInstaller"
    Start-Process -FilePath $PythonInstaller -ArgumentList "/quiet InstallAllUsers=1 PrependPath=1 Include_pip=1 Include_launcher=1" -Wait
    return
  }

  if (-not $NoWinget) {
    $winget = Get-Command "winget.exe" -ErrorAction SilentlyContinue
    if ($winget) {
      Write-Host "Installing Python with winget"
      & $winget.Source install --id Python.Python.3.12 -e --silent --accept-package-agreements --accept-source-agreements
      if ($LASTEXITCODE -eq 0) {
        return
      }
      Write-Host "winget Python install failed; trying next method if available"
    }
  }

  if ($PythonDownloadUrl) {
    $target = Join-Path $env:TEMP "python-installer.exe"
    Write-Host "Downloading Python installer: $PythonDownloadUrl"
    Invoke-WebRequest -Uri $PythonDownloadUrl -OutFile $target
    Start-Process -FilePath $target -ArgumentList "/quiet InstallAllUsers=1 PrependPath=1 Include_pip=1 Include_launcher=1" -Wait
    return
  }

  throw "Python was not found and no automatic install method succeeded. Provide -PythonInstaller or -PythonDownloadUrl, or allow winget."
}

$requirementsPath = Resolve-RequirementsPath $Requirements
$pythonExe = Resolve-PythonExe

if (-not $pythonExe) {
  Install-Python
  $pythonExe = Resolve-PythonExe
}

if (-not $pythonExe) {
  throw "Python installation completed but python.exe could not be resolved."
}

Write-Host "Using Python: $pythonExe"
& $pythonExe -c "import sys; print(sys.version)"

try {
  & $pythonExe -m pip --version | Out-Host
} catch {
  & $pythonExe -m ensurepip --upgrade
}

if ($Wheelhouse -and (Test-Path -LiteralPath $Wheelhouse)) {
  Write-Host "Installing guest dependencies from offline wheelhouse: $Wheelhouse"
  & $pythonExe -m pip install --no-index --find-links $Wheelhouse -r $requirementsPath
} else {
  & $pythonExe -m pip install --upgrade pip
  & $pythonExe -m pip install -r $requirementsPath
}

& $pythonExe -c "import pyautogui, pygetwindow, pyperclip, PIL; print('guest GUI automation environment ready')"
