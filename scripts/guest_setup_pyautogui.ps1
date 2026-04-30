param(
  [string]$Python = "",
  [string]$Requirements = ".\requirements-guest.txt",
  [string]$Wheelhouse = "",
  [string]$PythonInstaller = "",
  [string]$PythonDownloadUrl = "",
  [string]$EnvRoot = "",
  [string]$InstallDir = "",
  [string]$Manifest = "",
  [switch]$NoWinget,
  [switch]$Cleanup
)

$ErrorActionPreference = "Stop"
$ProgressPreference = "SilentlyContinue"
[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new($false)
$OutputEncoding = [System.Text.UTF8Encoding]::new($false)

function Resolve-FullPath {
  param([string]$PathValue)
  if ([string]::IsNullOrWhiteSpace($PathValue)) {
    return ""
  }
  return [System.IO.Path]::GetFullPath($PathValue)
}

function Test-PathUnderRoot {
  param([string]$PathValue, [string]$RootValue)
  if ([string]::IsNullOrWhiteSpace($PathValue) -or [string]::IsNullOrWhiteSpace($RootValue)) {
    return $false
  }
  $pathFull = Resolve-FullPath $PathValue
  $rootFull = (Resolve-FullPath $RootValue).TrimEnd('\')
  return $pathFull.Equals($rootFull, [System.StringComparison]::OrdinalIgnoreCase) -or
    $pathFull.StartsWith("$rootFull\", [System.StringComparison]::OrdinalIgnoreCase)
}

function Remove-TreeSafely {
  param([string]$PathValue, [string]$AllowedRoot, [string]$Label)
  if ([string]::IsNullOrWhiteSpace($PathValue) -or -not (Test-Path -LiteralPath $PathValue)) {
    return
  }
  if (-not (Test-PathUnderRoot -PathValue $PathValue -RootValue $AllowedRoot)) {
    throw "Refusing to remove $Label outside managed guest work root: $PathValue"
  }
  Remove-Item -LiteralPath $PathValue -Recurse -Force
}

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
  if ([string]::IsNullOrWhiteSpace($Exe) -or -not (Test-Path -LiteralPath $Exe -PathType Leaf)) {
    $cmd = Get-Command $Exe -ErrorAction SilentlyContinue
    if (-not $cmd) {
      return $false
    }
    $Exe = $cmd.Source
  }
  try {
    $output = & $Exe -c "import sys; print(sys.executable)" 2>$null
    return ($LASTEXITCODE -eq 0 -and -not [string]::IsNullOrWhiteSpace([string]$output))
  } catch {
    return $false
  }
}

function Resolve-PythonExePath {
  param([string]$Exe)
  if ([string]::IsNullOrWhiteSpace($Exe)) {
    return ""
  }
  if (Test-Path -LiteralPath $Exe -PathType Leaf) {
    return (Resolve-Path -LiteralPath $Exe).Path
  }
  $cmd = Get-Command $Exe -ErrorAction SilentlyContinue
  if ($cmd) {
    return $cmd.Source
  }
  return ""
}

function Get-RegistryPythonCandidates {
  $roots = @(
    "HKLM:\SOFTWARE\Python\PythonCore",
    "HKCU:\SOFTWARE\Python\PythonCore",
    "HKLM:\SOFTWARE\WOW6432Node\Python\PythonCore",
    "HKCU:\SOFTWARE\WOW6432Node\Python\PythonCore"
  )
  foreach ($root in $roots) {
    if (-not (Test-Path -LiteralPath $root)) {
      continue
    }
    foreach ($versionKey in Get-ChildItem -LiteralPath $root -ErrorAction SilentlyContinue) {
      $installKey = Join-Path $versionKey.PSPath "InstallPath"
      try {
        $key = Get-Item -LiteralPath $installKey -ErrorAction Stop
        $props = Get-ItemProperty -LiteralPath $installKey -ErrorAction Stop
        $defaultInstallPath = [string]$key.GetValue("")
        foreach ($value in @($props.ExecutablePath, $defaultInstallPath)) {
          if ([string]::IsNullOrWhiteSpace([string]$value)) {
            continue
          }
          $candidate = [string]$value
          if ((Split-Path -Leaf $candidate) -ne "python.exe") {
            $candidate = Join-Path $candidate "python.exe"
          }
          $candidate
        }
      } catch {
        continue
      }
    }
  }
}

function Resolve-PythonExe {
  param([string]$PreferredInstallDir)

  $candidates = [System.Collections.Generic.List[string]]::new()
  foreach ($value in @($Python, (Join-Path $PreferredInstallDir "python.exe"))) {
    if (-not [string]::IsNullOrWhiteSpace($value)) {
      [void]$candidates.Add($value)
    }
  }

  foreach ($name in @("python.exe", "python3.exe")) {
    $cmd = Get-Command $name -ErrorAction SilentlyContinue
    if ($cmd) {
      [void]$candidates.Add($cmd.Source)
    }
  }

  foreach ($root in @(
    "$env:LOCALAPPDATA\Programs\Python",
    "$env:ProgramFiles\Python*",
    "${env:ProgramFiles(x86)}\Python*",
    "C:\Python*"
  )) {
    Get-ChildItem -Path $root -Filter python.exe -Recurse -ErrorAction SilentlyContinue |
      Sort-Object FullName -Descending |
      ForEach-Object { [void]$candidates.Add($_.FullName) }
  }

  foreach ($candidate in Get-RegistryPythonCandidates) {
    [void]$candidates.Add($candidate)
  }

  $py = Get-Command "py.exe" -ErrorAction SilentlyContinue
  if ($py) {
    try {
      $resolved = & $py.Source -3 -c "import sys; print(sys.executable)" 2>$null
      if ($LASTEXITCODE -eq 0 -and -not [string]::IsNullOrWhiteSpace([string]$resolved)) {
        [void]$candidates.Add(([string]$resolved).Trim())
      }
    } catch {
      # Continue with other candidates.
    }
  }

  foreach ($candidate in $candidates | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique) {
    $resolvedCandidate = Resolve-PythonExePath $candidate
    if ($resolvedCandidate -and (Test-PythonExe $resolvedCandidate)) {
      return $resolvedCandidate
    }
  }
  return $null
}

function Install-Python {
  param([string]$TargetDir)
  if ($PythonInstaller -and (Test-Path -LiteralPath $PythonInstaller)) {
    New-Item -ItemType Directory -Force -Path $TargetDir | Out-Null
    Write-Host "Installing Python from local installer: $PythonInstaller"
    $args = @(
      "/quiet",
      "InstallAllUsers=0",
      "TargetDir=$TargetDir",
      "PrependPath=0",
      "Include_pip=1",
      "Include_launcher=0",
      "AssociateFiles=0",
      "Shortcuts=0",
      "CompileAll=0"
    )
    $process = Start-Process -FilePath $PythonInstaller -ArgumentList $args -Wait -PassThru
    if ($process.ExitCode -ne 0) {
      Write-Warning "Python installer exited with code $($process.ExitCode). Will try to resolve python.exe from known locations before failing."
    }
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
    New-Item -ItemType Directory -Force -Path $TargetDir | Out-Null
    $target = Join-Path $TargetDir "python-installer.exe"
    Write-Host "Downloading Python installer: $PythonDownloadUrl"
    Invoke-WebRequest -Uri $PythonDownloadUrl -OutFile $target
    $process = Start-Process -FilePath $target -ArgumentList @(
      "/quiet",
      "InstallAllUsers=0",
      "TargetDir=$TargetDir",
      "PrependPath=0",
      "Include_pip=1",
      "Include_launcher=0",
      "AssociateFiles=0",
      "Shortcuts=0",
      "CompileAll=0"
    ) -Wait -PassThru
    if ($process.ExitCode -ne 0) {
      Write-Warning "Downloaded Python installer exited with code $($process.ExitCode)."
    }
    return
  }

  throw "Python was not found and no automatic install method is available. Provide -PythonInstaller or -PythonDownloadUrl, or allow winget."
}

function New-GuestPythonManifest {
  param(
    [string]$PathValue,
    [object]$Data
  )
  $parent = Split-Path -Parent $PathValue
  if (-not [string]::IsNullOrWhiteSpace($parent)) {
    New-Item -ItemType Directory -Force -Path $parent | Out-Null
  }
  $Data | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $PathValue -Encoding UTF8
}

function Get-DefaultUserWorkRoot {
  if (-not [string]::IsNullOrWhiteSpace($env:USERPROFILE)) {
    return (Join-Path $env:USERPROFILE "CodexVmAudit")
  }
  if (-not [string]::IsNullOrWhiteSpace($env:USERNAME)) {
    return (Join-Path (Join-Path "C:\Users" $env:USERNAME) "CodexVmAudit")
  }
  return (Join-Path $env:TEMP "CodexVmAudit")
}

$defaultUserWorkRoot = Get-DefaultUserWorkRoot
if ([string]::IsNullOrWhiteSpace($EnvRoot)) {
  $EnvRoot = Join-Path $defaultUserWorkRoot "pyenv"
}
if ([string]::IsNullOrWhiteSpace($InstallDir)) {
  $InstallDir = Join-Path $defaultUserWorkRoot "python39"
}
if ([string]::IsNullOrWhiteSpace($Manifest)) {
  $Manifest = Join-Path $EnvRoot "python_env_manifest.json"
}

$EnvRoot = Resolve-FullPath $EnvRoot
$InstallDir = Resolve-FullPath $InstallDir
$Manifest = Resolve-FullPath $Manifest
$workRoot = Resolve-FullPath (Split-Path -Parent $EnvRoot)
if ([string]::IsNullOrWhiteSpace($workRoot)) {
  throw "Could not resolve managed guest work root from EnvRoot: $EnvRoot"
}

if ($Cleanup) {
  if (-not (Test-Path -LiteralPath $Manifest -PathType Leaf)) {
    Write-Host "No guest Python environment manifest found; nothing to clean: $Manifest"
    exit 0
  }
  $manifestData = Get-Content -LiteralPath $Manifest -Raw -Encoding UTF8 | ConvertFrom-Json
  $allowedRoot = if ($manifestData.work_root) { [string]$manifestData.work_root } else { $workRoot }
  Remove-TreeSafely -PathValue ([string]$manifestData.managed_venv_dir) -AllowedRoot $allowedRoot -Label "managed venv"
  Remove-TreeSafely -PathValue ([string]$manifestData.managed_python_dir) -AllowedRoot $allowedRoot -Label "managed Python"
  Write-Host "Guest Python environment cleanup complete"
  exit 0
}

$requirementsPath = Resolve-RequirementsPath $Requirements
$wheelhousePath = ""
if (-not [string]::IsNullOrWhiteSpace($Wheelhouse) -and (Test-Path -LiteralPath $Wheelhouse)) {
  $wheelhousePath = (Resolve-Path -LiteralPath $Wheelhouse).Path
}
$forceOfflinePython = -not [string]::IsNullOrWhiteSpace($PythonInstaller) -and (Test-Path -LiteralPath $PythonInstaller -PathType Leaf)

New-Item -ItemType Directory -Force -Path $workRoot | Out-Null
$foundCustomerPython = $false
$basePython = $null
if ($forceOfflinePython) {
  Remove-TreeSafely -PathValue $InstallDir -AllowedRoot $workRoot -Label "stale managed Python"
  Install-Python -TargetDir $InstallDir
  $managedPython = Join-Path $InstallDir "python.exe"
  if (Test-PythonExe $managedPython) {
    $basePython = (Resolve-Path -LiteralPath $managedPython).Path
  } else {
    throw "Offline Python installer did not produce a usable managed python.exe: $managedPython. The offline route intentionally does not fall back to the customer's existing Python."
  }
} else {
  $basePython = Resolve-PythonExe -PreferredInstallDir $InstallDir
  if ($basePython -and -not (Test-PathUnderRoot -PathValue $basePython -RootValue $InstallDir)) {
    $foundCustomerPython = $true
  }
  if (-not $basePython) {
    Install-Python -TargetDir $InstallDir
    $basePython = Resolve-PythonExe -PreferredInstallDir $InstallDir
  }
}

if (-not $basePython) {
  throw "Python installation completed but python.exe could not be resolved. Installer may have collided with an existing Python registration; pass -Python with the absolute python.exe path or remove the stale installation."
}

$managedPythonDir = $null
if (Test-PathUnderRoot -PathValue $basePython -RootValue $InstallDir) {
  $managedPythonDir = $InstallDir
}

if (Test-Path -LiteralPath $EnvRoot) {
  Remove-TreeSafely -PathValue $EnvRoot -AllowedRoot $workRoot -Label "stale managed venv"
}

Write-Host "Using base Python: $basePython"
& $basePython -m venv $EnvRoot
if ($LASTEXITCODE -ne 0) {
  throw "Failed to create managed guest Python venv: $EnvRoot"
}

$venvPython = Join-Path $EnvRoot "Scripts\python.exe"
if (-not (Test-PythonExe $venvPython)) {
  throw "Managed guest Python venv did not produce a usable python.exe: $venvPython"
}

try {
  & $venvPython -m pip --version | Out-Host
} catch {
  & $venvPython -m ensurepip --upgrade
}

if ($wheelhousePath) {
  Write-Host "Installing guest dependencies from offline wheelhouse: $wheelhousePath"
  & $venvPython -m pip install --no-index --find-links $wheelhousePath -r $requirementsPath
} else {
  & $venvPython -m pip install --upgrade pip
  & $venvPython -m pip install -r $requirementsPath
}
if ($LASTEXITCODE -ne 0) {
  throw "Failed to install guest GUI automation dependencies"
}

& $venvPython -c "import pyautogui, pygetwindow, pyperclip, PIL; print('guest GUI automation environment ready')"
if ($LASTEXITCODE -ne 0) {
  throw "Guest GUI automation imports failed after dependency installation"
}

$manifestData = [pscustomobject]@{
  version = 1
  created_at = (Get-Date).ToString("o")
  python_exe = $venvPython
  base_python_exe = $basePython
  found_customer_python = [bool]$foundCustomerPython
  forced_offline_python = [bool]$forceOfflinePython
  managed_venv_dir = $EnvRoot
  managed_python_dir = $managedPythonDir
  work_root = $workRoot
  requirements = $requirementsPath
  wheelhouse = $wheelhousePath
}
New-GuestPythonManifest -PathValue $Manifest -Data $manifestData
Write-Host "Guest Python environment ready: $venvPython"
Write-Host "Guest Python environment manifest: $Manifest"
