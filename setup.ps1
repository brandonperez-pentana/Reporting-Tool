# Reporter - first-time setup on a machine.
#
# Finds a Python, builds a virtual environment beside the app, installs what Reporter
# needs, and writes python-path.txt so "Start Reporter.bat" works on double-click.
#
# Run it from the Reporter folder:
#
#     powershell -ExecutionPolicy Bypass -File .\setup.ps1
#
# Safe to run again: an existing .venv is reused unless you pass -Force. It never
# touches user_settings.json, .env, or anything under data\ and Accounts\.

[CmdletBinding()]
param(
    [switch]$Force,
    [string]$PythonExe = ""
)

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $root

function Write-Step($text) { Write-Host "`n$text" -ForegroundColor Cyan }
function Write-Ok($text)   { Write-Host "  $text" -ForegroundColor Green }
function Write-Warn($text) { Write-Host "  $text" -ForegroundColor Yellow }

Write-Host "Reporter setup" -ForegroundColor White
Write-Host "Folder: $root"

# --- 1. Find an interpreter ---------------------------------------------------------
Write-Step "Looking for Python 3.10 or newer..."

function Test-Python($path) {
    if (-not $path) { return $false }
    try {
        $version = & $path -c "import sys; print('%d.%d' % sys.version_info[:2])" 2>$null
        if ($LASTEXITCODE -ne 0 -or -not $version) { return $false }
        $parts = $version.Trim().Split(".")
        return ([int]$parts[0] -eq 3 -and [int]$parts[1] -ge 10)
    } catch { return $false }
}

$candidates = New-Object System.Collections.Generic.List[string]
if ($PythonExe) { $candidates.Add($PythonExe) }

foreach ($name in @("py", "python", "python3")) {
    $found = Get-Command $name -ErrorAction SilentlyContinue
    if ($found) { $candidates.Add($found.Source) }
}

# Registry, which is where a per-user or all-users install actually records itself.
foreach ($hive in @("HKLM:\SOFTWARE\Python\PythonCore", "HKCU:\SOFTWARE\Python\PythonCore")) {
    if (Test-Path $hive) {
        Get-ChildItem $hive -ErrorAction SilentlyContinue | ForEach-Object {
            $installPath = (Get-ItemProperty "$($_.PSPath)\InstallPath" -ErrorAction SilentlyContinue)."(default)"
            if ($installPath) { $candidates.Add((Join-Path $installPath "python.exe")) }
        }
    }
}

# Common install locations, including the Store shim that Get-ChildItem cannot enumerate.
foreach ($guess in @(
    "$env:LOCALAPPDATA\Programs\Python\Python312\python.exe",
    "$env:LOCALAPPDATA\Programs\Python\Python311\python.exe",
    "$env:LOCALAPPDATA\Microsoft\WindowsApps\python.exe",
    "C:\Python312\python.exe",
    "C:\Python311\python.exe",
    "C:\Program Files\Python312\python.exe",
    "C:\Program Files\Python311\python.exe"
)) { $candidates.Add($guess) }

# An existing venv from another Reporter copy on this machine is a valid base too.
Get-ChildItem "$env:USERPROFILE" -Filter "python.exe" -Recurse -Depth 4 -ErrorAction SilentlyContinue |
    Where-Object { $_.FullName -like "*\.venv\Scripts\python.exe" -or $_.FullName -like "*\venv\Scripts\python.exe" } |
    ForEach-Object { $candidates.Add($_.FullName) }

$python = $null
foreach ($candidate in $candidates) {
    if (Test-Path $candidate -ErrorAction SilentlyContinue) {
        if (Test-Python $candidate) { $python = $candidate; break }
    }
}

if (-not $python) {
    Write-Host "`nNo suitable Python found on this machine." -ForegroundColor Red
    Write-Host @"

  Install Python 3.12 from https://www.python.org/downloads/ and tick
  "Add python.exe to PATH" on the first screen of the installer, then run
  this script again.

  If Python is installed somewhere unusual, point at it directly:

      .\setup.ps1 -PythonExe "D:\Tools\Python312\python.exe"

  On a managed build you may need IT to approve the install.
"@ -ForegroundColor Yellow
    exit 1
}

$pythonVersion = (& $python -c "import sys; print(sys.version.split()[0])").Trim()
Write-Ok "Using $python  (Python $pythonVersion)"

# --- 2. Build the virtual environment ----------------------------------------------
$venv = Join-Path $root ".venv"
$venvPython = Join-Path $venv "Scripts\python.exe"

if ((Test-Path $venvPython) -and -not $Force) {
    Write-Step "Virtual environment already present, reusing it."
    Write-Warn "Pass -Force to rebuild it from scratch."
} else {
    if (Test-Path $venv) {
        Write-Step "Removing the existing virtual environment..."
        Remove-Item $venv -Recurse -Force
    }
    Write-Step "Creating the virtual environment..."
    & $python -m venv $venv
    if ($LASTEXITCODE -ne 0) { Write-Host "  Could not create .venv" -ForegroundColor Red; exit 1 }
    Write-Ok "Created $venv"
}

# --- 3. Install dependencies --------------------------------------------------------
Write-Step "Installing dependencies (this takes a minute)..."
& $venvPython -m pip install --upgrade pip --quiet
& $venvPython -m pip install -r (Join-Path $root "requirements-web.txt") --quiet
if ($LASTEXITCODE -ne 0) {
    Write-Host "  Dependency install failed." -ForegroundColor Red
    Write-Warn "Behind a corporate proxy this usually needs an index URL or a trusted host."
    exit 1
}
Write-Ok "Dependencies installed"

# The Performance Intelligence report is drawn in a browser, so Chromium is fetched
# once per machine. It is a large download; the app still runs without it and simply
# cannot build that report.
Write-Step "Installing Chromium for the Performance Intelligence report..."
& $venvPython -m playwright install chromium
if ($LASTEXITCODE -eq 0) { Write-Ok "Chromium installed" }
else { Write-Warn "Chromium did not download - often a proxy rule on a managed build." 
       Write-Warn "Reporter will fall back to Microsoft Edge or Google Chrome if either is installed," 
       Write-Warn "so the Performance Intelligence report should still work. Settings shows which is in use." }

# --- 4. Point the launcher at it ----------------------------------------------------
Set-Content -Path (Join-Path $root "python-path.txt") -Value $venvPython -NoNewline -Encoding ASCII
Write-Ok "python-path.txt now points at this machine's interpreter"

# --- 5. Optional AI layer -----------------------------------------------------------
$envFile = Join-Path $root ".env"
$envExample = Join-Path $root ".env.example"
if (-not (Test-Path $envFile) -and (Test-Path $envExample)) {
    Write-Step "Optional: the AI insight engine"
    Write-Warn "No .env file yet, so AI assisted is unavailable and Reporter runs on the"
    Write-Warn "rules or funnel engine as normal. To enable it later:"
    Write-Warn "    copy .env.example .env      then add a key"
}

# --- 6. Prove it works --------------------------------------------------------------
Write-Step "Checking the install..."
& $venvPython -c "import pdfplumber, docx, docxtpl, pandas, openpyxl, jinja2; print('  All modules import cleanly')"
if ($LASTEXITCODE -ne 0) { Write-Host "  Something did not install correctly." -ForegroundColor Red; exit 1 }

Write-Host "`nSetup complete." -ForegroundColor Green
Write-Host @"

  Start Reporter by double-clicking "Start Reporter.bat", or run:

      .\.venv\Scripts\python.exe reporter_web.py

  The first screen will ask for your Accounts folder, an output folder and a
  contacts file. Nothing is configured for you: those paths are per-person.
"@
