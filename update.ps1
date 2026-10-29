# Reporter - update to a newer version.
#
# Replaces the code and leaves everything personal alone: settings, the API key, the
# accounts folder, contacts, metric history, past output, and the virtual environment
# all survive untouched. That separation is the whole point -- an update that wipes a
# month's output or a configured key is an update nobody runs twice.
#
#     powershell -ExecutionPolicy Bypass -File .\update.ps1 -Source \\server\share\Reporter-Web
#     powershell -ExecutionPolicy Bypass -File .\update.ps1 -Source C:\Downloads\Reporter-Web.zip
#
# Add -WhatIf to see exactly what would change without changing anything.

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory = $true)]
    [string]$Source,
    [switch]$SkipBackup
)

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $root

function Write-Step($text) { Write-Host "`n$text" -ForegroundColor Cyan }
function Write-Ok($text)   { Write-Host "  $text" -ForegroundColor Green }
function Write-Warn($text) { Write-Host "  $text" -ForegroundColor Yellow }

# Everything below is yours, not the app's. Nothing in an update may overwrite it.
$KeepFiles = @(
    "user_settings.json",   # your paths, sender details, engine choice
    ".env",                 # the AI key
    "python-path.txt",      # this machine's interpreter
    "*.local.json"
)
$KeepFolders = @(
    ".venv",                # rebuilt only when requirements change
    "venv",
    "Accounts",             # your source reports
    "data",                 # contacts and metric history
    "logs",
    "Insights Output"
)

Write-Host "Reporter update" -ForegroundColor White
Write-Host "Installed at: $root"

# --- 1. Resolve the source ----------------------------------------------------------
Write-Step "Reading the new version..."
$temp = $null
if ($Source -like "*.zip") {
    if (-not (Test-Path $Source)) { Write-Host "  No file at $Source" -ForegroundColor Red; exit 1 }
    $temp = Join-Path $env:TEMP ("reporter-update-" + [guid]::NewGuid().ToString("N").Substring(0, 8))
    Expand-Archive -Path $Source -DestinationPath $temp -Force
    # The zip contains a single Reporter-Web folder; step into it if so.
    $inner = Get-ChildItem $temp -Directory
    $newVersion = if ($inner.Count -eq 1 -and -not (Test-Path (Join-Path $temp "reporter_web.py"))) { $inner[0].FullName } else { $temp }
} else {
    if (-not (Test-Path $Source)) { Write-Host "  No folder at $Source" -ForegroundColor Red; exit 1 }
    $newVersion = $Source
}

if (-not (Test-Path (Join-Path $newVersion "reporter_web.py"))) {
    Write-Host "  That does not look like a Reporter package (no reporter_web.py)." -ForegroundColor Red
    exit 1
}

function Get-Version($folder) {
    $file = Join-Path $folder "VERSION"
    if (Test-Path $file) { return (Get-Content $file -First 1).Trim() }
    return "unknown"
}
$from = Get-Version $root
$to = Get-Version $newVersion
Write-Ok "$from  ->  $to"
if ($from -eq $to -and $from -ne "unknown") {
    Write-Warn "Already on this version. Continuing anyway will simply re-copy the same files."
}

# --- 2. Back up what cannot be re-derived -------------------------------------------
if (-not $SkipBackup) {
    Write-Step "Backing up your settings..."
    $backup = Join-Path $root ("backup-" + (Get-Date -Format "yyyyMMdd-HHmmss"))
    if ($PSCmdlet.ShouldProcess($backup, "Create settings backup")) {
        New-Item -ItemType Directory -Path $backup -Force | Out-Null
        foreach ($name in @("user_settings.json", ".env", "python-path.txt")) {
            $path = Join-Path $root $name
            if (Test-Path $path) { Copy-Item $path $backup -Force }
        }
        $dataFolder = Join-Path $root "data"
        if (Test-Path $dataFolder) { Copy-Item $dataFolder (Join-Path $backup "data") -Recurse -Force }
        Write-Ok "Saved to $backup"
        Write-Warn "Delete it once the update is confirmed good -- it holds your API key."
    }
}

# --- 3. Copy the new code over ------------------------------------------------------
Write-Step "Updating the application files..."

$excludeFiles = $KeepFiles -join " "
$excludeDirs = ($KeepFolders | ForEach-Object { "`"$(Join-Path $root $_)`"" }) -join " "

# /E every folder, /XF and /XD to protect what is yours, /NP for readable output.
# Deliberately NOT /MIR: mirroring would delete files an operator put here themselves.
$robocopyArgs = @(
    "`"$newVersion`"", "`"$root`"", "/E", "/NP", "/NFL", "/NDL", "/NJH", "/R:2", "/W:2"
)
if ($excludeFiles) { $robocopyArgs += "/XF"; $robocopyArgs += $KeepFiles }
# /XD full paths are matched against the SOURCE tree, so they must point into $newVersion.
if ($excludeDirs)  { $robocopyArgs += "/XD"; $robocopyArgs += ($KeepFolders | ForEach-Object { Join-Path $newVersion $_ }) }

if ($PSCmdlet.ShouldProcess($root, "Copy new application files")) {
    $output = & robocopy @robocopyArgs
    # Robocopy exit codes below 8 are success; 8 and above are real failures.
    if ($LASTEXITCODE -ge 8) {
        Write-Host "  Copy failed (robocopy code $LASTEXITCODE)." -ForegroundColor Red
        Write-Warn "Is Reporter still running? Close the browser tab and stop the server, then retry."
        exit 1
    }
    Write-Ok "Application files updated"
} else {
    Write-Host "  Would copy from $newVersion" -ForegroundColor DarkGray
    & robocopy @robocopyArgs "/L"
}

# Stale bytecode from the previous version confuses nobody more than it confuses Python.
Get-ChildItem $root -Filter "__pycache__" -Recurse -Directory -ErrorAction SilentlyContinue |
    Where-Object { $_.FullName -notlike "*\.venv\*" } |
    ForEach-Object { if ($PSCmdlet.ShouldProcess($_.FullName, "Remove stale bytecode")) { Remove-Item $_.FullName -Recurse -Force } }

# --- 4. Refresh dependencies if they changed ----------------------------------------
Write-Step "Checking dependencies..."
$venvPython = Join-Path $root ".venv\Scripts\python.exe"
if (Test-Path $venvPython) {
    if ($PSCmdlet.ShouldProcess("requirements-web.txt", "Install any new dependencies")) {
        & $venvPython -m pip install -r (Join-Path $root "requirements-web.txt") --quiet --upgrade
        if ($LASTEXITCODE -eq 0) { Write-Ok "Dependencies up to date" }
        else { Write-Warn "Dependency install reported a problem. Run setup.ps1 -Force if Reporter will not start." }
    }
} else {
    Write-Warn "No .venv here. Run setup.ps1 first."
}

if ($temp -and (Test-Path $temp)) { Remove-Item $temp -Recurse -Force -ErrorAction SilentlyContinue }

# --- 5. Confirm it still starts -----------------------------------------------------
if ((Test-Path $venvPython) -and $PSCmdlet.ShouldProcess("Reporter", "Verify the update")) {
    Write-Step "Verifying..."
    & $venvPython -c "import sys; sys.path.insert(0, r'$root'); import webapp.api; print('  Reporter imports cleanly')"
    if ($LASTEXITCODE -ne 0) {
        Write-Host "  Reporter did not import after the update." -ForegroundColor Red
        Write-Warn "Your settings are safe in the backup folder. Restore the previous package and report this."
        exit 1
    }
}

Write-Host "`nUpdate complete. Your settings, contacts, history and output were left alone." -ForegroundColor Green
