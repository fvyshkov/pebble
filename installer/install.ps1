#Requires -Version 5.1
<#
.SYNOPSIS
    Pebble -one-click installer for Windows.

.DESCRIPTION
    Downloads Pebble, installs Python if needed, creates desktop shortcut.

    Usage (run in PowerShell):
        irm https://raw.githubusercontent.com/fvyshkov/pebble/main/installer/install.ps1 | iex

    Or from CMD:
        powershell -ExecutionPolicy Bypass -Command "irm https://raw.githubusercontent.com/fvyshkov/pebble/main/installer/install.ps1 | iex"
#>

$ErrorActionPreference = "Stop"
$ProgressPreference = "SilentlyContinue"  # Speed up Invoke-WebRequest

# ── Configuration ──────────────────────────────────────────────────
# Change DOWNLOAD_URL to where you host pebble-release.zip
$DOWNLOAD_URL = $env:PEBBLE_DOWNLOAD_URL
if (-not $DOWNLOAD_URL) {
    $DOWNLOAD_URL = "https://github.com/fvyshkov/pebble/releases/latest/download/pebble-release.zip"
}
$INSTALL_DIR = "$env:LOCALAPPDATA\Pebble"
$PYTHON_VERSION = "3.12.7"
$PYTHON_URL = "https://www.python.org/ftp/python/$PYTHON_VERSION/python-$PYTHON_VERSION-amd64.exe"

# ── Helpers ────────────────────────────────────────────────────────

function Write-Step($msg) {
    Write-Host ""
    Write-Host "  [Pebble] $msg" -ForegroundColor Cyan
}

function Write-Ok($msg) {
    Write-Host "  [OK] $msg" -ForegroundColor Green
}

function Write-Err($msg) {
    Write-Host "  [ERROR] $msg" -ForegroundColor Red
}

# ── Check admin not required ───────────────────────────────────────

Write-Host ""
Write-Host "  ╔══════════════════════════════════════╗" -ForegroundColor Cyan
Write-Host "  ║       Pebble -Installation          ║" -ForegroundColor Cyan
Write-Host "  ╚══════════════════════════════════════╝" -ForegroundColor Cyan
Write-Host ""

# ── 1. Install Python if missing ──────────────────────────────────

function Find-Python {
    # Try python, then py launcher
    foreach ($cmd in @("python", "py")) {
        $p = Get-Command $cmd -ErrorAction SilentlyContinue
        if ($p) {
            try {
                $ver = & $p.Source --version 2>&1
                if ($ver -match "3\.1[0-9]") {
                    return $p.Source
                }
            } catch {}
        }
    }
    # Check common install locations
    $locations = @(
        "$env:LOCALAPPDATA\Programs\Python\Python312\python.exe",
        "$env:LOCALAPPDATA\Programs\Python\Python311\python.exe",
        "$env:LOCALAPPDATA\Programs\Python\Python310\python.exe",
        "$env:ProgramFiles\Python312\python.exe",
        "$env:ProgramFiles\Python311\python.exe"
    )
    foreach ($loc in $locations) {
        if (Test-Path $loc) { return $loc }
    }
    return $null
}

$python = Find-Python
if (-not $python) {
    Write-Step "Python not found. Installing Python $PYTHON_VERSION..."

    $installer = "$env:TEMP\python-installer.exe"

    # Try winget first (faster, no UAC prompt needed for per-user)
    $winget = Get-Command winget -ErrorAction SilentlyContinue
    if ($winget) {
        Write-Step "Installing via winget..."
        & winget install Python.Python.3.12 --accept-source-agreements --accept-package-agreements --silent 2>$null

        # Refresh PATH
        $env:PATH = [System.Environment]::GetEnvironmentVariable("PATH", "User") + ";" + [System.Environment]::GetEnvironmentVariable("PATH", "Machine")
        $python = Find-Python
    }

    if (-not $python) {
        Write-Step "Downloading Python installer..."
        Invoke-WebRequest -Uri $PYTHON_URL -OutFile $installer

        Write-Step "Running Python installer..."
        Start-Process -FilePath $installer -ArgumentList "/quiet", "InstallAllUsers=0", "PrependPath=1", "Include_launcher=1" -Wait
        Remove-Item $installer -ErrorAction SilentlyContinue

        # Refresh PATH
        $env:PATH = [System.Environment]::GetEnvironmentVariable("PATH", "User") + ";" + [System.Environment]::GetEnvironmentVariable("PATH", "Machine")
        $python = Find-Python
    }

    if (-not $python) {
        Write-Err "Python installation failed. Please install Python 3.10+ manually from https://python.org"
        Read-Host "Press Enter to exit"
        exit 1
    }
    Write-Ok "Python installed: $python"
} else {
    Write-Ok "Python found: $python"
}

# ── Kill old Pebble processes before updating ─────────────────────

Write-Step "Stopping any running Pebble..."

# Kill by port 8000
$existingConn = Get-NetTCPConnection -LocalPort 8000 -ErrorAction SilentlyContinue
if ($existingConn) {
    $existingConn | ForEach-Object { Stop-Process -Id $_.OwningProcess -Force -ErrorAction SilentlyContinue }
}

# Kill ALL processes that might lock files in the install dir
$ErrorActionPreference = "SilentlyContinue"
taskkill /F /IM python.exe 2>&1 | Out-Null
taskkill /F /IM pythonw.exe 2>&1 | Out-Null
taskkill /F /IM pip.exe 2>&1 | Out-Null
# Kill cmd processes that may have CWD inside Pebble dir (locks the folder)
Get-Process cmd -ErrorAction SilentlyContinue | Where-Object {
    $_.Id -ne $PID
} | Stop-Process -Force -ErrorAction SilentlyContinue
$ErrorActionPreference = "Stop"

Start-Sleep -Seconds 3

# ── 2. Download Pebble ────────────────────────────────────────────

if ($DOWNLOAD_URL -match "YOUR-SERVER" -or $DOWNLOAD_URL -match "^$") {
    # Local mode: look for zip in same directory as script, or use existing install
    $scriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { Get-Location }
    $localZip = Join-Path $scriptDir "pebble-release.zip"

    if (Test-Path $localZip) {
        Write-Step "Found local pebble-release.zip, extracting..."
        if (Test-Path $INSTALL_DIR) { Remove-Item $INSTALL_DIR -Recurse -Force }
        Expand-Archive -Path $localZip -DestinationPath $INSTALL_DIR -Force
    } elseif (Test-Path (Join-Path $scriptDir "backend\main.py")) {
        # We're inside the app directory already -install in-place
        Write-Step "App files found in current directory. Installing in-place..."
        $INSTALL_DIR = $scriptDir
    } else {
        Write-Err "DOWNLOAD_URL not configured and no local pebble-release.zip found."
        Write-Host "  Set `$env:PEBBLE_DOWNLOAD_URL or place pebble-release.zip next to this script."
        Read-Host "Press Enter to exit"
        exit 1
    }
} else {
    Write-Step "Downloading Pebble..."
    $zipPath = "$env:TEMP\pebble-release.zip"
    Invoke-WebRequest -Uri $DOWNLOAD_URL -OutFile $zipPath

    Write-Step "Extracting..."

    # Clean old installation completely
    if (Test-Path $INSTALL_DIR) {
        Write-Step "Removing old installation..."
        for ($attempt = 1; $attempt -le 5; $attempt++) {
            # Use cmd rmdir -more reliable than PowerShell Remove-Item for locked files
            & cmd /c "rmdir /s /q `"$INSTALL_DIR`"" 2>$null
            if (-not (Test-Path $INSTALL_DIR)) { break }

            Write-Host "    Retry $attempt... (files still locked)" -ForegroundColor Yellow
            & taskkill /F /IM python.exe 2>$null
            & taskkill /F /IM pythonw.exe 2>$null
            Start-Sleep -Seconds 3
        }
        if (Test-Path $INSTALL_DIR) {
            Write-Err "Cannot remove old installation. Close all Pebble/Python windows and try again."
            exit 1
        }
    }

    # Extract to temp, then move inner 'pebble' folder to INSTALL_DIR
    $tempExtract = "$env:TEMP\pebble-extract"
    if (Test-Path $tempExtract) { Remove-Item $tempExtract -Recurse -Force }
    Expand-Archive -Path $zipPath -DestinationPath $tempExtract -Force
    Remove-Item $zipPath -ErrorAction SilentlyContinue

    $inner = Join-Path $tempExtract "pebble"
    if (-not (Test-Path $inner)) { $inner = $tempExtract }

    Move-Item $inner $INSTALL_DIR
    Remove-Item $tempExtract -Recurse -Force -ErrorAction SilentlyContinue

    Write-Ok "Extracted to $INSTALL_DIR"
}

# ── 3. Create venv & install deps ─────────────────────────────────

$venvDir = Join-Path $INSTALL_DIR ".venv"
if (-not (Test-Path $venvDir)) {
    Write-Step "Creating virtual environment..."
    & $python -m venv $venvDir
}

$venvPython = Join-Path $venvDir "Scripts\python.exe"
$venvPip = Join-Path $venvDir "Scripts\pip.exe"
$reqsFile = Join-Path $INSTALL_DIR "requirements.txt"

if (Test-Path $reqsFile) {
    Write-Step "Installing dependencies..."
    & $venvPip install -q -r $reqsFile
    Write-Ok "Dependencies installed"
}

# Install bundled pebble_calc Rust wheel (mandatory — formula_engine imports it)
$wheelsDir = Join-Path $INSTALL_DIR "wheels"
if (Test-Path $wheelsDir) {
    $wheels = Get-ChildItem -Path $wheelsDir -Filter "pebble_calc-*.whl" -ErrorAction SilentlyContinue
    if ($wheels) {
        Write-Step "Installing pebble_calc engine..."
        foreach ($w in $wheels) {
            & $venvPip install -q --force-reinstall $w.FullName
        }
        Write-Ok "pebble_calc installed"
    } else {
        Write-Err "No pebble_calc wheel found in $wheelsDir — backend will not start."
    }
} else {
    Write-Err "wheels/ folder missing in install — pebble_calc engine not installed, backend will crash."
}

# ── 4. Create desktop shortcut ────────────────────────────────────

Write-Step "Creating desktop shortcut..."

$desktopPath = [Environment]::GetFolderPath("Desktop")
$shortcutPath = Join-Path $desktopPath "Pebble.lnk"

$launcherBat = Join-Path $INSTALL_DIR "Pebble.bat"

# Create shortcut via COM
$shell = New-Object -ComObject WScript.Shell
$shortcut = $shell.CreateShortcut($shortcutPath)
$shortcut.TargetPath = $launcherBat
$shortcut.WorkingDirectory = $INSTALL_DIR
$shortcut.Description = "Pebble -Financial Modeling"
$shortcut.WindowStyle = 7  # Minimized
$shortcut.Save()

Write-Ok "Desktop shortcut created"

# ── 5. Create Start Menu shortcut ─────────────────────────────────

$startMenuDir = Join-Path $env:APPDATA "Microsoft\Windows\Start Menu\Programs"
$startShortcut = Join-Path $startMenuDir "Pebble.lnk"
$shortcut2 = $shell.CreateShortcut($startShortcut)
$shortcut2.TargetPath = $launcherBat
$shortcut2.WorkingDirectory = $INSTALL_DIR
$shortcut2.Description = "Pebble -Financial Modeling"
$shortcut2.WindowStyle = 7
$shortcut2.Save()

Write-Ok "Start Menu shortcut created"

# ── Done ───────────────────────────────────────────────────────────

Write-Host ""
Write-Host "  ╔══════════════════════════════════════╗" -ForegroundColor Green
Write-Host "  ║     Pebble installed successfully!   ║" -ForegroundColor Green
Write-Host "  ║                                      ║" -ForegroundColor Green
Write-Host "  ║  Double-click 'Pebble' on Desktop    ║" -ForegroundColor Green
Write-Host "  ║  to start the application.           ║" -ForegroundColor Green
Write-Host "  ╚══════════════════════════════════════╝" -ForegroundColor Green
Write-Host ""
Write-Host "  Install location: $INSTALL_DIR" -ForegroundColor Gray
Write-Host ""

# ── 6. Launch ─────────────────────────────────────────────────────

Write-Step "Starting Pebble..."
# Start in a new minimized window that stays open
Start-Process cmd -ArgumentList "/k", "cd /d `"$INSTALL_DIR`" && `"$launcherBat`"" -WorkingDirectory $INSTALL_DIR -WindowStyle Minimized
