# compile_and_setup_win.ps1
# rdgr.szal.kkbbs
# Stop script on error
$ErrorActionPreference = "Stop"

function Draw-Progress {
    param(
        [int]$pct,
        [string]$label
    )
    $barLen = 30
    $filled = [math]::Floor($pct * $barLen / 100)
    $empty = $barLen - $filled
    $barFilled = "#" * $filled
    $barEmpty = "." * $empty
    Write-Host -NoNewline "`r[$barFilled$barEmpty] $pct% - $label"
}

Write-Host "==================================================" -ForegroundColor Green
Write-Host "    Starting Daftar Radiologi Setup (Windows)     " -ForegroundColor Green
Write-Host "==================================================" -ForegroundColor Green

# 1. Get project paths (relative to script location)
$SCRIPT_DIR = Split-Path -Parent $MyInvocation.MyCommand.Path
$PROJECT_DIR = Split-Path -Parent $SCRIPT_DIR
$DESKTOP_DIR = [System.IO.Path]::Combine($env:USERPROFILE, "Desktop")
$APP_TARGET_DIR = [System.IO.Path]::Combine($DESKTOP_DIR, "DaftarRadiologi_App")
$SHORTCUT_PATH = [System.IO.Path]::Combine($DESKTOP_DIR, "Daftar Radiologi.lnk")

Draw-Progress 10 "Initializing paths..."
Start-Sleep -Seconds 1

# 2. Activate Virtual Environment
Draw-Progress 20 "Checking virtual environment..."
Set-Location $PROJECT_DIR
if (Test-Path ".venv") {
    & .venv\Scripts\Activate.ps1
} else {
    Draw-Progress 25 "Virtual environment not found. Creating a new one..."
    python -m venv .venv
    & .venv\Scripts\Activate.ps1
    Draw-Progress 40 "Installing dependencies (Flask, openpyxl, pyinstaller)..."
    pip install --quiet Flask openpyxl pyinstaller
}

# 3. Compile project using PyInstaller
Draw-Progress 50 "Preparing compilation..."
Set-Location "$PROJECT_DIR\Daftar_Radiologi"
if (Test-Path "build") { Remove-Item -Recurse -Force "build" }
if (Test-Path "dist") { Remove-Item -Recurse -Force "dist" }

# Run pyinstaller and update progress dynamically
Draw-Progress 60 "Compiling codebase with PyInstaller (please wait)..."
$job = Start-Process -FilePath "pyinstaller" -ArgumentList "--clean", "-y", "DaftarRadiologi.spec" -NoNewWindow -PassThru -RedirectStandardOutput "pyinstaller.log" -RedirectStandardError "pyinstaller_err.log"
$pct = 60
while (-not $job.HasExited) {
    if ($pct -lt 90) { $pct += 1 }
    Draw-Progress $pct "Compiling codebase with PyInstaller (please wait)..."
    Start-Sleep -Milliseconds 500
}

# Cleanup pyinstaller logs
if (Test-Path "pyinstaller.log") { Remove-Item -Force "pyinstaller.log" }
if (Test-Path "pyinstaller_err.log") { Remove-Item -Force "pyinstaller_err.log" }

# Return to root folder
Set-Location $PROJECT_DIR

Draw-Progress 90 "Structuring destination folders..."
# 4. Set up Application Folder on Desktop
if (Test-Path $APP_TARGET_DIR) {
    Draw-Progress 90 "Previous installation found. Keeping 'Pendaftaran' database intact..."
    Get-ChildItem -Path $APP_TARGET_DIR -Exclude "Pendaftaran" | Remove-Item -Recurse -Force
} else {
    New-Item -ItemType Directory -Force -Path $APP_TARGET_DIR > $null
}

# Copy executable and assets
Copy-Item -Path "$PROJECT_DIR\Daftar_Radiologi\dist\DaftarRadiologi\*" -Destination $APP_TARGET_DIR -Recurse -Force

# Copy database templates (Pendaftaran) if not already present
$PENDAFTARAN_TARGET = Join-Path $APP_TARGET_DIR "Pendaftaran"
if (-not (Test-Path $PENDAFTARAN_TARGET)) {
    $PENDAFTARAN_SOURCE = "$PROJECT_DIR\Daftar_Radiologi\Pendaftaran"
    if (Test-Path $PENDAFTARAN_SOURCE) {
        Copy-Item -Path $PENDAFTARAN_SOURCE -Destination $APP_TARGET_DIR -Recurse -Force
    } else {
        New-Item -ItemType Directory -Force -Path "$APP_TARGET_DIR\Pendaftaran" > $null
    }
} else {
    Draw-Progress 92 "Database folder already exists. Skipping templates copy..."
    Start-Sleep -Milliseconds 500
}

# 5. Create Desktop Shortcut (.lnk) using COM object WScript.Shell
Draw-Progress 95 "Creating Desktop shortcut..."
if (Test-Path $SHORTCUT_PATH) {
    Remove-Item -Force $SHORTCUT_PATH
}

$WshShell = New-Object -ComObject WScript.Shell
$Shortcut = $WshShell.CreateShortcut($SHORTCUT_PATH)
$Shortcut.TargetPath = "$APP_TARGET_DIR\DaftarRadiologi.exe"
$Shortcut.WorkingDirectory = $APP_TARGET_DIR
$Shortcut.Description = "Daftar Radiologi"
$Shortcut.Save()

Draw-Progress 100 "Setup completed successfully!"
Write-Host ""
Write-Host "==================================================" -ForegroundColor Green
Write-Host "Setup Successful!" -ForegroundColor Green
Write-Host "App folder created: $APP_TARGET_DIR" -ForegroundColor Green
Write-Host "Desktop Shortcut created: $SHORTCUT_PATH" -ForegroundColor Green
Write-Host "==================================================" -ForegroundColor Green
