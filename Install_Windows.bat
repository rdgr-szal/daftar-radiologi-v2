@echo off
:: rdgr.szal.kkbbs
cd /d "%~dp0"
echo ==================================================
echo Starting installation of Daftar Radiologi (Windows)...
echo ==================================================

:: Check if Python is installed
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo Python not found.
    echo Downloading and installing Python automatically...
    powershell -Command "[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12; (New-Object System.Net.WebClient).DownloadFile('https://www.python.org/ftp/python/3.10.11/python-3.10.11-amd64.exe', 'python_installer.exe')"
    if exist python_installer.exe (
        echo Installing Python 3.10... Please follow the installation steps.
        echo IMPORTANT: Please make sure to check 'Add Python to PATH' during installation!
        start /wait python_installer.exe /passive PrependPath=1
        del python_installer.exe
        echo Python installed successfully!
    ) else (
        echo Failed to download Python automatically.
        echo Please visit https://www.python.org/downloads/ to install Python manually.
        pause
        exit /b 1
    )
)

:: Jalankan skrip pemasangan utama
powershell -NoProfile -ExecutionPolicy Bypass -File "./scripts/compile_and_setup_win.ps1"
echo.
echo Installation completed. Please check your Desktop for the shortcut.
pause
