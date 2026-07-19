#!/bin/bash
# rdgr.szal.kkbbs
cd "$(dirname "$0")"
clear
echo "=================================================="
echo "Starting installation of Daftar Radiologi (Mac)..."
echo "=================================================="

# Check if Python3 Installed
python3 --version >/dev/null 2>&1
if [ $? -ne 0 ]; then
    echo "Python not found in your system."
    echo "Try to install Python..."
    if command -v brew >/dev/null 2>&1; then
        echo "Homebrew found. Installing Python..."
        brew install python
    else
        echo "Homebrew not found. Opening web browser to download..."
        open "https://www.python.org/downloads/"
        echo "Please install Python from your web browser, then run this file again."
        read -p "Press [Enter] to exit..."
        exit 1
    fi
fi

chmod +x scripts/compile_and_setup_mac.sh
./scripts/compile_and_setup_mac.sh
echo ""
echo "Installation completed. You can close this window."
read -p "Press [Enter] to exit..."
