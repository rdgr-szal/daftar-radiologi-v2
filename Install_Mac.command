#!/bin/bash
# rdgr.szal.kkbbs
cd "$(dirname "$0")"
clear
echo "=================================================="
echo "Starting installation of Daftar Radiologi (Mac)..."
echo "=================================================="

# Check if Python3 Installed
if ! python3 --version > /dev/null 2>&1; then
    echo "Python tidak ditemui dalam sistem anda."
    if command -v brew > /dev/null 2>&1; then
        echo "Homebrew ditemui. Memasang Python..."
        brew install python
    else
        echo "Homebrew tidak ditemui. Membuka pelayar web untuk muat turun..."
        open "https://www.python.org/downloads/"
        echo "Sila pasang Python dari pelayar web anda, kemudian jalankan fail ini semula."
        read -p "Tekan [Enter] untuk keluar..."
        exit 1
    fi
fi

chmod +x scripts/compile_and_setup_mac.sh
./scripts/compile_and_setup_mac.sh
RESULT=$?

echo ""
if [ $RESULT -eq 0 ]; then
    echo "Installation completed. You can close this window."
else
    echo "=================================================="
    echo "INSTALLATION FAILED (exit code: $RESULT)"
    echo "Sila semak mesej error di atas dan hubungi pembangun."
    echo "=================================================="
fi
read -p "Press [Enter] to exit..."
