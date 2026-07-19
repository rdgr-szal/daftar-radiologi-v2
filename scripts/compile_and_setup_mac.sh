#!/bin/bash
# rdgr.szal.kkbbs

# Function to draw progress bar
draw_progress() {
    local pct=$1
    local label=$2
    local bar_len=30
    local filled=$((pct * bar_len / 100))
    local empty=$((bar_len - filled))
    local bar_filled=$(printf "%${filled}s" | tr ' ' '#')
    local bar_empty=$(printf "%${empty}s" | tr ' ' '.')
    printf "\r\033[K[%s%s] %d%% - %s" "$bar_filled" "$bar_empty" "$pct" "$label"
}

echo "=================================================="
echo "      Starting Daftar Radiologi Setup (Mac)       "
echo "=================================================="

# 1. Get project paths
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
PROJECT_DIR="$(dirname "$SCRIPT_DIR")"
DESKTOP_DIR="$HOME/Desktop"
APP_TARGET_DIR="$DESKTOP_DIR/DaftarRadiologi_App"
SHORTCUT_PATH="$DESKTOP_DIR/Daftar Radiologi.command"
LOG_FILE="$PROJECT_DIR/pyinstaller_err.log"

draw_progress 10 "Initializing paths and directories..."
sleep 1

# 2. Activate or create Virtual Environment
draw_progress 20 "Checking virtual environment..."
cd "$PROJECT_DIR"

if [ ! -d ".venv" ]; then
    draw_progress 25 "Creating new virtual environment..."
    python3 -m venv .venv
fi

source .venv/bin/activate

# 3. Always ensure all required packages are installed
# (even if .venv existed before, it may be missing packages)
draw_progress 35 "Checking and installing dependencies..."
python3 -m pip install --quiet --upgrade pip
python3 -m pip install --quiet Flask openpyxl pyinstaller

# Verify pyinstaller is available before proceeding
if ! python3 -m PyInstaller --version > /dev/null 2>&1; then
    echo ""
    echo ""
    echo "ERROR: PyInstaller could not be installed. Check your internet connection."
    exit 1
fi

draw_progress 50 "Preparing compilation..."

# 4. Compile project using PyInstaller
cd "$PROJECT_DIR/Daftar_Radiologi"
rm -rf build dist

# Run pyinstaller in background, save full output to log file
python3 -m PyInstaller --clean -y DaftarRadiologi.spec > "$LOG_FILE" 2>&1 &
PID=$!
pct=55

while kill -0 $PID 2>/dev/null; do
    if [ $pct -lt 88 ]; then
        pct=$((pct + 1))
    fi
    draw_progress $pct "Compiling codebase with PyInstaller (please wait)..."
    sleep 0.6
done

# Capture actual exit code from PyInstaller
wait $PID
PYINSTALLER_EXIT=$?

# If PyInstaller failed - show the actual error log
if [ $PYINSTALLER_EXIT -ne 0 ]; then
    echo ""
    echo ""
    echo "=================================================="
    echo "ERROR: PyInstaller gagal (exit code: $PYINSTALLER_EXIT)"
    echo "Log penuh tersimpan di: $LOG_FILE"
    echo ""
    echo "--- 20 baris terakhir log error ---"
    tail -20 "$LOG_FILE"
    echo "=================================================="
    exit 1
fi

# Verify output folder was actually created
if [ ! -d "dist/DaftarRadiologi" ]; then
    echo ""
    echo ""
    echo "ERROR: Folder dist/DaftarRadiologi tidak ditemui selepas compilation."
    echo "Semak log: $LOG_FILE"
    exit 1
fi

draw_progress 90 "Structuring destination folders..."

# 5. Set up Application Folder on Desktop
if [ -d "$APP_TARGET_DIR" ]; then
    draw_progress 90 "Previous installation found. Keeping 'Pendaftaran' database intact..."
    find "$APP_TARGET_DIR" -mindepth 1 -maxdepth 1 ! -name 'Pendaftaran' -exec rm -rf {} +
else
    mkdir -p "$APP_TARGET_DIR"
fi

# Copy compiled executable and assets
cp -R dist/DaftarRadiologi/* "$APP_TARGET_DIR/"

# Copy Pendaftaran folder only if it doesn't exist (protect existing data)
if [ ! -d "$APP_TARGET_DIR/Pendaftaran" ]; then
    if [ -d "$PROJECT_DIR/Daftar_Radiologi/Pendaftaran" ]; then
        cp -R "$PROJECT_DIR/Daftar_Radiologi/Pendaftaran" "$APP_TARGET_DIR/"
    else
        mkdir -p "$APP_TARGET_DIR/Pendaftaran"
    fi
else
    draw_progress 92 "Database folder exists. Skipping copy to prevent data loss..."
    sleep 0.5
fi

# 6. Create Desktop Launcher (.command)
draw_progress 95 "Creating Desktop launcher shortcut..."
rm -f "$SHORTCUT_PATH"

cat << 'EOF' > "$SHORTCUT_PATH"
#!/bin/bash
# Double-click to launch Daftar Radiologi
cd "$(dirname "$0")/DaftarRadiologi_App"
clear
echo "Starting Daftar Radiologi Application..."
./DaftarRadiologi
EOF

chmod +x "$SHORTCUT_PATH"
chmod +x "$APP_TARGET_DIR/DaftarRadiologi" 2>/dev/null || true

draw_progress 100 "Setup completed successfully!"
echo ""
echo "=================================================="
echo "Setup Berjaya!"
echo "App folder : $APP_TARGET_DIR"
echo "Desktop shortcut: $SHORTCUT_PATH"
echo "=================================================="
