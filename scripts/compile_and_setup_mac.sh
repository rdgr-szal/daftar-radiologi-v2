#!/bin/bash
# rdgr.szal.kkbbs

# Ensure script stops on error
set -e

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

# 1. Get project paths (relative to script location)
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
PROJECT_DIR="$(dirname "$SCRIPT_DIR")"
DESKTOP_DIR="$HOME/Desktop"
APP_TARGET_DIR="$DESKTOP_DIR/DaftarRadiologi_App"
SHORTCUT_PATH="$DESKTOP_DIR/Daftar Radiologi.command"

draw_progress 10 "Initializing paths and directories..."
sleep 1

# 2. Activate Virtual Environment
draw_progress 20 "Checking virtual environment..."
cd "$PROJECT_DIR"
if [ -d ".venv" ]; then
    source .venv/bin/activate
else
    draw_progress 25 "Virtual environment not found. Creating a new one..."
    python3 -m venv .venv
    source .venv/bin/activate
    draw_progress 40 "Installing dependencies (Flask, openpyxl, pyinstaller)..."
    pip install --quiet Flask openpyxl pyinstaller
fi

# 3. Compile project using PyInstaller
draw_progress 50 "Preparing compilation..."
cd "$PROJECT_DIR/Daftar_Radiologi"
rm -rf build dist

# Run pyinstaller in the background to show dynamic progress
pyinstaller --clean -y DaftarRadiologi.spec > /dev/null 2>&1 &
PID=$!
pct=60

while kill -0 $PID 2>/dev/null; do
    if [ $pct -lt 90 ]; then
        pct=$((pct + 1))
    fi
    draw_progress $pct "Compiling codebase with PyInstaller (please wait)..."
    sleep 0.5
done
wait $PID

draw_progress 90 "Structuring destination folders..."
# 4. Set up Application Folder on Desktop
if [ -d "$APP_TARGET_DIR" ]; then
    draw_progress 90 "Previous installation found. Keeping 'Pendaftaran' database intact..."
    find "$APP_TARGET_DIR" -mindepth 1 -maxdepth 1 ! -name 'Pendaftaran' -exec rm -rf {} +
else
    mkdir -p "$APP_TARGET_DIR"
fi

# Copy executable and assets
cp -R dist/DaftarRadiologi/* "$APP_TARGET_DIR/"

# Copy templates directory (Pendaftaran) if not already present
if [ ! -d "$APP_TARGET_DIR/Pendaftaran" ]; then
    if [ -d "$PROJECT_DIR/Daftar_Radiologi/Pendaftaran" ]; then
        cp -R "$PROJECT_DIR/Daftar_Radiologi/Pendaftaran" "$APP_TARGET_DIR/"
    else
        mkdir -p "$APP_TARGET_DIR/Pendaftaran"
    fi
else
    draw_progress 92 "Database folder already exists. Skipping templates copy to prevent data loss..."
    sleep 0.5
fi

# 5. Create Desktop Launcher (.command)
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

# Grant executable permissions
chmod +x "$SHORTCUT_PATH"
chmod +x "$APP_TARGET_DIR/DaftarRadiologi"

draw_progress 100 "Setup completed successfully!"
echo ""
echo "=================================================="
echo "Setup Successful!"
echo "App folder created: $APP_TARGET_DIR"
echo "Desktop Shortcut created: $SHORTCUT_PATH"
echo "=================================================="
