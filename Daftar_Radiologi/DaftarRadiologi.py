import os
import sys
import threading
import webbrowser
from flask import Flask

from core.config import TEMPLATE_FOLDER, STATIC_FOLDER
from routes.registration import registration_bp
from routes.patients import patients_bp
from routes.phris import phris_bp
from routes.dashboard import dashboard_bp
from routes.export import export_bp
from routes.settings import settings_bp

app = Flask(__name__, template_folder=TEMPLATE_FOLDER, static_folder=STATIC_FOLDER)

# Register Modular Blueprints
app.register_blueprint(registration_bp)
app.register_blueprint(patients_bp)
app.register_blueprint(phris_bp)
app.register_blueprint(dashboard_bp)
app.register_blueprint(export_bp)
app.register_blueprint(settings_bp)

def start_flask_server(port=5005):
    """Menjalankan pelayan Flask tempatan."""
    app.run(host='127.0.0.1', port=port, debug=False, use_reloader=False)

def ensure_windows_webview2():
    """Memeriksa dan memasang Microsoft Edge WebView2 Runtime secara senyap jika dijalankan pada Windows."""
    if sys.platform != 'win32':
        return

    # Check registry for WebView2 Runtime
    has_webview2 = False
    try:
        import winreg
        key_paths = [
            r"SOFTWARE\WOW6432Node\Microsoft\EdgeUpdate\Clients\{F3017226-FE2A-4295-8BDF-00C3A9A7E4C5}",
            r"SOFTWARE\Microsoft\EdgeUpdate\Clients\{F3017226-FE2A-4295-8BDF-00C3A9A7E4C5}"
        ]
        for path in key_paths:
            try:
                key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, path)
                winreg.CloseKey(key)
                has_webview2 = True
                break
            except Exception:
                continue
    except Exception as e:
        print(f"[WebView2 Check Warning] {e}")

    if has_webview2:
        print("[DaftarRadiologi] WebView2 Runtime dikesan.")
        return

    print("[DaftarRadiologi] WebView2 Runtime tidak dijumpai. Memulakan muat turun & pemasangan automatik...")
    try:
        import urllib.request
        import subprocess

        # 1. Semak fail installer terbina jika ada
        installer_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "MicrosoftEdgeWebview2Setup.exe")
        
        # 2. Jika tidak ada fail terbina, muat turun dari Microsoft jika ada talian internet
        if not os.path.exists(installer_path):
            temp_dir = os.environ.get("TEMP", os.path.dirname(os.path.abspath(__file__)))
            installer_path = os.path.join(temp_dir, "MicrosoftEdgeWebview2Setup.exe")
            url = "https://go.microsoft.com/fwlink/p/?LinkId=2124703" # Microsoft Evergreen Installer
            print(f"[WebView2 AutoInstall] Muat turun installer dari {url}...")
            urllib.request.urlretrieve(url, installer_path)

        # 3. Jalankan silent installation
        if os.path.exists(installer_path):
            print("[WebView2 AutoInstall] Memasang WebView2 Runtime secara senyap (silent mode)...")
            subprocess.run([installer_path, "/silent", "/install"], check=True)
            print("[WebView2 AutoInstall] Pemasangan WebView2 berjaya!")
    except Exception as err:
        print(f"[WebView2 AutoInstall Warning] Gagal memasang WebView2 secara automatik: {err}")

def apply_pending_updates():
    """
    Memeriksa dan menggantikan fail .new yang dicipta semasa kemas kini aplikasi
    (disebabkan Windows file locking pada fail .exe / .dll semasa aplikasi sedang berjalan).
    """
    try:
        from core.config import BASE_DIR
        for root, dirs, files in os.walk(BASE_DIR):
            for file in files:
                if file.endswith('.new'):
                    new_file_path = os.path.join(root, file)
                    target_file_path = os.path.join(root, file[:-4])
                    try:
                        if os.path.exists(target_file_path):
                            os.remove(target_file_path)
                        os.rename(new_file_path, target_file_path)
                        print(f"[Pending Update Applied] {file} -> {os.path.basename(target_file_path)}")
                    except Exception as e_rename:
                        print(f"[Pending Update Warning] Gagal menggantikan {file}: {e_rename}")
    except Exception as e:
        print(f"[Pending Update System Warning] {e}")

def find_available_port(preferred_port=5005):
    import socket
    try:
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            if s.connect_ex(('127.0.0.1', preferred_port)) != 0:
                return preferred_port
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            s.bind(('127.0.0.1', 0))
            return s.getsockname()[1]
    except Exception:
        return preferred_port

def main():
    # 0. Semak jika dijalankan terus dari fail ZIP (Temp Directory Guard)
    exec_path = sys.executable if getattr(sys, 'frozen', False) else os.path.abspath(__file__)
    if "AppData\\Local\\Temp" in exec_path or "AppData/Local/Temp" in exec_path:
        try:
            import tkinter as tk
            from tkinter import messagebox
            root = tk.Tk()
            root.withdraw()
            messagebox.showwarning(
                "Zip Extraction Required / Sila Ekstrak ZIP",
                "You are running this application directly from inside a ZIP file.\n\n"
                "Please EXTRACT the ZIP folder to your Desktop or Documents folder before launching.\n\n"
                "Anda menjalankan aplikasi ini terus dari dalam fail ZIP. Sila EXTRACT folder ZIP ke Desktop atau Documents terlebih dahulu."
            )
        except Exception:
            pass
        sys.exit(1)

    # 0.5 Terapkan kemas kini fail .new tergendala akibat Windows File Lock
    apply_pending_updates()

    # Semakan & Cipta Backup Harian Automatik MyGovUC
    try:
        from core.backup_engine import auto_daily_backup_check
        auto_daily_backup_check()
    except Exception as e:
        print(f"[Main AutoBackup Warning] {e}")

    # Mulakan Pelayan DICOM MWL (jika diaktifkan dalam tetapan)
    try:
        from core.dicom_engine import start_dicom_server_from_config
        start_dicom_server_from_config()
    except Exception as e:
        print(f"[Main DICOM Server Warning] {e}")

    port = find_available_port(5005)
    url = f"http://127.0.0.1:{port}"

    # Mulakan pelayan Flask di thread latar belakang
    server_thread = threading.Thread(target=start_flask_server, args=(port,), daemon=True)
    server_thread.start()

    # Pastikan WebView2 tersedia jika di Windows
    ensure_windows_webview2()

    # 1. Cuba jalankan sebagai PyWebView Native Desktop Window
    try:
        import webview
        print(f"[DaftarRadiologi] Mulakan tetingkap PyWebView di {url}...")
        window = webview.create_window(
            title="Daftar Radiologi (PER.SS-RA 101 Compliance)",
            url=url,
            width=1152,
            height=1080,
            min_size=(648, 608),
            resizable=True
        )
        # Cuba guna edgehtml / cef / qt jika mshtml/pythonnet bermasalah
        webview.start()
        print("[DaftarRadiologi] Aplikasi ditutup dengan selamat.")
        sys.exit(0)
    except Exception as e1:
        print(f"[DaftarRadiologi] PyWebView gagal dibuka ({type(e1).__name__}: {e1}). Mencuba enjin PySide6 QtWebEngine...")

    # 2. Cuba jalankan sebagai PySide6 QtWebEngine Window (Embedded Chromium)
    try:
        from PySide6.QtCore import QUrl
        from PySide6.QtWidgets import QApplication, QMainWindow
        from PySide6.QtWebEngineWidgets import QWebEngineView
        
        from PySide6.QtGui import QIcon
        print(f"[DaftarRadiologi] Mulakan tetingkap PySide6 QtWebEngine di {url}...")
        qt_app = QApplication(sys.argv)
        icon_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'static', 'img', 'logo.png')
        if os.path.exists(icon_path):
            qt_app.setWindowIcon(QIcon(icon_path))
        main_win = QMainWindow()
        if os.path.exists(icon_path):
            main_win.setWindowIcon(QIcon(icon_path))
        browser = QWebEngineView()
        browser.setUrl(QUrl(url))
        main_win.setCentralWidget(browser)
        main_win.setWindowTitle("Daftar Radiologi (PER.SS-RA 101 Compliance)")
        main_win.resize(1152, 900)
        main_win.show()
        qt_app.exec()
        print("[DaftarRadiologi] Aplikasi Qt ditutup dengan selamat.")
        sys.exit(0)
    except Exception as e2:
        print(f"[DaftarRadiologi] QtWebEngine gagal dibuka ({type(e2).__name__}: {e2}). Membuka pelayar web tempatan...")

    # 3. Fallback terakhir ke pelayar web jika tiada enjin tetingkap GUI ditemui
    threading.Timer(1.2, lambda: webbrowser.open(url)).start()
    import time
    while True:
        time.sleep(1)

if __name__ == '__main__':
    main()
