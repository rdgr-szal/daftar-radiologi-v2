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

def main():
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

    port = 5005
    url = f"http://127.0.0.1:{port}"

    # Mulakan pelayan Flask di thread latar belakang
    server_thread = threading.Thread(target=start_flask_server, args=(port,), daemon=True)
    server_thread.start()

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
        webview.start()
        print("[DaftarRadiologi] Aplikasi ditutup dengan selamat.")
        sys.exit(0)
    except Exception as e1:
        print(f"[DaftarRadiologi] PyWebView gagal dibuka ({e1}). Mencuba enjin PySide6 QtWebEngine...")

    # 2. Cuba jalankan sebagai PySide6 QtWebEngine Window (Embedded Chromium)
    try:
        from PySide6.QtCore import QUrl
        from PySide6.QtWidgets import QApplication, QMainWindow
        from PySide6.QtWebEngineWidgets import QWebEngineView
        
        print(f"[DaftarRadiologi] Mulakan tetingkap PySide6 QtWebEngine di {url}...")
        qt_app = QApplication(sys.argv)
        main_win = QMainWindow()
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
        print(f"[DaftarRadiologi] QtWebEngine gagal dibuka ({e2}). Membuka pelayar web tempatan...")

    # 3. Fallback terakhir ke pelayar web jika tiada enjin tetingkap GUI ditemui
    threading.Timer(1.2, lambda: webbrowser.open(url)).start()
    import time
    while True:
        time.sleep(1)

if __name__ == '__main__':
    main()
