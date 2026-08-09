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

    port = 5005
    url = f"http://127.0.0.1:{port}"
    
    # 1. Cuba jalankan sebagai PyWebView Native Desktop Window
    try:
        import webview
        
        # Mulakan pelayan Flask di thread latar belakang
        server_thread = threading.Thread(target=start_flask_server, args=(port,), daemon=True)
        server_thread.start()
        
        print(f"[DaftarRadiologi] Starting Desktop Native Window at {url}...")
        
        # Cipta Tetingkap Desktop Native (Nisbah Skrin 2580x1992 MacBook)
        window = webview.create_window(
            title="Daftar Radiologi (PER.SS-RA 101 Compliance)",
            url=url,
            width=1152,
            height=1080,
            min_size=(648, 608),
            resizable=True
        )
        
        webview.start()
        print("[DaftarRadiologi] Apps closed. Exiting safely.")
        sys.exit(0)

    except ImportError:
        print("[DaftarRadiologi] App not found. Switching to browser mode...")
        threading.Timer(1.5, lambda: webbrowser.open(url)).start()
        app.run(host='127.0.0.1', port=port, debug=False)

if __name__ == '__main__':
    main()
