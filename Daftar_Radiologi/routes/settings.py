import os
import sys
import threading
import datetime
import json
import urllib.request
import urllib.error
import zipfile
import shutil
from flask import Blueprint, render_template, request, jsonify, redirect, url_for, send_file
from core.config import (
    load_config,
    save_config,
    PENDAFTARAN_DIR,
    MONTH_MAP,
    ALL_MODALITIES_CATALOG,
    load_sync_queue,
    APP_VERSION,
    GITHUB_REPO
)
from core.excel_engine import get_excel_path, repair_excel_file
from core.db_engine import (
    test_db_connection,
    init_db_schema,
    process_sync_queue
)
from core.backup_engine import (
    create_zip_backup,
    restore_from_zip,
    list_available_backups
)

settings_bp = Blueprint('settings', __name__)

@settings_bp.route('/setup')
def setup_wizard():
    config = load_config()
    return render_template('setup.html', config=config, modalities_catalog=ALL_MODALITIES_CATALOG, current_page='setup')

@settings_bp.route('/settings')
def settings_view():
    config = load_config()
    today = datetime.date.today()
    queue = load_sync_queue()
    pending_sync_count = len(queue)
    backups = list_available_backups()
    return render_template(
        'settings.html',
        config=config,
        modalities_catalog=ALL_MODALITIES_CATALOG,
        current_year=today.year,
        pending_sync_count=pending_sync_count,
        backups=backups,
        app_version=APP_VERSION,
        current_page='settings'
    )

@settings_bp.route('/api/setup', methods=['POST'])
def api_setup():
    data = request.get_json() or {}
    klinik_asal = data.get('klinik_asal', '').strip()
    singkatan = data.get('singkatan_klinik', '').strip().upper()
    facility_type = data.get('facility_type', 'KK').strip().upper()
    hospital_scope = data.get('hospital_scope', 'ALL').strip().upper()
    single_modality = data.get('single_modality', 'General Radiography').strip()
    klinik_rujukan = data.get('klinik_rujukan', [])
    staff_list = data.get('staff', data.get('juru_xray', []))
    default_staff = data.get('default_staff', data.get('default_juru_xray', '')).strip()
    
    active_modalities = data.get('active_modalities')
    if not active_modalities or not isinstance(active_modalities, list):
        if facility_type == 'KK':
            active_modalities = ["GENERAL RADIOGRAPHY"]
        else:
            active_modalities = ["GENERAL RADIOGRAPHY", "MOBILE - GENERAL RADIOGRAPHY", "ULTRASOUND", "DENTAL"]

    consumables_list = data.get('consumables', ["-", "CD [1]", "CD [2]", "FILEM 14X17 [1]", "FILEM 10X12 [1]"])
    db_config = data.get('db_config', {})
    mygovuc_config = data.get('mygovuc_backup', {})
    
    if not klinik_asal or not singkatan:
        return jsonify({"success": False, "message": "Nama Fasiliti dan Singkatan wajib diisi."}), 400
        
    existing_config = load_config()
    config = {
        "is_configured": True,
        "facility_type": facility_type,
        "hospital_scope": hospital_scope,
        "single_modality": single_modality,
        "klinik_asal": klinik_asal,
        "singkatan_klinik": singkatan,
        "klinik_rujukan": klinik_rujukan if isinstance(klinik_rujukan, list) else [],
        "staff": staff_list if isinstance(staff_list, list) else [],
        "default_staff": default_staff,
        "juru_xray": staff_list if isinstance(staff_list, list) else [],
        "default_juru_xray": default_staff,
        "active_modalities": active_modalities,
        "consumables": consumables_list if isinstance(consumables_list, list) else [],
        "custom_smrp_orderables": data.get("custom_smrp_orderables", existing_config.get("custom_smrp_orderables", {})),
        "db_config": db_config if isinstance(db_config, dict) else {
            "enabled": False,
            "provider": "sqlite",
            "host": "localhost",
            "port": 5432,
            "database": "radiologi_db",
            "username": "",
            "password": "",
            "endpoint_url": "",
            "api_key": "",
            "table_prefix": "rad_"
        },
        "mygovuc_backup": mygovuc_config if isinstance(mygovuc_config, dict) else {
            "enabled": True,
            "drive_path": "",
            "auto_daily": True,
            "retention_days": 30,
            "last_backup": ""
        }
    }
    
    save_config(config)
    
    # Inisialisasi fail Excel tahun semasa
    today = datetime.date.today()
    get_excel_path(today)
    
    # Jika database diaktifkan, inisialisasi skema secara automatik
    if config["db_config"].get("enabled", False):
        try:
            init_db_schema(config)
        except Exception as e:
            print(f"[Setup DB Warning] Inisialisasi DB: {e}")
    
    return jsonify({"success": True, "message": "Konfigurasi fasiliti berjaya disimpan!"})

@settings_bp.route('/api/settings/save', methods=['POST'])
def api_save_settings():
    data = request.get_json() or {}
    config = load_config()
    
    config["is_configured"] = True
    config["facility_type"] = data.get('facility_type', config.get('facility_type', 'KK')).strip().upper()
    config["hospital_scope"] = data.get('hospital_scope', config.get('hospital_scope', 'ALL')).strip().upper()
    config["single_modality"] = data.get('single_modality', config.get('single_modality', 'General Radiography')).strip()
    config["klinik_asal"] = data.get('klinik_asal', config.get('klinik_asal', '')).strip()
    config["singkatan_klinik"] = data.get('singkatan_klinik', config.get('singkatan_klinik', '')).strip().upper()
    
    ref_list = data.get('klinik_rujukan')
    if isinstance(ref_list, list):
        config["klinik_rujukan"] = ref_list

    staff_list = data.get('staff', data.get('juru_xray'))
    if isinstance(staff_list, list):
        config["staff"] = staff_list
        config["juru_xray"] = staff_list

    if 'default_staff' in data or 'default_juru_xray' in data:
        def_st = str(data.get('default_staff', data.get('default_juru_xray', '')) or '').strip()
        config["default_staff"] = def_st
        config["default_juru_xray"] = def_st

    if 'active_modalities' in data and isinstance(data.get('active_modalities'), list):
        config["active_modalities"] = data.get('active_modalities')

    if 'consumables' in data and isinstance(data.get('consumables'), list):
        config["consumables"] = data.get('consumables')

    if 'custom_starting_xray_no' in data:
        try:
            config["custom_starting_xray_no"] = int(data.get('custom_starting_xray_no', 0) or 0)
        except (ValueError, TypeError):
            config["custom_starting_xray_no"] = 0

    if 'db_config' in data and isinstance(data.get('db_config'), dict):
        config["db_config"] = data.get('db_config')

    if 'mygovuc_backup' in data and isinstance(data.get('mygovuc_backup'), dict):
        config["mygovuc_backup"] = data.get('mygovuc_backup')

    if 'dicom_config' in data and isinstance(data.get('dicom_config'), dict):
        config["dicom_config"] = data.get('dicom_config')
        # Selaraskan status pelayan DICOM
        try:
            from core.dicom_engine import DicomMWLServerDaemon
            d_cfg = config["dicom_config"]
            daemon = DicomMWLServerDaemon.get_instance()
            if d_cfg.get("enabled", False):
                daemon.restart(
                    host=d_cfg.get("host", "0.0.0.0"),
                    port=d_cfg.get("port", 104),
                    ae_title=d_cfg.get("ae_title", "KAUNTER")
                )
            else:
                daemon.stop()
        except Exception as e_d:
            print(f"[Settings DICOM Sync Warning] {e_d}")

    save_config(config)
    return jsonify({"success": True, "message": "Semua tetapan berjaya dikemaskini!"})

# --- API BACKUP & RESTORE MYGOVUC GOOGLE DRIVE ---

@settings_bp.route('/api/backup/create', methods=['POST'])
def api_create_backup():
    data = request.get_json(silent=True) or {}
    drive_path = data.get('drive_path', '').strip()
    
    if drive_path:
        config = load_config()
        if "mygovuc_backup" not in config:
            config["mygovuc_backup"] = {}
        config["mygovuc_backup"]["drive_path"] = drive_path
        save_config(config)

    success, msg, zip_path = create_zip_backup(custom_dest_dir=drive_path if drive_path else None)
    backups = list_available_backups()
    return jsonify({
        "success": success,
        "message": msg,
        "backups": backups
    }), (200 if success else 500)

@settings_bp.route('/api/backup/list', methods=['GET'])
def api_list_backups():
    backups = list_available_backups()
    return jsonify({"success": True, "backups": backups})

@settings_bp.route('/api/backup/download/<filename>', methods=['GET'])
def api_download_backup(filename):
    from core.backup_engine import BACKUP_LOCAL_DIR
    target_path = os.path.join(BACKUP_LOCAL_DIR, filename)
    if os.path.exists(target_path) and filename.endswith(".zip"):
        return send_file(target_path, as_attachment=True)
    return jsonify({"success": False, "message": "Fail backup tidak dijumpai."}), 404

@settings_bp.route('/api/backup/restore', methods=['POST'])
def api_restore_backup():
    if 'file' not in request.files:
        return jsonify({"success": False, "message": "Tiada fail backup (.zip) dimuat naik."}), 400
    file = request.files['file']
    if not file.filename.endswith(".zip"):
        return jsonify({"success": False, "message": "Hanya fail format .zip sahaja dibenarkan."}), 400
        
    temp_path = os.path.join(PENDAFTARAN_DIR, "temp_restore.zip")
    file.save(temp_path)
    
    success, msg = restore_from_zip(temp_path)
    if os.path.exists(temp_path):
        os.remove(temp_path)
        
    return jsonify({"success": success, "message": msg}), (200 if success else 500)

@settings_bp.route('/api/backup/restore-local/<filename>', methods=['POST'])
def api_restore_local_backup(filename):
    from core.backup_engine import BACKUP_LOCAL_DIR
    target_path = os.path.join(BACKUP_LOCAL_DIR, filename)
    if not os.path.exists(target_path) or not filename.endswith(".zip"):
        return jsonify({"success": False, "message": "Fail backup tidak dijumpai."}), 404
        
    success, msg = restore_from_zip(target_path)
    return jsonify({"success": success, "message": msg}), (200 if success else 500)

@settings_bp.route('/api/db/test', methods=['POST'])
def api_test_db():
    data = request.get_json() or {}
    db_config = data.get('db_config')
    if not db_config:
        config = load_config()
        db_config = config.get('db_config', {})
        
    success, msg = test_db_connection(db_config)
    return jsonify({"success": success, "message": msg}), (200 if success else 400)

@settings_bp.route('/api/db/init', methods=['POST'])
def api_init_db():
    config = load_config()
    data = request.get_json() or {}
    if 'db_config' in data:
        config['db_config'] = data['db_config']
        
    success, msg = init_db_schema(config)
    return jsonify({"success": success, "message": msg}), (200 if success else 500)

@settings_bp.route('/api/db/sync', methods=['POST'])
def api_sync_queue():
    config = load_config()
    success, msg = process_sync_queue(config)
    queue = load_sync_queue()
    return jsonify({
        "success": success,
        "message": msg,
        "pending_count": len(queue)
    })

@settings_bp.route('/api/open-folder', methods=['POST', 'GET'])
def api_open_folder():
    try:
        folder_path = PENDAFTARAN_DIR
        if not os.path.exists(folder_path):
            os.makedirs(folder_path, exist_ok=True)
            
        if os.name == 'nt':  # Windows
            os.startfile(folder_path)
        elif sys.platform == 'darwin':  # macOS
            import subprocess
            subprocess.Popen(['open', folder_path])
        else:  # Linux
            import subprocess
            subprocess.Popen(['xdg-open', folder_path])
            
        return jsonify({"success": True, "message": "Folder Excel dibuka."})
    except Exception as e:
        return jsonify({"success": False, "message": f"Ralat membuka folder: {str(e)}"}), 500

@settings_bp.route('/api/select-folder', methods=['POST', 'GET'])
def api_select_folder():
    try:
        import webview
        if hasattr(webview, 'windows') and webview.windows:
            win = webview.windows[0]
            result = win.create_file_dialog(webview.FOLDER_DIALOG)
            if result and len(result) > 0:
                return jsonify({"success": True, "folder_path": result[0]})
            return jsonify({"success": False, "message": "Tiada folder dipilih."})
    except Exception:
        pass

    try:
        import tkinter as tk
        from tkinter import filedialog
        root = tk.Tk()
        root.withdraw()
        root.attributes('-topmost', True)
        folder_selected = filedialog.askdirectory(title="Pilih Folder MyGovUC Google Drive")
        root.destroy()
        if folder_selected:
            return jsonify({"success": True, "folder_path": folder_selected})
        return jsonify({"success": False, "message": "Tiada folder dipilih."})
    except Exception as e:
        return jsonify({"success": False, "message": f"Ralat membuka dialog folder: {str(e)}"}), 500

@settings_bp.route('/api/excel/repair', methods=['POST'])
def api_repair_excel():
    try:
        data = request.get_json() or {}
        year = data.get('year')
        if year:
            year = int(year)
        else:
            year = datetime.date.today().year

        success, msg, rescued_count = repair_excel_file(year)
        return jsonify({
            "success": success,
            "message": msg,
            "rescued_count": rescued_count
        }), (200 if success else 500)
    except Exception as e:
        return jsonify({"success": False, "message": f"Ralat semasa pemulihan fail Excel: {str(e)}"}), 500

@settings_bp.route('/api/settings/add-custom-orderable', methods=['POST'])
def api_add_custom_orderable():
    """Menambah Jenis atau Bahagian Pemeriksaan Baru ke dalam konfigurasi."""
    data = request.get_json() or {}
    jenis = str(data.get('jenis_pemeriksaan', '')).strip().upper()
    bahagian = str(data.get('bahagian_pemeriksaan', '')).strip()

    if not jenis:
        return jsonify({'success': False, 'message': 'Sila pilih atau masukkan Jenis Pemeriksaan.'}), 400

    config = load_config()
    custom_map = config.get('custom_smrp_orderables', {})
    if not isinstance(custom_map, dict):
        custom_map = {}

    if jenis not in custom_map:
        custom_map[jenis] = []

    msg = f"Kategori jenis pemeriksaan '{jenis}' berjaya ditambah."

    if bahagian:
        if bahagian not in custom_map[jenis]:
            custom_map[jenis].append(bahagian)
            msg = f"Bahagian pemeriksaan '{bahagian}' berjaya ditambah di bawah '{jenis}'."

    config['custom_smrp_orderables'] = custom_map
    save_config(config)

    return jsonify({
        'success': True,
        'message': msg
    })

@settings_bp.route('/api/settings/delete-custom-orderable', methods=['POST'])
def api_delete_custom_orderable():
    """Padam Jenis/Bahagian Pemeriksaan Kustom daripada konfigurasi."""
    data = request.get_json() or {}
    jenis = str(data.get('jenis_pemeriksaan', '')).strip().upper()
    bahagian = str(data.get('bahagian_pemeriksaan', '')).strip()

    if not jenis:
        return jsonify({'success': False, 'message': 'Jenis Pemeriksaan diperlukan.'}), 400

    config = load_config()
    custom_map = config.get('custom_smrp_orderables', {})
    if not isinstance(custom_map, dict):
        custom_map = {}

    msg = "Tiada rekod dipadam."

    if jenis in custom_map:
        if bahagian:
            if bahagian in custom_map[jenis]:
                custom_map[jenis].remove(bahagian)
                msg = f"Bahagian '{bahagian}' berjaya dipadam daripada '{jenis}'."
        else:
            del custom_map[jenis]
            msg = f"Kategori jenis pemeriksaan '{jenis}' berjaya dipadam."

    config['custom_smrp_orderables'] = custom_map
    save_config(config)

    return jsonify({
        'success': True,
        'message': msg
    })

@settings_bp.route('/api/shutdown', methods=['POST'])
def api_shutdown():
    def shutdown_server():
        os._exit(0)
    threading.Timer(0.5, shutdown_server).start()
    return jsonify({"success": True, "message": "Aplikasi ditutup secara selamat."})

def get_clean_zip_member_path(member, zip_namelist):
    """
    Mengambil jalur relatif fail dalam arkib zip dengan membuang folder utama paling atas jika wujud
    (contoh: 'DaftarRadiologi-v2-win10/templates/patient_list.html' -> 'templates/patient_list.html').
    """
    norm_member = member.replace('\\', '/').strip('/')
    if not norm_member:
        return ""

    all_items = [m.replace('\\', '/').strip('/') for m in zip_namelist if m.replace('\\', '/').strip('/')]
    prefix_candidate = None
    if all_items:
        first_item = all_items[0]
        parts = first_item.split('/')
        if len(parts) > 1:
            candidate = parts[0]
            if candidate.lower() not in ['templates', 'static', 'core', 'routes', 'pendaftaran']:
                if all(item.startswith(candidate + '/') or item == candidate for item in all_items):
                    prefix_candidate = candidate + '/'

    if prefix_candidate and norm_member.startswith(prefix_candidate.rstrip('/')):
        if norm_member == prefix_candidate.rstrip('/'):
            return ""
        rel_member = norm_member[len(prefix_candidate):]
    else:
        rel_member = norm_member

    return rel_member

# ==============================================================================
# SELF-UPDATE ENGINE (PyInstaller frozen build) - staging + restart updater
# ------------------------------------------------------------------------------
# For a frozen onedir build the running code executes from inside the .exe /
# .app bundle, NOT from loose source files. Writing files into BASE_DIR does
# NOT update the running app. To update properly we must:
#   1) extract the downloaded release bundle into a staging dir,
#   2) spawn a tiny detached updater carrying (PID, staging, install dir),
#   3) force-exit the app so file locks are released,
#   4) updater waits for exit -> swaps the bundle -> relaunches the app.
# ==============================================================================

def _running_frozen():
    return bool(getattr(sys, 'frozen', False))

def _frozen_install_layout():
    """Return a dict describing the frozen install layout, or None if not frozen."""
    if not _running_frozen():
        return None
    exec_path = os.path.abspath(sys.executable)
    exec_name = os.path.basename(exec_path)
    if sys.platform == 'darwin':
        # macOS onedir: .../DaftarRadiologi.app/Contents/MacOS/DaftarRadiologi
        app_dir = os.path.abspath(os.path.join(os.path.dirname(exec_path), '..', '..'))
        return {
            "platform": "mac",
            "root": app_dir,
            "is_app": app_dir.lower().endswith('.app'),
            "exec_name": exec_name,
        }
    # Windows / Linux onedir: all runtime files sit next to the executable
    root = os.path.dirname(exec_path)
    return {
        "platform": "windows",
        "root": root,
        "is_app": False,
        "exec_name": exec_name,
    }

def _extract_zip_into(zip_path, dest_root, keep_top_app=False):
    """
    Safely extract zip into dest_root. Returns number of files extracted.
    Protects against zip-slip and never writes into Pendaftaran/.
    When keep_top_app is True (macOS), a top-level '*.app' folder is preserved
    as the container so the whole bundle can be swapped at once.
    """
    extracted = 0
    with zipfile.ZipFile(zip_path, 'r') as zipf:
        namelist = zipf.namelist()
        for member in namelist:
            norm_member = member.replace('\\', '/').strip('/')
            if not norm_member:
                continue

            if keep_top_app:
                # Keep the first component (the .app bundle) as the container
                parts = norm_member.split('/')
                rel_member = norm_member
                if len(parts) > 1 and parts[0].lower().endswith('.app'):
                    rel_member = norm_member  # preserve .app/Contents/...
                elif len(parts) == 1:
                    rel_member = norm_member
            else:
                rel_member = get_clean_zip_member_path(member, namelist)
                if not rel_member:
                    continue

            norm_path = os.path.normpath(rel_member)
            path_parts = norm_path.replace('\\', '/').split('/')
            if 'Pendaftaran' in path_parts or 'pendaftaran' in path_parts:
                continue

            target_file_path = os.path.abspath(os.path.join(dest_root, norm_path))
            if not target_file_path.startswith(os.path.abspath(dest_root)):
                continue

            if rel_member.endswith('/') or rel_member.endswith('\\'):
                os.makedirs(target_file_path, exist_ok=True)
                continue

            os.makedirs(os.path.dirname(target_file_path), exist_ok=True)
            try:
                with zipf.open(member) as source, open(target_file_path, 'wb') as target:
                    shutil.copyfileobj(source, target)
                extracted += 1
            except Exception:
                continue
    return extracted

def _write_windows_updater(updater_path, pid, src, dst, exe):
    esc = lambda s: s.replace('"', '""')
    content = (
        '@echo off\r\n'
        'setlocal EnableExtensions\r\n'
        f'set "SRC={esc(src)}"\r\n'
        f'set "DST={esc(dst)}"\r\n'
        f'set "EXE={esc(exe)}"\r\n'
        f'set "PID={int(pid)}"\r\n'
        ':WAITLOOP\r\n'
        'tasklist /FI "PID eq %PID%" 2>nul | find "%PID%" >nul\r\n'
        'if not errorlevel 1 (\r\n'
        '  timeout /t 1 /nobreak >nul\r\n'
        '  goto WAITLOOP\r\n'
        ')\r\n'
        'robocopy "%SRC%" "%DST%" /E /R:1 /W:1 /NFL /NDL /NJH /NJS >nul 2>nul\r\n'
        'if errorlevel 8 xcopy /E /Y /I /Q "%SRC%\\*" "%DST%\\" >nul 2>nul\r\n'
        'rd /S /Q "%SRC%" >nul 2>nul\r\n'
        'start "" "%DST%\\%EXE%"\r\n'
        'endlocal\r\n'
        'exit /b\r\n'
    )
    with open(updater_path, 'w', encoding='utf-8') as f:
        f.write(content)
    return updater_path

def _write_macos_updater(updater_path, pid, src_app, dst_app, staging_dir):
    esc = lambda s: s.replace('"', '\\"')
    content = (
        '#!/bin/bash\n'
        f'APP_PID="{int(pid)}"\n'
        f'SRC="{esc(src_app)}"\n'
        f'DST="{esc(dst_app)}"\n'
        f'STAGING="{esc(staging_dir)}"\n'
        'while /bin/kill -0 "$APP_PID" 2>/dev/null; do /bin/sleep 1; done\n'
        '/bin/sleep 1\n'
        '/bin/rm -rf "$DST"\n'
        '/bin/cp -R "$SRC" "$DST"\n'
        '/bin/rm -rf "$STAGING"\n'
        '/usr/bin/open "$DST" >/dev/null 2>&1 || /usr/bin/open -a "$DST" >/dev/null 2>&1\n'
        'exit 0\n'
    )
    with open(updater_path, 'w', encoding='utf-8') as f:
        f.write(content)
    os.chmod(updater_path, 0o755)
    return updater_path

def _perform_frozen_self_update(zip_path, is_major):
    """
    Stage the downloaded release bundle and spawn a detached updater that
    swaps the running app and relaunches it. Returns (success, message, needs_restart).
    """
    import subprocess
    layout = _frozen_install_layout()
    if layout is None:
        return False, "Not a packaged (frozen) build.", False

    staging_root = os.path.join(layout["root"], "__daftar_staging__")
    if os.path.exists(staging_root):
        try:
            shutil.rmtree(staging_root)
        except Exception:
            pass
    os.makedirs(staging_root, exist_ok=True)

    keep_top_app = layout["is_app"]
    try:
        extracted = _extract_zip_into(zip_path, staging_root, keep_top_app=keep_top_app)
    except Exception as e:
        try:
            shutil.rmtree(staging_root)
        except Exception:
            pass
        return False, f"Gagal mengekstrak bundle update: {e}", False

    if extracted <= 0:
        try:
            shutil.rmtree(staging_root)
        except Exception:
            pass
        return False, "Bundle update tidak mengandungi fail yang sah.", False

    pid = os.getpid()

    try:
        if layout["platform"] == "mac" and layout["is_app"]:
            # Locate the staged .app bundle (may be nested under a prefix dir)
            app_candidate = staging_root
            try:
                for dirpath, dirnames, filenames in os.walk(staging_root):
                    for d in dirnames:
                        if d.lower().endswith('.app'):
                            app_candidate = os.path.join(dirpath, d)
                            break
                    if os.path.abspath(app_candidate) != os.path.abspath(staging_root):
                        break
            except Exception:
                pass
            updater_path = os.path.join(staging_root, "__daftar_selfupdate.sh")
            _write_macos_updater(updater_path, pid, app_candidate, layout["root"], staging_root)
            with open(os.devnull, 'w') as devnull:
                subprocess.Popen(
                    ['/bin/bash', updater_path],
                    stdin=devnull, stdout=devnull, stderr=devnull,
                    close_fds=True, start_new_session=True,
                )
            if is_major:
                threading.Timer(1.0, _try_init_db_after_update).start()
            threading.Timer(3.5, os._exit, args=(0,)).start()
            return True, "Bundle update telah dipentaskan. Aplikasi akan dimulakan semula secara automatik.", True

        # Windows / Linux onedir: swap staged contents over install root
        import tempfile
        updater_path = os.path.join(tempfile.gettempdir(), "daftar_selfupdate.bat")
        _write_windows_updater(updater_path, pid, staging_root, layout["root"], layout["exec_name"])
        devnull = open(os.devnull, 'w')
        CREATE_NO_WINDOW = 0x08000000
        DETACHED_PROCESS = 0x00000008
        try:
            subprocess.Popen(
                ['cmd.exe', '/c', updater_path],
                stdin=devnull, stdout=devnull, stderr=devnull,
                close_fds=True, creationflags=DETACHED_PROCESS | CREATE_NO_WINDOW,
            )
        finally:
            try:
                devnull.close()
            except Exception:
                pass
        if is_major:
            threading.Timer(1.0, _try_init_db_after_update).start()
        threading.Timer(3.5, os._exit, args=(0,)).start()
        return True, "Bundle update telah dipentaskan. Aplikasi akan ditutup dan dimulakan semula secara automatik.", True

    except Exception as e:
        return False, f"Ralat memulakan updater: {e}", False

def _try_init_db_after_update():
    try:
        config = load_config()
        if config.get("db_config", {}).get("enabled", False):
            init_db_schema(config)
    except Exception as e_db:
        print(f"[Self-Update Major DB Warning] {e_db}")

@settings_bp.route('/api/update/apply-patch', methods=['POST'])
def api_apply_update_patch():
    """
    Kemas kini Aplikasi via Patch File (.zip):
    1. Simpan fail zip tempatan dalam folder temp sistem.
    2. Semak dan auto-kesan sekiranya terdapat fail manifest.json (is_major: true) atau perubahan skema database major.
    3. Jika dikesan major update atau is_major = True: jalankan Auto Backup dahulu.
    4. Ekstrak fail zip ke BASE_DIR dan BUNDLE_DIR (melindungi folder Pendaftaran/ dan fail yang dikunci Windows).
    5. Jalankan init_db_schema jika major update dikesan.
    """
    if 'file' not in request.files:
        return jsonify({"success": False, "message": "No update file (.zip) uploaded."}), 400
        
    file = request.files['file']
    if not file.filename.endswith(".zip"):
        return jsonify({"success": False, "message": "Please upload a valid .zip update file."}), 400

    form_is_major = request.form.get('is_major', 'false').lower() == 'true'
    
    from core.config import BASE_DIR, BUNDLE_DIR, PENDAFTARAN_DIR, load_config
    import zipfile
    import tempfile
    
    temp_zip = os.path.join(tempfile.gettempdir(), f"temp_update_patch_{int(datetime.datetime.now().timestamp())}.zip")
    try:
        file.save(temp_zip)
    except Exception as e:
        return jsonify({"success": False, "message": f"Error saving uploaded file: {str(e)}"}), 500

    # Auto-detect sekiranya patch mengandungi petunjuk Major Update
    auto_detected_major = False
    try:
        with zipfile.ZipFile(temp_zip, 'r') as zipf:
            namelist = zipf.namelist()
            for member in namelist:
                fname = os.path.basename(member).lower()
                if fname in ["manifest.json", "update_manifest.json"]:
                    try:
                        m_data = json.loads(zipf.read(member).decode('utf-8'))
                        if m_data.get("is_major", False) or m_data.get("requires_backup", False):
                            auto_detected_major = True
                            break
                    except Exception:
                        pass
                if "migration" in fname or "schema" in fname or "db_engine.py" in fname:
                    auto_detected_major = True
    except Exception as e_check:
        print(f"[Update Auto-Detect Warning] {e_check}")

    is_major = form_is_major or auto_detected_major

    # 1. Auto Backup jika Major Update dikesan
    backup_msg = ""
    if is_major:
        try:
            success_b, msg_b, _ = create_zip_backup()
            if not success_b:
                if os.path.exists(temp_zip): os.remove(temp_zip)
                return jsonify({"success": False, "message": f"Failed to create auto backup before major update: {msg_b}"}), 500
            backup_msg = " [Auto Backup Successfully Created]"
        except Exception as e:
            if os.path.exists(temp_zip): os.remove(temp_zip)
            return jsonify({"success": False, "message": f"Auto Backup error: {str(e)}"}), 500

    # Packaged (frozen) build => full bundle self-update with automatic restart
    if _running_frozen():
        run_ok, run_msg, needs_restart = _perform_frozen_self_update(temp_zip, is_major)
        if not run_ok:
            if os.path.exists(temp_zip): os.remove(temp_zip)
            return jsonify({"success": False, "message": run_msg}), 500
        auto_clean = ""
        try:
            if os.path.exists(temp_zip): os.remove(temp_zip)
        except Exception:
            auto_clean = " (Fail temp dibersihkan secara automatik oleh updater)"
        return jsonify({
            "success": True,
            "needs_restart": True,
            "message": f"{run_msg}{backup_msg}{auto_clean} Aplikasi akan ditutup & dimulakan semula."
        })

    # 2. Ekstrak & Terapkan Kemas Kini dengan Selamat (Safety Extraction)
    target_dirs = [BASE_DIR]
    if BUNDLE_DIR and os.path.abspath(BUNDLE_DIR) != os.path.abspath(BASE_DIR) and os.path.exists(BUNDLE_DIR):
        target_dirs.append(BUNDLE_DIR)

    extracted_count = 0
    locked_files = []

    try:
        with zipfile.ZipFile(temp_zip, 'r') as zipf:
            namelist = zipf.namelist()

            for member in namelist:
                rel_member = get_clean_zip_member_path(member, namelist)
                if not rel_member:
                    continue

                norm_path = os.path.normpath(rel_member)
                path_parts = norm_path.replace('\\', '/').split('/')

                # 🛡️ Perlindungan Pangkalan Data: Jangan sekali-kali ganti folder Pendaftaran/
                if 'Pendaftaran' in path_parts or 'pendaftaran' in path_parts:
                    continue

                for t_dir in target_dirs:
                    target_file_path = os.path.abspath(os.path.join(t_dir, rel_member))
                    
                    # 🛡️ Zip-Slip Security Safeguard
                    if not target_file_path.startswith(os.path.abspath(t_dir)):
                        continue

                    if rel_member.endswith('/') or rel_member.endswith('\\'):
                        os.makedirs(target_file_path, exist_ok=True)
                        continue

                    os.makedirs(os.path.dirname(target_file_path), exist_ok=True)

                    # Tulis fail dengan perlindungan Windows file lock
                    try:
                        with zipf.open(member) as source, open(target_file_path, 'wb') as target:
                            shutil.copyfileobj(source, target)
                        extracted_count += 1
                    except (PermissionError, OSError):
                        # Fail sedang dikunci (cth: fail .exe yang sedang berjalan di Windows)
                        locked_files.append(os.path.basename(rel_member))
                        try:
                            fallback_path = target_file_path + ".new"
                            with zipf.open(member) as source, open(fallback_path, 'wb') as target:
                                shutil.copyfileobj(source, target)
                        except Exception:
                            pass
                    except Exception as e_extract_single:
                        print(f"[Update Extraction File Warning] {member}: {e_extract_single}")

        if os.path.exists(temp_zip):
            try: os.remove(temp_zip)
            except Exception: pass
            
        # 3. Kemas kini DB schema jika major update
        if is_major:
            try:
                config = load_config()
                if config.get("db_config", {}).get("enabled", False):
                    init_db_schema(config)
            except Exception as e_db:
                print(f"[Update Major DB Schema Warning] {e_db}")

        detected_label = " (Major Patch Detected)" if auto_detected_major else ""
        locked_notice = ""
        if locked_files:
            unique_locked = list(set(locked_files))
            locked_notice = f" (Note: {', '.join(unique_locked)} locked by Windows, will take full effect on app restart)"

        return jsonify({
            "success": True,
            "message": f"Application update successfully applied ({extracted_count} files updated)!{detected_label}{backup_msg}{locked_notice} Page will now reload."
        })
        
    except Exception as e:
        if os.path.exists(temp_zip):
            os.remove(temp_zip)
        return jsonify({"success": False, "message": f"Error extracting update package: {str(e)}"}), 500

# --- ONLINE UPDATE TRACKER & PLATFORM MATCHING ---

ONLINE_UPDATE_TRACKER = {
    "state": "idle",
    "progress_percent": 0,
    "downloaded_bytes": 0,
    "total_bytes": 0,
    "downloaded_mb": 0.0,
    "total_mb": 0.0,
    "message": "",
    "filepath": "",
    "filename": "",
    "is_major": False
}

@settings_bp.route('/api/update/check-github', methods=['GET'])
def api_check_github_release():
    """
    Semak kemas kini terkini daripada GitHub Release API mengikut OS dan senibina.
    """
    import ssl
    import platform
    try:
        url = f"https://api.github.com/repos/{GITHUB_REPO}/releases/latest"
        req = urllib.request.Request(url, headers={'User-Agent': 'DaftarRadiologi-App'})
        
        ctx = ssl.create_default_context()
        try:
            ctx.check_hostname = False
            ctx.verify_mode = ssl.CERT_NONE
        except Exception:
            pass

        with urllib.request.urlopen(req, timeout=8, context=ctx) as response:
            if response.status == 200:
                data = json.loads(response.read().decode('utf-8'))
                tag_name = data.get('tag_name', '').lstrip('v')
                release_name = data.get('name', '')
                body = data.get('body', '')
                html_url = data.get('html_url', '')

                current_sys = sys.platform  # 'win32', 'darwin', 'linux'
                machine_arch = platform.machine().lower() # 'arm64', 'x86_64', 'amd64'

                target_keyword = ""
                platform_label = ""
                if current_sys == 'win32':
                    win_rel = platform.release()
                    if win_rel in ['7', '8', '8.1'] or (hasattr(sys, 'getwindowsversion') and sys.getwindowsversion().major < 10):
                        target_keyword = "win-legacy"
                        platform_label = "Windows 7 / 8 / 8.1 Legacy x64"
                    else:
                        target_keyword = "win10"
                        platform_label = "Windows 10 / 11 x64"
                elif current_sys == 'darwin':
                    if 'arm64' in machine_arch or 'aarch64' in machine_arch:
                        target_keyword = "mac-applesilicon"
                        platform_label = "macOS Apple Silicon (M1/M2/M3/M4)"
                    else:
                        target_keyword = "mac-intel"
                        platform_label = "macOS Intel"
                else:
                    target_keyword = "zip"
                    platform_label = "Generic Desktop OS"

                zip_asset = None
                matched_asset = None
                fallback_asset = None

                assets = data.get('assets', [])
                for a in assets:
                    aname = a.get('name', '').lower()
                    asset_dict = {
                        "name": a.get('name'),
                        "download_url": a.get('browser_download_url'),
                        "size_mb": round(a.get('size', 0) / (1024 * 1024), 2)
                    }

                    if target_keyword and target_keyword in aname and (aname.endswith('.zip') or aname.endswith('.dmg')):
                        matched_asset = asset_dict
                        break
                    
                    if current_sys == 'win32' and 'win' in aname and aname.endswith('.zip'):
                        fallback_asset = asset_dict
                    elif current_sys == 'darwin' and 'mac' in aname and (aname.endswith('.zip') or aname.endswith('.dmg')):
                        fallback_asset = asset_dict
                    elif not fallback_asset and (aname.endswith('.zip') or aname.endswith('.dmg')):
                        fallback_asset = asset_dict

                zip_asset = matched_asset or fallback_asset

                current_ver = APP_VERSION.lstrip('v')
                has_update = tag_name != current_ver and tag_name > current_ver

                return jsonify({
                    "success": True,
                    "current_version": APP_VERSION,
                    "latest_version": tag_name,
                    "release_name": release_name,
                    "has_update": has_update,
                    "body": body,
                    "html_url": html_url,
                    "asset": zip_asset,
                    "platform_label": platform_label
                })
    except urllib.error.HTTPError as e:
        if e.code == 404:
            return jsonify({"success": False, "message": f"No GitHub Release found for repository '{GITHUB_REPO}'."}), 404
        elif e.code == 403:
            return jsonify({"success": False, "message": f"Error (403). Please try again in 1 hour or upload the update zip file manually."}), 403
        return jsonify({"success": False, "message": f"GitHub API error ({e.code}): {e.reason}"}), 500
    except Exception as e:
        return jsonify({"success": False, "message": f"Failed to check online update: {str(e)}"}), 500

@settings_bp.route('/api/update/online-status', methods=['GET'])
def api_update_online_status():
    """Status muat turun dan pemasangan online semasa."""
    return jsonify({
        "success": True,
        "tracker": ONLINE_UPDATE_TRACKER
    })

@settings_bp.route('/api/update/start-download', methods=['POST'])
def api_start_online_download():
    """
    Mulakan muat turun fail update GitHub di latar belakang dengan indikator progress.
    """
    import ssl
    data = request.get_json() or {}
    download_url = data.get('download_url', '').strip()
    filename = data.get('filename', 'DaftarRadiologi_Update.zip').strip()
    is_major = data.get('is_major', False)

    if not download_url:
        return jsonify({"success": False, "message": "Pautan muat turun GitHub tidak sah."}), 400

    # Tentukan folder muat turun pengguna (Downloads folder)
    user_downloads = os.path.join(os.path.expanduser('~'), 'Downloads')
    if not os.path.exists(user_downloads):
        user_downloads = PENDAFTARAN_DIR

    dest_filepath = os.path.join(user_downloads, filename)

    ONLINE_UPDATE_TRACKER["state"] = "downloading"
    ONLINE_UPDATE_TRACKER["progress_percent"] = 0
    ONLINE_UPDATE_TRACKER["downloaded_bytes"] = 0
    ONLINE_UPDATE_TRACKER["total_bytes"] = 0
    ONLINE_UPDATE_TRACKER["downloaded_mb"] = 0.0
    ONLINE_UPDATE_TRACKER["total_mb"] = 0.0
    ONLINE_UPDATE_TRACKER["message"] = "Starting update package download..."
    ONLINE_UPDATE_TRACKER["filepath"] = dest_filepath
    ONLINE_UPDATE_TRACKER["filename"] = filename
    ONLINE_UPDATE_TRACKER["is_major"] = is_major

    def download_worker():
        try:
            req = urllib.request.Request(download_url, headers={'User-Agent': 'DaftarRadiologi-App'})
            ctx = ssl.create_default_context()
            try:
                ctx.check_hostname = False
                ctx.verify_mode = ssl.CERT_NONE
            except Exception:
                pass

            with urllib.request.urlopen(req, timeout=60, context=ctx) as response, open(dest_filepath, 'wb') as out_file:
                total_bytes = int(response.headers.get('content-length', 0))
                ONLINE_UPDATE_TRACKER["total_bytes"] = total_bytes
                ONLINE_UPDATE_TRACKER["total_mb"] = round(total_bytes / (1024 * 1024), 2)
                
                downloaded = 0
                chunk_size = 65536
                while True:
                    chunk = response.read(chunk_size)
                    if not chunk:
                        break
                    out_file.write(chunk)
                    downloaded += len(chunk)
                    ONLINE_UPDATE_TRACKER["downloaded_bytes"] = downloaded
                    ONLINE_UPDATE_TRACKER["downloaded_mb"] = round(downloaded / (1024 * 1024), 2)
                    if total_bytes > 0:
                        pct = int((downloaded / total_bytes) * 100)
                        ONLINE_UPDATE_TRACKER["progress_percent"] = min(pct, 99)
                        ONLINE_UPDATE_TRACKER["message"] = f"Downloading update package... {pct}% ({ONLINE_UPDATE_TRACKER['downloaded_mb']} MB / {ONLINE_UPDATE_TRACKER['total_mb']} MB)"

            ONLINE_UPDATE_TRACKER["state"] = "ready"
            ONLINE_UPDATE_TRACKER["progress_percent"] = 100
            ONLINE_UPDATE_TRACKER["message"] = "Download Completed! Ready to install update."
        except Exception as e_dl:
            ONLINE_UPDATE_TRACKER["state"] = "error"
            ONLINE_UPDATE_TRACKER["message"] = f"Download failed: {str(e_dl)}"
            if os.path.exists(dest_filepath):
                try: os.remove(dest_filepath)
                except Exception: pass

    threading.Thread(target=download_worker, daemon=True).start()

    return jsonify({
        "success": True,
        "message": "Update package download started in background.",
        "filepath": dest_filepath
    })

@settings_bp.route('/api/update/apply-downloaded', methods=['POST'])
def api_apply_downloaded_update():
    """
    Terapkan kemas kini daripada fail ZIP yang telah dimuat turun & PADAM AUTOMATIK fail ZIP tersebut.
    """
    data = request.get_json() or {}
    filepath = data.get('filepath', '').strip() or ONLINE_UPDATE_TRACKER.get("filepath", "")
    is_major = data.get('is_major', False) or ONLINE_UPDATE_TRACKER.get("is_major", False)

    if not filepath or not os.path.exists(filepath):
        return jsonify({"success": False, "message": "Fail update yang dimuat turun tidak ditemui di lokasi muat turun."}), 400

    from core.config import BASE_DIR, BUNDLE_DIR, load_config

    ONLINE_UPDATE_TRACKER["state"] = "applying"
    ONLINE_UPDATE_TRACKER["progress_percent"] = 10
    ONLINE_UPDATE_TRACKER["message"] = "Preparing system auto-backup..."

    # 1. Auto backup jika major update
    backup_msg = ""
    if is_major:
        try:
            success_b, msg_b, _ = create_zip_backup()
            if not success_b:
                ONLINE_UPDATE_TRACKER["state"] = "error"
                return jsonify({"success": False, "message": f"Gagal membuat auto backup sebelum major update: {msg_b}"}), 500
            backup_msg = " [Auto Backup Berjaya Dicipta]"
        except Exception as e:
            ONLINE_UPDATE_TRACKER["state"] = "error"
            return jsonify({"success": False, "message": f"Ralat Auto Backup: {str(e)}"}), 500

    ONLINE_UPDATE_TRACKER["progress_percent"] = 40
    ONLINE_UPDATE_TRACKER["message"] = "Extracting update package & patching application..."

    # Packaged (frozen) build => full bundle self-update with automatic restart
    if _running_frozen():
        run_ok, run_msg, needs_restart = _perform_frozen_self_update(filepath, is_major)
        ONLINE_UPDATE_TRACKER["state"] = "completed" if run_ok else "error"
        if not run_ok:
            ONLINE_UPDATE_TRACKER["message"] = run_msg
            return jsonify({"success": False, "message": run_msg}), 500
        ONLINE_UPDATE_TRACKER["message"] = run_msg
        return jsonify({
            "success": True,
            "needs_restart": True,
            "message": f"{run_msg}{backup_msg} Aplikasi akan ditutup & dimulakan semula secara automatik."
        })

    # 2. Ekstrak & Terapkan Kemas Kini dengan Selamat (Safety Extraction)
    target_dirs = [BASE_DIR]
    if BUNDLE_DIR and os.path.abspath(BUNDLE_DIR) != os.path.abspath(BASE_DIR) and os.path.exists(BUNDLE_DIR):
        target_dirs.append(BUNDLE_DIR)

    extracted_count = 0
    locked_files = []

    try:
        with zipfile.ZipFile(filepath, 'r') as zipf:
            namelist = zipf.namelist()

            for member in namelist:
                rel_member = get_clean_zip_member_path(member, namelist)
                if not rel_member:
                    continue

                norm_path = os.path.normpath(rel_member)
                path_parts = norm_path.replace('\\', '/').split('/')

                if 'Pendaftaran' in path_parts or 'pendaftaran' in path_parts:
                    continue

                for t_dir in target_dirs:
                    target_file_path = os.path.abspath(os.path.join(t_dir, rel_member))
                    if not target_file_path.startswith(os.path.abspath(t_dir)):
                        continue

                    if rel_member.endswith('/') or rel_member.endswith('\\'):
                        os.makedirs(target_file_path, exist_ok=True)
                        continue

                    os.makedirs(os.path.dirname(target_file_path), exist_ok=True)

                    try:
                        with zipf.open(member) as source, open(target_file_path, 'wb') as target:
                            shutil.copyfileobj(source, target)
                        extracted_count += 1
                    except (PermissionError, OSError):
                        locked_files.append(os.path.basename(rel_member))
                        try:
                            fallback_path = target_file_path + ".new"
                            with zipf.open(member) as source, open(fallback_path, 'wb') as target:
                                shutil.copyfileobj(source, target)
                        except Exception:
                            pass
                    except Exception as e_extract_single:
                        print(f"[Online Update Extraction File Warning] {member}: {e_extract_single}")

        ONLINE_UPDATE_TRACKER["progress_percent"] = 85
        ONLINE_UPDATE_TRACKER["message"] = "Cleaning up downloaded update file..."

        # 3. AUTO DELETEDownloaded Update File
        if os.path.exists(filepath):
            try:
                os.remove(filepath)
                print(f"[Online Update] Auto-deleted downloaded file: {filepath}")
            except Exception as e_del:
                print(f"[Online Update Auto Delete Warning] {e_del}")

        if is_major:
            try:
                config = load_config()
                if config.get("db_config", {}).get("enabled", False):
                    init_db_schema(config)
            except Exception as e_db:
                print(f"[Online Update Major DB Warning] {e_db}")

        locked_notice = ""
        if locked_files:
            unique_locked = list(set(locked_files))
            locked_notice = f" (Nota: {', '.join(unique_locked)} dikunci oleh Windows, akan beroperasi penuh selepas aplikasi dibuka semula)"

        ONLINE_UPDATE_TRACKER["state"] = "completed"
        ONLINE_UPDATE_TRACKER["progress_percent"] = 100
        ONLINE_UPDATE_TRACKER["message"] = "Update Completed Successfully! Reloading page..."

        return jsonify({
            "success": True,
            "message": f"Kemas kini Online dari GitHub berjaya diterap ({extracted_count} fail)!{backup_msg}{locked_notice} Fail muat turun telah dipadam secara automatik."
        })
    except Exception as e:
        ONLINE_UPDATE_TRACKER["state"] = "error"
        ONLINE_UPDATE_TRACKER["message"] = f"Ralat mengekstrak fail: {str(e)}"
        return jsonify({"success": False, "message": f"Ralat mengekstrak fail dari GitHub: {str(e)}"}), 500

# ==============================================================================
# DICOM MODALITY WORKLIST (MWL) & CONSOLE INTEGRATION API
# ==============================================================================

@settings_bp.route('/api/dicom/status', methods=['GET'])
def api_dicom_status():
    """Mendapatkan status semasa DICOM MWL Server dan bilangan item worklist."""
    try:
        from core.dicom_engine import DicomMWLServerDaemon, get_local_lan_ip, load_dicom_worklist
        daemon = DicomMWLServerDaemon.get_instance()
        status = daemon.get_status()
        config = load_config()
        dicom_cfg = config.get("dicom_config", {})
        
        status["config_enabled"] = dicom_cfg.get("enabled", False)
        status["config_port"] = dicom_cfg.get("port", 104)
        status["config_ae"] = dicom_cfg.get("ae_title", "KAUNTER")
        status["console_ae"] = dicom_cfg.get("console_ae_title", "CARESTREAM")
        status["console_ip"] = dicom_cfg.get("console_ip", "")
        status["console_port"] = dicom_cfg.get("console_port", 104)
        return jsonify({"success": True, "data": status})
    except Exception as e:
        return jsonify({"success": False, "message": str(e)}), 500

@settings_bp.route('/api/dicom/toggle', methods=['POST'])
def api_dicom_toggle():
    """Menghidupkan atau mematikan pelayan DICOM MWL."""
    try:
        data = request.get_json() or {}
        enabled = bool(data.get('enabled', False))
        
        config = load_config()
        if "dicom_config" not in config:
            config["dicom_config"] = {}
        config["dicom_config"]["enabled"] = enabled
        if data.get('ae_title'):
            config["dicom_config"]["ae_title"] = data.get('ae_title').strip()
        if data.get('port'):
            config["dicom_config"]["port"] = int(data.get('port'))
        save_config(config)
        
        from core.dicom_engine import DicomMWLServerDaemon
        daemon = DicomMWLServerDaemon.get_instance()
        
        if enabled:
            host = config["dicom_config"].get("host", "0.0.0.0")
            port = config["dicom_config"].get("port", 104)
            ae_title = config["dicom_config"].get("ae_title", "KAUNTER")
            ok, msg = daemon.start(host=host, port=port, ae_title=ae_title)
            return jsonify({"success": ok, "message": msg, "running": daemon.is_running})
        else:
            ok, msg = daemon.stop()
            return jsonify({"success": ok, "message": msg, "running": False})
    except Exception as e:
        return jsonify({"success": False, "message": f"Server error: {str(e)}"}), 500

@settings_bp.route('/api/dicom/test-connection', methods=['POST'])
def api_dicom_test_connection():
    """
    Universal DICOM Connection Test Endpoint.
    Tests full 5-stage connectivity: local server, ARP/network, C-ECHO (DICOM Ping), and C-FIND MWL inquiry.
    """
    try:
        data = request.get_json() or {}
        console_ip = data.get('console_ip')
        console_port = data.get('console_port')
        console_ae = data.get('console_ae_title')
        my_ae = data.get('ae_title')
        local_port = data.get('port')

        from core.dicom_engine import test_dicom_connection
        report = test_dicom_connection(
            console_ip=console_ip,
            console_port=console_port,
            console_ae=console_ae,
            my_ae=my_ae,
            local_port=local_port
        )
        return jsonify(report)
    except Exception as e:
        return jsonify({
            "ok": False,
            "testedAt": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "summary": "Internal error executing DICOM test.",
            "failure_reason": str(e),
            "checks": [{"name": "Diagnostic Execution", "ok": False, "detail": str(e)}]
        }), 500

@settings_bp.route('/api/dicom/test-echo', methods=['POST'])
def api_dicom_test_echo():
    """Test DICOM C-ECHO connectivity to the Modality Console (Backwards-compatible)."""
    try:
        data = request.get_json() or {}
        console_ip = data.get('console_ip')
        console_port = data.get('console_port')
        console_ae = data.get('console_ae_title')
        my_ae = data.get('ae_title')
        local_port = data.get('port')

        from core.dicom_engine import test_dicom_connection
        report = test_dicom_connection(
            console_ip=console_ip,
            console_port=console_port,
            console_ae=console_ae,
            my_ae=my_ae,
            local_port=local_port
        )
        return jsonify({
            "success": report.get("ok", False),
            "message": report.get("summary") or report.get("failure_reason", ""),
            "report": report
        })
    except Exception as e:
        return jsonify({"success": False, "message": f"DICOM C-ECHO error: {str(e)}"}), 500

@settings_bp.route('/api/dicom/worklist', methods=['GET'])
def api_dicom_worklist():
    """Get active patient examinations in the DICOM Worklist queue."""
    try:
        from core.dicom_engine import load_dicom_worklist
        items = load_dicom_worklist()
        return jsonify({"success": True, "count": len(items), "worklist": items})
    except Exception as e:
        return jsonify({"success": False, "message": str(e)}), 500

@settings_bp.route('/api/dicom/clear-worklist', methods=['POST'])
def api_dicom_clear_worklist():
    """Clear the DICOM Worklist queue."""
    try:
        from core.dicom_engine import clear_dicom_worklist
        clear_dicom_worklist()
        return jsonify({"success": True, "message": "DICOM Worklist queue cleared successfully."})
    except Exception as e:
        return jsonify({"success": False, "message": str(e)}), 500

# ==============================================================================
# DICOM MPPS (MODALITY PERFORMED PROCEDURE STEP) & REJECT ANALYSIS API
# ==============================================================================

@settings_bp.route('/api/mpps/records', methods=['GET'])
def api_mpps_records():
    """Mendapatkan senarai transaksi MPPS dengan penapis status, tarikh dan carian teks."""
    try:
        from core.mpps_engine import get_mpps_records_list
        limit = int(request.args.get('limit', 50))
        offset = int(request.args.get('offset', 0))
        status = request.args.get('status', 'ALL')
        query = request.args.get('q', '')
        start_date = request.args.get('start_date', '')
        end_date = request.args.get('end_date', '')

        result = get_mpps_records_list(
            limit=limit,
            offset=offset,
            status=status,
            query=query,
            start_date=start_date,
            end_date=end_date
        )
        return jsonify({"success": True, "data": result})
    except Exception as e:
        return jsonify({"success": False, "message": str(e)}), 500

@settings_bp.route('/api/mpps/records/<path:sop_uid>', methods=['GET'])
def api_mpps_record_details(sop_uid):
    """Mendapatkan perincian transaksi MPPS termasuk data audit DICOM Dataset JSON."""
    try:
        from core.mpps_engine import get_mpps_record_details
        details = get_mpps_record_details(sop_uid)
        if not details:
            return jsonify({"success": False, "message": "Rekod MPPS tidak ditemui."}), 404
        return jsonify({"success": True, "data": details})
    except Exception as e:
        return jsonify({"success": False, "message": str(e)}), 500

@settings_bp.route('/api/mpps/export', methods=['GET'])
def api_mpps_export_csv():
    """Mengeksport senarai transaksi MPPS dan Reject Analysis ke format fail CSV."""
    import io
    import csv
    from flask import Response
    try:
        from core.mpps_engine import get_mpps_records_list
        status = request.args.get('status', 'ALL')
        query = request.args.get('q', '')
        start_date = request.args.get('start_date', '')
        end_date = request.args.get('end_date', '')

        result = get_mpps_records_list(limit=5000, offset=0, status=status, query=query, start_date=start_date, end_date=end_date)
        records = result.get("records", [])

        output = io.StringIO()
        writer = csv.writer(output)
        
        # Header CSV
        writer.writerow([
            "ID", "SOP Instance UID", "Accession Number", "Patient ID", "Patient Name",
            "Modality", "Station AE", "Station Name", "Status", "Start Date", "Start Time",
            "End Date", "End Time", "Total Images", "Reject Count", "Reject Categories",
            "Status Reason", "Comments", "Created At"
        ])

        for r in records:
            writer.writerow([
                r.get("id"),
                r.get("sop_instance_uid"),
                r.get("accession_number"),
                r.get("patient_id"),
                r.get("patient_name"),
                r.get("modality"),
                r.get("station_ae"),
                r.get("station_name"),
                r.get("status"),
                r.get("start_date"),
                r.get("start_time"),
                r.get("end_date"),
                r.get("end_time"),
                r.get("total_images_count"),
                r.get("reject_count"),
                r.get("reject_categories"),
                r.get("status_reason"),
                r.get("comments"),
                r.get("created_at")
            ])

        output.seek(0)
        filename = f"mpps_reject_analysis_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
        return Response(
            output.getvalue(),
            mimetype="text/csv",
            headers={"Content-Disposition": f"attachment;filename={filename}"}
        )
    except Exception as e:
        return jsonify({"success": False, "message": str(e)}), 500

@settings_bp.route('/api/import/preview', methods=['POST'])
def api_import_preview():
    if 'file' not in request.files:
        return jsonify({"success": False, "message": "Tiada fail dimuat naik."}), 400
        
    file = request.files['file']
    if not file or file.filename == '':
        return jsonify({"success": False, "message": "Fail tidak sah."}), 400
        
    filename = file.filename
    ext = os.path.splitext(filename)[1].lower()
    if ext not in ['.xlsx', '.xls', '.csv']:
        return jsonify({"success": False, "message": "Sila muat naik fail Excel (.xlsx/.xls) atau CSV (.csv)."}), 400

    import tempfile
    temp_dir = os.path.join(tempfile.gettempdir(), "temp_import")
    os.makedirs(temp_dir, exist_ok=True)
    temp_path = os.path.join(temp_dir, f"import_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}{ext}")
    file.save(temp_path)
    
    from core.import_engine import parse_uploaded_file, auto_suggest_mapping
    res = parse_uploaded_file(temp_path)
    if not res.get("success"):
        return jsonify(res), 400
        
    suggested = auto_suggest_mapping(res.get("headers", []))
    res["suggested_mapping"] = suggested
    res["temp_file_path"] = temp_path
    return jsonify(res)

@settings_bp.route('/api/import/execute', methods=['POST'])
def api_import_execute():
    data = request.get_json() or {}
    temp_file_path = data.get("temp_file_path")
    column_mapping = data.get("column_mapping", {})
    sheet_name = data.get("sheet_name")
    
    if not temp_file_path or not os.path.exists(temp_file_path):
        return jsonify({"success": False, "message": "Fail sesi import telah tamat tempoh. Sila muat naik semula."}), 400
        
    from core.import_engine import process_data_migration
    result = process_data_migration(temp_file_path, column_mapping, sheet_name=sheet_name)
    
    # Padam fail sementara
    try:
        os.remove(temp_file_path)
    except Exception:
        pass
        
    return jsonify(result)



