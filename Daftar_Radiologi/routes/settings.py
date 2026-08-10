import os
import sys
import threading
import datetime
import json
import urllib.request
import urllib.error
import zipfile
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
    active_modalities = data.get('active_modalities', ["General Radiography", "Ultrasound", "Dental"])
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
        "active_modalities": active_modalities if isinstance(active_modalities, list) else [],
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

@settings_bp.route('/api/update/apply-patch', methods=['POST'])
def api_apply_update_patch():
    """
    Kemas kini Aplikasi via Patch File (.zip):
    1. Simpan fail zip tempatan.
    2. Semak dan auto-kesan sekiranya terdapat fail manifest.json (is_major: true) atau perubahan skema database major.
    3. Jika dikesan major update atau is_major = True: jalankan Auto Backup dahulu.
    4. Ekstrak fail zip ke BUNDLE_DIR (melindungi folder Pendaftaran/).
    5. Jalankan init_db_schema jika major update dikesan.
    """
    if 'file' not in request.files:
        return jsonify({"success": False, "message": "No update file (.zip) uploaded."}), 400
        
    file = request.files['file']
    if not file.filename.endswith(".zip"):
        return jsonify({"success": False, "message": "Please upload a valid .zip update file."}), 400

    form_is_major = request.form.get('is_major', 'false').lower() == 'true'
    
    from core.config import BUNDLE_DIR, load_config
    import zipfile
    
    temp_zip = os.path.join(PENDAFTARAN_DIR, "temp_update_patch.zip")
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

    # 2. Ekstrak & Terapkan Kemas Kini
    try:
        with zipfile.ZipFile(temp_zip, 'r') as zipf:
            namelist = zipf.namelist()
            # Keselamatan: Lindungi folder Pendaftaran/
            for member in namelist:
                norm_path = os.path.normpath(member)
                if norm_path.startswith("Pendaftaran") or norm_path.startswith("Pendaftaran/") or norm_path.startswith("Pendaftaran\\"):
                    continue
                zipf.extract(member, BUNDLE_DIR)
                
        if os.path.exists(temp_zip):
            os.remove(temp_zip)
            
        # 3. Kemas kini DB schema jika major update
        if is_major:
            try:
                config = load_config()
                if config.get("db_config", {}).get("enabled", False):
                    init_db_schema(config)
            except Exception as e_db:
                print(f"[Update Major DB Schema Warning] {e_db}")

        detected_label = " (Major Patch Detected)" if auto_detected_major else ""
        return jsonify({
            "success": True,
            "message": f"Application update successfully applied!{detected_label}{backup_msg} Please reload page or restart app if needed."
        })
        
    except Exception as e:
        if os.path.exists(temp_zip):
            os.remove(temp_zip)
        return jsonify({"success": False, "message": f"Error extracting update: {str(e)}"}), 500

@settings_bp.route('/api/update/check-github', methods=['GET'])
def api_check_github_release():
    """
    Semak kemas kini terkini daripada GitHub Release API.
    """
    import ssl
    try:
        url = f"https://api.github.com/repos/{GITHUB_REPO}/releases/latest"
        req = urllib.request.Request(url, headers={'User-Agent': 'DaftarRadiologi-App'})
        
        # SSL Context setup with unverified fallback for environment SSL certificate issues
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
                
                # Cari fail asset zip
                zip_asset = None
                assets = data.get('assets', [])
                for a in assets:
                    aname = a.get('name', '').lower()
                    if aname.endswith('.zip') or aname.endswith('.dmg'):
                        zip_asset = {
                            "name": a.get('name'),
                            "download_url": a.get('browser_download_url'),
                            "size_mb": round(a.get('size', 0) / (1024 * 1024), 2)
                        }
                        # Utamakan fail .zip yang mengandungi App
                        if "app" in aname or "windows" in aname or "macos" in aname:
                            break

                current_ver = APP_VERSION.lstrip('v')
                # Perbandingan ringkas versi
                has_update = tag_name != current_ver and tag_name > current_ver

                return jsonify({
                    "success": True,
                    "current_version": APP_VERSION,
                    "latest_version": tag_name,
                    "release_name": release_name,
                    "has_update": has_update,
                    "body": body,
                    "html_url": html_url,
                    "asset": zip_asset
                })
    except urllib.error.HTTPError as e:
        if e.code == 404:
            return jsonify({"success": False, "message": f"No GitHub Release found for repository '{GITHUB_REPO}'."}), 404
        elif e.code == 403:
            return jsonify({"success": False, "message": f"Error (403). Please try again in 1 hour or upload the update zip file manually."}), 403
        return jsonify({"success": False, "message": f"GitHub API error ({e.code}): {e.reason}"}), 500
    except Exception as e:
        return jsonify({"success": False, "message": f"Failed to check online update: {str(e)}"}), 500

@settings_bp.route('/api/update/apply-online', methods=['POST'])
def api_apply_online_update():
    """
    Muat turun pakej update dari GitHub Release URL & pasang secara automatik.
    """
    import ssl
    data = request.get_json() or {}
    download_url = data.get('download_url', '').strip()
    is_major = data.get('is_major', False)

    if not download_url:
        return jsonify({"success": False, "message": "Pautan muat turun GitHub tidak sah."}), 400

    from core.config import BUNDLE_DIR, load_config

    # 1. Auto backup jika major update
    backup_msg = ""
    if is_major:
        try:
            success_b, msg_b, _ = create_zip_backup()
            if not success_b:
                return jsonify({"success": False, "message": f"Gagal membuat auto backup sebelum major update: {msg_b}"}), 500
            backup_msg = " [Auto Backup Berjaya Dicipta]"
        except Exception as e:
            return jsonify({"success": False, "message": f"Ralat Auto Backup: {str(e)}"}), 500

    # 2. Muat turun fail zip dari GitHub
    temp_zip = os.path.join(PENDAFTARAN_DIR, "temp_github_update.zip")
    try:
        req = urllib.request.Request(download_url, headers={'User-Agent': 'DaftarRadiologi-App'})
        ctx = ssl.create_default_context()
        try:
            ctx.check_hostname = False
            ctx.verify_mode = ssl.CERT_NONE
        except Exception:
            pass

        with urllib.request.urlopen(req, timeout=60, context=ctx) as resp, open(temp_zip, 'wb') as out_file:
            out_file.write(resp.read())
    except Exception as e:
        if os.path.exists(temp_zip):
            os.remove(temp_zip)
        return jsonify({"success": False, "message": f"Ralat memuat turun dari GitHub: {str(e)}"}), 500

    # 3. Ekstrak & Terapkan Kemas Kini
    try:
        with zipfile.ZipFile(temp_zip, 'r') as zipf:
            namelist = zipf.namelist()
            for member in namelist:
                norm_path = os.path.normpath(member)
                if norm_path.startswith("Pendaftaran") or norm_path.startswith("Pendaftaran/") or norm_path.startswith("Pendaftaran\\"):
                    continue # Abaikan folder data
                zipf.extract(member, BUNDLE_DIR)

        if os.path.exists(temp_zip):
            os.remove(temp_zip)

        if is_major:
            try:
                config = load_config()
                if config.get("db_config", {}).get("enabled", False):
                    init_db_schema(config)
            except Exception as e_db:
                print(f"[Online Update Major DB Warning] {e_db}")

        return jsonify({
            "success": True,
            "message": f"Kemas kini Online dari GitHub berjaya diterap!{backup_msg} Sila muat semula halaman."
        })
    except Exception as e:
        if os.path.exists(temp_zip):
            os.remove(temp_zip)
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

@settings_bp.route('/api/dicom/test-echo', methods=['POST'])
def api_dicom_test_echo():
    """Test DICOM C-ECHO connectivity to the Modality Console."""
    try:
        data = request.get_json() or {}
        console_ip = data.get('console_ip', '').strip()
        console_port = data.get('console_port', 104)
        console_ae = data.get('console_ae_title', '').strip()
        my_ae = data.get('ae_title', '').strip() or 'KAUNTER'
        
        if not console_ip or not console_ae:
            config = load_config()
            dicom_cfg = config.get("dicom_config", {})
            if not console_ip:
                console_ip = dicom_cfg.get("console_ip", "").strip()
            if not console_ae:
                console_ae = dicom_cfg.get("console_ae_title", "").strip()
            if not my_ae:
                my_ae = dicom_cfg.get("ae_title", "KAUNTER").strip() or "KAUNTER"

        if not console_ip:
            return jsonify({"success": False, "message": "Modality Console IP address is empty. Please enter the console IP address."})
        if not console_ae:
            return jsonify({"success": False, "message": "Modality Console AE Title is empty. Please enter the console AE Title."})

        from core.dicom_engine import test_dicom_echo_scu
        success, message, ms = test_dicom_echo_scu(
            console_ip=console_ip,
            console_port=console_port,
            console_ae=console_ae,
            my_ae=my_ae
        )
        return jsonify({
            "success": success,
            "message": message,
            "elapsed_ms": ms
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



