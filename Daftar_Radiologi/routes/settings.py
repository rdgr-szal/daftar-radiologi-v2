import os
import sys
import threading
import datetime
from flask import Blueprint, render_template, request, jsonify, redirect, url_for, send_file
from core.config import (
    load_config,
    save_config,
    PENDAFTARAN_DIR,
    MONTH_MAP,
    ALL_MODALITIES_CATALOG,
    load_sync_queue
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
