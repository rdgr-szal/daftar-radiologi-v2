import os
import sys
import json
import datetime
import re
from pathlib import Path

# Dynamic Versioning Info
APP_VERSION = "2.1.6"
GITHUB_REPO = "rdgr-szal/daftar-radiologi-v2" # Format: username/repo

# Tangani lokasi fail asas mengikut mod PyInstaller vs Dev
if getattr(sys, 'frozen', False):
    # PyInstaller Bundle Mode
    BUNDLE_DIR = sys._MEIPASS
    CORE_DIR = os.path.join(BUNDLE_DIR, 'core')
    BASE_DIR = os.path.dirname(sys.executable)
else:
    # Source Development Mode
    CORE_DIR = os.path.dirname(os.path.abspath(__file__))
    BUNDLE_DIR = os.path.dirname(CORE_DIR)
    BASE_DIR = BUNDLE_DIR

# Folder Templat & Static
TEMPLATE_FOLDER = os.path.join(BUNDLE_DIR, 'templates')
STATIC_FOLDER = os.path.join(BUNDLE_DIR, 'static')
APP_TEMPLATE_XLSX = os.path.join(BUNDLE_DIR, 'template.xlsx')

# Dapatkan lokasi folder data kekal mengikut OS (OS App Data Directory)
def get_user_data_dir():
    if sys.platform == 'win32':
        appdata = os.getenv('LOCALAPPDATA') or os.getenv('APPDATA') or os.path.expanduser('~')
        return os.path.join(appdata, 'DaftarRadiologi')
    elif sys.platform == 'darwin':
        return os.path.expanduser('~/Library/Application Support/DaftarRadiologi')
    else:
        return os.path.expanduser('~/.local/share/DaftarRadiologi')

if getattr(sys, 'frozen', False):
    USER_DATA_ROOT = get_user_data_dir()
    PENDAFTARAN_DIR = os.path.join(USER_DATA_ROOT, "Pendaftaran")
    
    # Auto-migrasi data legacy dari lokasi BASE_DIR (jika ada)
    legacy_pendaftaran = os.path.join(BASE_DIR, "Pendaftaran")
    if os.path.exists(legacy_pendaftaran) and legacy_pendaftaran != PENDAFTARAN_DIR:
        try:
            os.makedirs(PENDAFTARAN_DIR, exist_ok=True)
            import shutil
            for item in os.listdir(legacy_pendaftaran):
                s = os.path.join(legacy_pendaftaran, item)
                d = os.path.join(PENDAFTARAN_DIR, item)
                if not os.path.exists(d):
                    if os.path.isdir(s):
                        shutil.copytree(s, d)
                    else:
                        shutil.copy2(s, d)
        except Exception as e_mig:
            print(f"[Data Migration Warning] {e_mig}")
else:
    PENDAFTARAN_DIR = os.path.join(BASE_DIR, "Pendaftaran")

CONFIG_PATH = os.path.join(PENDAFTARAN_DIR, "config.json")
EXTRA_DATA_PATH = os.path.join(PENDAFTARAN_DIR, "extra_phris_data.json")
SYNC_QUEUE_PATH = os.path.join(PENDAFTARAN_DIR, "sync_queue.json")
DICOM_WORKLIST_PATH = os.path.join(PENDAFTARAN_DIR, "dicom_worklist.json")
SMRP_TAXONOMY_PATH = os.path.join(CORE_DIR, "smrp_taxonomy.json")

# Pemetaan Bulan Bahasa Melayu untuk Folder & Nama Fail
MONTH_MAP = {
    1: ("01.JANUARI", "JANUARI"),
    2: ("02.FEBRUARI", "FEBRUARI"),
    3: ("03.MAC", "MAC"),
    4: ("04.APRIL", "APRIL"),
    5: ("05.MEI", "MEI"),
    6: ("06.JUN", "JUN"),
    7: ("07.JULAI", "JULAI"),
    8: ("08.OGOS", "OGOS"),
    9: ("09.SEPTEMBER", "SEPTEMBER"),
    10: ("10.OKTOBER", "OKTOBER"),
    11: ("11.NOVEMBER", "NOVEMBER"),
    12: ("12.DISEMBER", "DISEMBER")
}

# Katalog Semua Modaliti Radiologi yang Disokong
ALL_MODALITIES_CATALOG = [
    {"code": "GENERAL RADIOGRAPHY", "name": "General X-Ray", "filename_prefix": "DAFTAR XRAY"},
    {"code": "MOBILE RADIOGRAPHY", "name": "Mobile X-Ray", "filename_prefix": "DAFTAR MOBILE XRAY"},
    {"code": "ULTRASOUND", "name": "Ultrasound", "filename_prefix": "DAFTAR ULTRASOUND"},
    {"code": "MAMMOGRAPHY", "name": "Mammography", "filename_prefix": "DAFTAR MAMMOGRAPHY"},
    {"code": "CT", "name": "CT-Scan", "filename_prefix": "DAFTAR CT SCAN"},
    {"code": "MRI", "name": "MRI", "filename_prefix": "DAFTAR MRI"},
    {"code": "FLUOROSCOPY", "name": "Fluoroscopy", "filename_prefix": "DAFTAR FLUOROSCOPY"},
    {"code": "ANGIOGRAPHY", "name": "Angiography", "filename_prefix": "DAFTAR ANGIOGRAPHY"},
    {"code": "DENTAL", "name": "Dental / OPG", "filename_prefix": "DAFTAR DENTAL"},
    {"code": "BMD", "name": "Bone Mineral Density (BMD)", "filename_prefix": "DAFTAR BMD"}
]

def get_smrp_taxonomy():
    """Membaca data taksonomi SMRP 2.0 (Modality -> Region & Orderable -> Sub_Region & Orderable)."""
    if os.path.exists(SMRP_TAXONOMY_PATH):
        try:
            with open(SMRP_TAXONOMY_PATH, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception as e:
            print(f"[ERROR Config] Ralat membaca smrp_taxonomy.json: {e}")
    return {}

def build_smrp_option_map(active_modalities=None, custom_orderables=None):
    """
    Membina peta optionMap dinamik bagi borang.html berasaskan modaliti aktif SMRP 2.0.
    Stripping trailing Right/Left daripada Sub-Region untuk mengelakkan pertindihan dengan lajur Lateraliti.
    Termasuk sokongan untuk penambahan Bahagian Pemeriksaan Baru (Custom SMRP Orderables).
    """
    tax = get_smrp_taxonomy()
    if not tax:
        return {}

    if not active_modalities:
        active_modalities = ["GENERAL RADIOGRAPHY", "MOBILE - GENERAL RADIOGRAPHY", "ULTRASOUND", "DENTAL"]

    active_upper = [str(m).strip().upper() for m in active_modalities]

    def is_modality_active(mod_name):
        mod_u = mod_name.strip().upper()
        if mod_u in active_upper:
            return True
        for act in active_upper:
            act_u = act.strip().upper()
            if act_u == mod_u:
                return True
            if act_u in ["GENERAL RADIOGRAPHY", "GENERAL"]:
                if mod_u == "GENERAL RADIOGRAPHY":
                    return True
            elif act_u in ["MOBILE RADIOGRAPHY", "MOBILE - GENERAL RADIOGRAPHY"]:
                if mod_u == "MOBILE - GENERAL RADIOGRAPHY":
                    return True
            elif act_u in ["ULTRASOUND", "USG"]:
                if mod_u == "ULTRASOUND":
                    return True
            elif act_u == "MOBILE - ULTRASOUND":
                if mod_u == "MOBILE - ULTRASOUND":
                    return True
            elif act_u in ["DENTAL", "OPG"]:
                if mod_u == "DENTAL":
                    return True
            elif act_u in ["CT", "CT-SCAN", "CT SCAN"]:
                if mod_u == "CT":
                    return True
        return False

    option_map = {}

    for mod_name, mdata in tax.items():
        if not is_modality_active(mod_name):
            continue

        regions = mdata.get("regions", {})
        for reg_name, rdata in regions.items():
            jenis_key = reg_name.strip().upper()
            if jenis_key not in option_map:
                option_map[jenis_key] = []

            sub_regs = rdata.get("sub_regions", [])
            for sr in sub_regs:
                sr_name = sr.get("name", "").strip()
                if sr_name:
                    # Strip trailing Right / Left to prevent duplication with Lateraliti column
                    clean_sr_name = re.sub(r'\s+(Right|Left|RIGHT|LEFT)$', '', sr_name, flags=re.IGNORECASE).strip()
                    if clean_sr_name and not any(o["val"] == clean_sr_name for o in option_map[jenis_key]):
                        option_map[jenis_key].append({
                            "val": clean_sr_name,
                            "label": clean_sr_name
                        })

    # Merge custom SMRP orderables added by user
    if custom_orderables and isinstance(custom_orderables, dict):
        for j_key, c_list in custom_orderables.items():
            upper_k = j_key.strip().upper()
            if upper_k not in option_map:
                option_map[upper_k] = []
            if isinstance(c_list, list):
                for item in c_list:
                    item_str = str(item).strip()
                    if item_str and not any(o["val"] == item_str for o in option_map[upper_k]):
                        option_map[upper_k].append({"val": item_str, "label": f"{item_str} *"})

    return option_map

def ensure_dirs():
    """Memastikan direktori Pendaftaran wujud."""
    os.makedirs(PENDAFTARAN_DIR, exist_ok=True)

def get_unconfigured_default():
    """Mengembalikan konfigurasi bersih bagi pemasangan baru (Belum dikonfigurasi)."""
    return {
        "is_configured": False,
        "facility_type": "KK",  # 'KK' (Klinik Kesihatan) atau 'HOSPITAL' (Hospital/Institut)
        "hospital_scope": "ALL",  # 'ALL' (Keseluruhan Jabatan) atau 'SINGLE' (Modaliti Khusus)
        "single_modality": "General Radiography",
        "klinik_asal": "",
        "singkatan_klinik": "",
        "klinik_rujukan": [],
        "staff": [],
        "default_staff": "",
        # Modaliti aktif SMRP 2.0 yang ditawarkan bagi Klinik Kesihatan
        "active_modalities": [
            "GENERAL RADIOGRAPHY"
        ],
        "consumables": ["-", "CD [1]", "CD [2]", "FILEM 14X17 [1]", "FILEM 10X12 [1]"],
        "custom_starting_xray_no": 0,
        # Konfigurasi Pangkalan Data Pilihan (Offline-First Hybrid Sync)
        "db_config": {
            "enabled": False,
            "provider": "sqlite",  # 'sqlite', 'postgres', 'mysql', 'cloudflare_d1', 'rest_api'
            "host": "localhost",
            "port": 5432,
            "database": "radiologi_db",
            "username": "",
            "password": "",
            "endpoint_url": "",
            "api_key": "",
            "table_prefix": "rad_"
        },
        # Konfigurasi Integrasi DICOM Modality Worklist (MWL Server / Console)
        "dicom_config": {
            "enabled": False,
            "ae_title": "",
            "port": 104,
            "host": "0.0.0.0",
            "default_modality": "CR",
            "console_name": "",
            "console_ae_title": "",
            "console_ip": "",
            "console_port": 104,
            "auto_clear_hours": 24
        }
    }

def load_config():
    """
    Membaca konfigurasi dari config.json.
    Sekiranya fail tidak wujud, pulangkan is_configured = False tanpa data hardcoded KKBBS.
    """
    ensure_dirs()
    if not os.path.exists(CONFIG_PATH):
        return get_unconfigured_default()
    
    try:
        with open(CONFIG_PATH, "r", encoding="utf-8") as f:
            data = json.load(f)
            
            # Pastikan kunci-kunci standard wujud & backward compatibility
            if "staff" not in data:
                # Migrasi dari juru_xray jika wujud
                data["staff"] = data.get("juru_xray", [])
            if "default_staff" not in data:
                data["default_staff"] = data.get("default_juru_xray", "")
            if "juru_xray" not in data:
                data["juru_xray"] = data.get("staff", [])
            if "default_juru_xray" not in data:
                data["default_juru_xray"] = data.get("default_staff", "")
            if "facility_type" not in data:
                data["facility_type"] = "KK"
            if "hospital_scope" not in data:
                data["hospital_scope"] = "ALL"
            if "single_modality" not in data:
                data["single_modality"] = "General Radiography"
            if "active_modalities" not in data:
                data["active_modalities"] = ["GENERAL RADIOGRAPHY"]
            if "consumables" not in data:
                data["consumables"] = ["-", "CD [1]", "CD [2]", "FILEM 14X17 [1]", "FILEM 10X12 [1]"]
            if "custom_starting_xray_no" not in data:
                data["custom_starting_xray_no"] = 0
            if "db_config" not in data:
                data["db_config"] = {
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
                }
            if "dicom_config" not in data:
                data["dicom_config"] = {
                    "enabled": False,
                    "ae_title": "",
                    "port": 104,
                    "host": "0.0.0.0",
                    "default_modality": "CR",
                    "console_name": "",
                    "console_ae_title": "",
                    "console_ip": "",
                    "console_port": 104,
                    "auto_clear_hours": 24
                }
            return data
    except Exception as e:
        print(f"[ERROR Config] Ralat membaca config.json: {e}")
        return get_unconfigured_default()

def save_config(config_data):
    """Menyimpan konfigurasi klinik ke config.json."""
    ensure_dirs()
    try:
        # Selaraskan kedua-dua staff & juru_xray untuk keserasian templat lama
        if "staff" in config_data:
            config_data["juru_xray"] = config_data["staff"]
        if "default_staff" in config_data:
            config_data["default_juru_xray"] = config_data["default_staff"]
            
        with open(CONFIG_PATH, "w", encoding="utf-8") as f:
            json.dump(config_data, f, indent=4, ensure_ascii=False)
        return True
    except Exception as e:
        print(f"[ERROR Config] Ralat menyimpan config.json: {e}")
        return False

# --- PENGURUSAN OFFLINE SYNC QUEUE ---

def load_sync_queue():
    """Membaca rekod tertunggak dari sync_queue.json."""
    ensure_dirs()
    if not os.path.exists(SYNC_QUEUE_PATH):
        return []
    try:
        with open(SYNC_QUEUE_PATH, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception as e:
        print(f"[ERROR SyncQueue] Ralat membaca sync_queue.json: {e}")
        return []

def save_sync_queue(queue_items):
    """Menyimpan senarai giliran larasan ke fail."""
    ensure_dirs()
    try:
        with open(SYNC_QUEUE_PATH, "w", encoding="utf-8") as f:
            json.dump(queue_items, f, indent=4, ensure_ascii=False)
        return True
    except Exception as e:
        print(f"[ERROR SyncQueue] Ralat menyimpan sync_queue.json: {e}")
        return False

def add_to_sync_queue(action_type, payload):
    """
    Menambah satu tugasan pelarasan ke dalam giliran offline.
    action_type: 'INSERT', 'UPDATE', 'CANCEL'
    """
    queue = load_sync_queue()
    item = {
        "id": f"sync-{datetime.datetime.now().strftime('%Y%m%d%H%M%S%f')}",
        "action": action_type,
        "payload": payload,
        "created_at": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "retry_count": 0
    }
    queue.append(item)
    save_sync_queue(queue)
    return item["id"]

def load_extra_phris_data():
    """Membaca data sampingan PHRIS."""
    ensure_dirs()
    if not os.path.exists(EXTRA_DATA_PATH):
        return {}
    try:
        with open(EXTRA_DATA_PATH, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception as e:
        print(f"[ERROR Config] Ralat membaca extra_phris_data.json: {e}")
        return {}

def save_extra_phris_data(data):
    """Menyimpan data sampingan PHRIS."""
    ensure_dirs()
    try:
        with open(EXTRA_DATA_PATH, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=4, ensure_ascii=False)
        return True
    except Exception as e:
        print(f"[ERROR Config] Ralat menyimpan extra_phris_data.json: {e}")
        return False

def parse_mykad(ic_number):
    """
    Ekstraksi automatik Umur & Jantina daripada No. Kad Pengenalan Malaysia (12 digit).
    Format IC: YYMMDD-PB-###G (G = Ganjil untuk Lelaki, Genap untuk Perempuan)
    """
    if not ic_number:
        return None
    clean_ic = str(ic_number).replace("-", "").replace(" ", "").strip()
    if len(clean_ic) != 12 or not clean_ic.isdigit():
        return None
    
    try:
        yy = int(clean_ic[0:2])
        mm = int(clean_ic[2:4])
        dd = int(clean_ic[4:6])
        last_digit = int(clean_ic[-1])
        
        # Penentuan Tahun Lahir (Tahun 2000+ vs 1900+)
        current_year = datetime.datetime.now().year
        short_current_year = current_year % 100
        
        if yy <= short_current_year:
            birth_year = 2000 + yy
        else:
            birth_year = 1900 + yy
            
        today = datetime.date.today()
        birth_date = datetime.date(birth_year, mm, dd)
        
        # Pengiraan Umur
        age = today.year - birth_date.year - ((today.month, today.day) < (birth_date.month, birth_date.day))
        if age < 0:
            age = 0
            
        # Penentuan Jantina (Ganjil = Male, Genap = Female)
        gender = "M" if (last_digit % 2 != 0) else "F"
        gender_label = "LELAKI" if gender == "M" else "PEREMPUAN"
        
        return {
            "valid": True,
            "ic": clean_ic,
            "age": age,
            "gender": gender,
            "gender_label": gender_label,
            "birth_date": birth_date.strftime("%Y-%m-%d")
        }
    except Exception as e:
        print(f"[parse_mykad] Ralat parse IC {ic_number}: {e}")
        return None
