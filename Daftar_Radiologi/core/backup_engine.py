import os
import shutil
import zipfile
import datetime
from pathlib import Path
from core.config import PENDAFTARAN_DIR, CONFIG_PATH, load_config, save_config

BACKUP_LOCAL_DIR = os.path.join(PENDAFTARAN_DIR, "backups")

def ensure_backup_dir():
    """Memastikan folder backup tempatan wujud."""
    os.makedirs(BACKUP_LOCAL_DIR, exist_ok=True)

def create_zip_backup(custom_dest_dir=None):
    """
    Mencipta bungkusan sandaran .zip yang mengandungi:
    - Semua fail Excel Buku Daftar (.xlsx)
    - Pangkalan Data Tempatan (radiologi_local.db)
    - Konfigurasi (config.json)
    - Giliran Offline Sync (sync_queue.json)
    
    Menyimpan ke folder tempatan Pendaftaran/backups/ dan menyalin ke MyGovUC Google Drive jika dikonfigurasi.
    """
    ensure_backup_dir()
    config = load_config()
    singkatan = str(config.get("singkatan_klinik", "RAD")).strip().upper() or "RAD"
    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    zip_filename = f"BACKUP_RADIOLOGI_{singkatan}_{timestamp}.zip"
    zip_filepath = os.path.join(BACKUP_LOCAL_DIR, zip_filename)

    try:
        with zipfile.ZipFile(zip_filepath, 'w', zipfile.ZIP_DEFLATED) as zipf:
            for root, dirs, files in os.walk(PENDAFTARAN_DIR):
                # Elakkan menyalin folder backups itu sendiri
                if "backups" in root:
                    continue
                for file in files:
                    if file.endswith(".tmp") or file.startswith("~$"):
                        continue
                    abs_file_path = os.path.join(root, file)
                    rel_path = os.path.relpath(abs_file_path, PENDAFTARAN_DIR)
                    zipf.write(abs_file_path, rel_path)

        # Salin ke MyGovUC Google Drive jika laluan diisi
        mygovuc_config = config.get("mygovuc_backup", {})
        drive_path = (custom_dest_dir or mygovuc_config.get("drive_path", "")).strip()
        drive_copied = False
        drive_msg = ""

        if drive_path:
            try:
                os.makedirs(drive_path, exist_ok=True)
                target_drive_file = os.path.join(drive_path, zip_filename)
                shutil.copy2(zip_filepath, target_drive_file)
                drive_copied = True
                drive_msg = f" dan disalin ke Google Drive ({drive_path})"
            except Exception as de:
                drive_msg = f" (Ralat menyalin ke Google Drive: {str(de)})"

        # Kemaskini rekod tarikh backup terakhir
        if "mygovuc_backup" not in config:
            config["mygovuc_backup"] = {}
        config["mygovuc_backup"]["last_backup"] = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        save_config(config)

        # Bersihkan backup lama yang melebihi tempoh simpanan (retention days)
        clean_old_backups(mygovuc_config.get("retention_days", 30))

        return True, f"Backup berjaya dicipta: {zip_filename}{drive_msg}", zip_filepath

    except Exception as e:
        print(f"[BackupEngine ERROR] create_zip_backup: {e}")
        return False, f"Ralat mencipta backup: {str(e)}", None

def clean_old_backups(retention_days=30):
    """Memadam fail backup tempatan yang melebihi tempoh simpanan (hari)."""
    try:
        ensure_backup_dir()
        now = datetime.datetime.now()
        for f in os.listdir(BACKUP_LOCAL_DIR):
            if f.endswith(".zip") and f.startswith("BACKUP_RADIOLOGI_"):
                fp = os.path.join(BACKUP_LOCAL_DIR, f)
                file_time = datetime.datetime.fromtimestamp(os.path.getmtime(fp))
                if (now - file_time).days > retention_days:
                    os.remove(fp)
                    print(f"[BackupEngine] Backup lama dipadam: {f}")
    except Exception as e:
        print(f"[BackupEngine WARNING] clean_old_backups: {e}")

def restore_from_zip(zip_filepath):
    """
    Ekstrak dan pulihkan semula folder Pendaftaran/ daripada fail backup .zip.
    """
    if not os.path.exists(zip_filepath):
        return False, "Fail backup .zip tidak dijumpai."

    try:
        with zipfile.ZipFile(zip_filepath, 'r') as zipf:
            zipf.extractall(PENDAFTARAN_DIR)
            
        config = load_config()
        return True, f"Data berjaya dipulihkan daripada {os.path.basename(zip_filepath)}!"
    except Exception as e:
        print(f"[BackupEngine ERROR] restore_from_zip: {e}")
        return False, f"Ralat memulihkan data: {str(e)}"

def list_available_backups():
    """Mengembalikan senarai fail backup tempatan yang wujud."""
    ensure_backup_dir()
    backups = []
    try:
        for f in os.listdir(BACKUP_LOCAL_DIR):
            if f.endswith(".zip"):
                fp = os.path.join(BACKUP_LOCAL_DIR, f)
                stat = os.stat(fp)
                size_mb = round(stat.st_size / (1024 * 1024), 2)
                mtime = datetime.datetime.fromtimestamp(stat.st_mtime).strftime("%Y-%m-%d %H:%M:%S")
                backups.append({
                    "filename": f,
                    "filepath": fp,
                    "size_mb": size_mb if size_mb > 0.01 else 0.01,
                    "created_at": mtime
                })
        backups.sort(key=lambda x: x["created_at"], reverse=True)
    except Exception as e:
        print(f"[BackupEngine ERROR] list_available_backups: {e}")
    return backups

def auto_daily_backup_check():
    """
    Semakan automatik harian: Jika backup belum dibuat untuk hari ini, cipta backup automatik.
    """
    config = load_config()
    mygovuc_config = config.get("mygovuc_backup", {})
    if not mygovuc_config.get("auto_daily", True):
        return
        
    last_backup = mygovuc_config.get("last_backup", "")
    today_str = datetime.date.today().strftime("%Y-%m-%d")
    
    if not last_backup or not last_backup.startswith(today_str):
        print(f"[BackupEngine] Menjalankan Backup Harian Automatik MyGovUC bagi {today_str}...")
        create_zip_backup()
