import os
import json
import sqlite3
import datetime
from core.config import PENDAFTARAN_DIR, add_to_sync_queue, load_sync_queue, save_sync_queue

# Path lalai bagi fail pangkalan data SQLite Tempatan / Intranet
LOCAL_SQLITE_PATH = os.path.join(PENDAFTARAN_DIR, "radiologi_local.db")

def get_connection(db_config):
    """
    Membina sambungan ke pangkalan data berdasarkan pembekal (provider) yang mematuhi sektor awam:
    1. SQLite (On-Premise / NAS Shared Drive)
    2. PostgreSQL / MyGovCloud@PDSA (Pusat Data Sektor Awam)
    3. MySQL / MariaDB (Pelayan Intranet Hospital/KK)
    """
    provider = str(db_config.get("provider", "sqlite")).lower().strip()
    
    if provider == "sqlite":
        db_path = db_config.get("database") or LOCAL_SQLITE_PATH
        if not os.path.isabs(db_path):
            db_path = os.path.join(PENDAFTARAN_DIR, db_path)
        conn = sqlite3.connect(db_path, timeout=10)
        conn.row_factory = sqlite3.Row
        return conn, "sqlite"
        
    elif provider in ["postgres", "postgresql", "mygovcloud"]:
        try:
            import psycopg2
            from psycopg2.extras import RealDictCursor
            conn = psycopg2.connect(
                host=db_config.get("host", "localhost"),
                port=int(db_config.get("port", 5432) or 5432),
                dbname=db_config.get("database", "radiologi_db"),
                user=db_config.get("username", "postgres"),
                password=db_config.get("password", ""),
                connect_timeout=5
            )
            return conn, "postgres"
        except ImportError:
            raise Exception("Library 'psycopg2' belum dipasang untuk sambungan PostgreSQL / MyGovCloud.")
            
    elif provider in ["mysql", "mariadb"]:
        try:
            import pymysql
            conn = pymysql.connect(
                host=db_config.get("host", "localhost"),
                port=int(db_config.get("port", 3306) or 3306),
                database=db_config.get("database", "radiologi_db"),
                user=db_config.get("username", "root"),
                password=db_config.get("password", ""),
                connect_timeout=5,
                cursorclass=pymysql.cursors.DictCursor
            )
            return conn, "mysql"
        except ImportError:
            raise Exception("Library 'pymysql' belum dipasang untuk sambungan MySQL / MariaDB.")
            
    else:
        raise Exception(f"Pembekal pangkalan data '{provider}' tidak disokong di bawah pekeliling keselamatan data sektor awam.")

def test_db_connection(db_config):
    """
    Menguji kesahan sambungan ke pangkalan data tempatan / intranet / MyGovCloud.
    Memulangkan (success: bool, message: str)
    """
    if not db_config or not db_config.get("enabled", False):
        return True, "Integrasi Pangkalan Data adalah Luar Talian (Mod Standalone Excel)."
        
    try:
        conn, ptype = get_connection(db_config)
        cur = conn.cursor()
        cur.execute("SELECT 1;")
        res = cur.fetchone()
        cur.close()
        conn.close()
        return True, f"Sambungan pangkalan data ({ptype.upper()}) berjaya dan mematuhi piawaian keselamatan sektor awam!"
    except Exception as e:
        return False, f"Ralat sambungan: {str(e)}"

def init_db_schema(config):
    """
    Menjana skema jadual pendaftaran pesakit & pemeriksaan mengikut piawaian PER.SS-RA 101 & SMRP 2.0.
    """
    db_config = config.get("db_config", {})
    if not db_config.get("enabled", False):
        return True, "Database disabled (Excel Mode)."
        
    prefix = db_config.get("table_prefix", "rad_")
    
    try:
        conn, ptype = get_connection(db_config)
        cur = conn.cursor()
        
        # 1. Jadual Pendaftaran Pesakit & Pemeriksaan Utama (PER.SS-RA 101)
        cur.execute(f"""
        CREATE TABLE IF NOT EXISTS {prefix}patients (
            id VARCHAR(64) PRIMARY KEY,
            xray_no VARCHAR(32) NOT NULL,
            bil_kes VARCHAR(32),
            tarikh DATE NOT NULL,
            lmp VARCHAR(32),
            ic_pasport VARCHAR(32),
            nama VARCHAR(255) NOT NULL,
            umur INT,
            jantina VARCHAR(16),
            warganegara VARCHAR(32),
            kakitangan_kerajaan VARCHAR(64),
            bangsa VARCHAR(64),
            alamat TEXT,
            modality VARCHAR(64),
            jenis_pemeriksaan VARCHAR(128),
            bahagian_pemeriksaan VARCHAR(128),
            lateraliti VARCHAR(32),
            klinik_rujukan VARCHAR(255),
            kategori VARCHAR(64),
            cd_filem VARCHAR(64),
            total_expose INT DEFAULT 1,
            total_reject INT DEFAULT 0,
            pegawai_rujukan VARCHAR(255),
            operator VARCHAR(255),
            catatan TEXT,
            status VARCHAR(32) DEFAULT 'AKTIF',
            cancellation_reason TEXT,
            cancelled_by VARCHAR(255),
            cancelled_at DATETIME,
            facility_name VARCHAR(255),
            facility_code VARCHAR(32),
            created_at DATETIME,
            updated_at DATETIME
        );
        """)
        
        # 2. Indeks carian pantas
        try:
            cur.execute(f"CREATE INDEX IF NOT EXISTS idx_{prefix}xray_no ON {prefix}patients (xray_no);")
            cur.execute(f"CREATE INDEX IF NOT EXISTS idx_{prefix}tarikh ON {prefix}patients (tarikh);")
            cur.execute(f"CREATE INDEX IF NOT EXISTS idx_{prefix}ic ON {prefix}patients (ic_pasport);")
        except Exception:
            pass
            
        conn.commit()
        cur.close()
        conn.close()
        return True, "Skema pangkalan data berjaya dimulakan (Initialised)."
    except Exception as e:
        print(f"[DB Engine ERROR] init_db_schema: {e}")
        return False, str(e)

def sync_patient_record(patient_data, examinations, config):
    """
    Menyimpan rekod pesakit ke pangkalan data sekunder.
    Jika offline, ditambah ke sync_queue.json secara automatik.
    """
    db_config = config.get("db_config", {})
    if not db_config.get("enabled", False):
        return True, "DB Disabled (Excel only)"
        
    payload = {
        "patient_data": patient_data,
        "examinations": examinations,
        "facility_name": config.get("klinik_asal", ""),
        "facility_code": config.get("singkatan_klinik", "")
    }
    
    try:
        conn, ptype = get_connection(db_config)
        prefix = db_config.get("table_prefix", "rad_")
        cur = conn.cursor()
        
        tarikh_str = patient_data.get("tarikh") or datetime.date.today().strftime("%Y-%m-%d")
        nama = str(patient_data.get("nama", "")).upper()
        ic_no = str(patient_data.get("ic_pasport", "")).strip()
        umur = int(patient_data.get("umur", 0) or 0)
        jantina = str(patient_data.get("jantina", "M")).upper()
        warga = str(patient_data.get("warganegara", "YA")).upper()
        gov_staff = str(patient_data.get("kakitangan_kerajaan", "")).strip()
        bangsa = str(patient_data.get("bangsa", "MELAYU")).upper()
        alamat = str(patient_data.get("alamat", "")).upper()
        klinik_rujukan = str(patient_data.get("klinik_rujukan", config.get("klinik_asal", ""))).upper()
        kategori = str(patient_data.get("kategori", "PESAKIT LUAR")).upper()
        pegawai = str(patient_data.get("pegawai_rujukan", "")).upper()
        operator = str(patient_data.get("operator", "")).upper()
        lmp = str(patient_data.get("lmp", "")).strip()
        
        now_str = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        
        for exam in examinations:
            xray_no = str(exam.get("xray_no", "")).strip()
            rec_id = f"{tarikh_str}_{xray_no}_{ic_no}"
            bil_kes = str(exam.get("bil_kes", ""))
            modality = str(exam.get("modality", "General Radiography")).upper()
            jenis = str(exam.get("jenis_pemeriksaan", "DADA")).upper()
            bahagian = str(exam.get("bahagian", exam.get("bahagian_pemeriksaan", "CXR"))).upper()
            lat = str(exam.get("lateraliti", "")).upper()
            cd_filem = str(exam.get("cd_filem", "CD [1]")).upper()
            total_exp = int(exam.get("total_expose", 1) or 1)
            total_rej = int(exam.get("total_reject", 0) or 0)
            catatan = str(exam.get("catatan", "")).upper()
            
            cur.execute(f"""
            INSERT INTO {prefix}patients (
                id, xray_no, bil_kes, tarikh, lmp, ic_pasport, nama, umur, jantina,
                warganegara, kakitangan_kerajaan, bangsa, alamat, modality,
                jenis_pemeriksaan, bahagian_pemeriksaan, lateraliti, klinik_rujukan,
                kategori, cd_filem, total_expose, total_reject, pegawai_rujukan,
                operator, catatan, status, facility_name, facility_code, created_at, updated_at
            ) VALUES (
                ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 'AKTIF', ?, ?, ?, ?
            )
            """, (
                rec_id, xray_no, bil_kes, tarikh_str, lmp, ic_no, nama, umur, jantina,
                warga, gov_staff, bangsa, alamat, modality,
                jenis, bahagian, lat, klinik_rujukan,
                kategori, cd_filem, total_exp, total_rej, pegawai,
                operator, catatan, config.get("klinik_asal", ""), config.get("singkatan_klinik", ""), now_str, now_str
            ))
            
        conn.commit()
        cur.close()
        conn.close()
        return True, "Rekod berjaya disimpan ke Pangkalan Data!"
    except Exception as e:
        print(f"[DB Engine Offline Fallback] Ralat menulis ke DB: {e}. Menambah ke sync_queue.json.")
        add_to_sync_queue("INSERT", payload)
        return False, f"Ralat DB (Disimpan ke Giliran Offline): {str(e)}"

def update_patient_in_db(xray_no, updated_data, config):
    """
    Mengemaskini rekod pesakit sedia ada di pangkalan data.
    """
    db_config = config.get("db_config", {})
    if not db_config.get("enabled", False):
        return True, "DB Disabled"
        
    payload = {
        "xray_no": xray_no,
        "updated_data": updated_data
    }
    
    try:
        conn, ptype = get_connection(db_config)
        prefix = db_config.get("table_prefix", "rad_")
        cur = conn.cursor()
        now_str = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        
        cur.execute(f"""
        UPDATE {prefix}patients
        SET nama = ?, ic_pasport = ?, umur = ?, jantina = ?, bangsa = ?,
            klinik_rujukan = ?, operator = ?, catatan = ?, updated_at = ?
        WHERE xray_no = ?
        """, (
            str(updated_data.get("nama", "")).upper(),
            str(updated_data.get("ic_pasport", "")).strip(),
            int(updated_data.get("umur", 0) or 0),
            str(updated_data.get("jantina", "M")).upper(),
            str(updated_data.get("bangsa", "MELAYU")).upper(),
            str(updated_data.get("klinik_rujukan", "")).upper(),
            str(updated_data.get("operator", "")).upper(),
            str(updated_data.get("catatan", "")).upper(),
            now_str,
            str(xray_no).strip()
        ))
        
        conn.commit()
        cur.close()
        conn.close()
        return True, "Rekod di DB berjaya dikemaskini."
    except Exception as e:
        print(f"[DB Engine ERROR] update_patient_in_db: {e}. Queuing for sync.")
        add_to_sync_queue("UPDATE", payload)
        return False, f"Ralat kemaskini DB: {str(e)}"

def cancel_patient_in_db(xray_no, reason, staff_name, config):
    """
    Menandakan rekod sebagai BATAL di DB (Audit Preservation).
    """
    db_config = config.get("db_config", {})
    if not db_config.get("enabled", False):
        return True, "DB Disabled"
        
    payload = {
        "xray_no": xray_no,
        "reason": reason,
        "cancelled_by": staff_name
    }
    
    try:
        conn, ptype = get_connection(db_config)
        prefix = db_config.get("table_prefix", "rad_")
        cur = conn.cursor()
        now_str = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        
        cur.execute(f"""
        UPDATE {prefix}patients
        SET status = 'BATAL', cancellation_reason = ?, cancelled_by = ?, cancelled_at = ?, updated_at = ?
        WHERE xray_no = ?
        """, (
            str(reason).upper(),
            str(staff_name).upper(),
            now_str,
            now_str,
            str(xray_no).strip()
        ))
        
        conn.commit()
        cur.close()
        conn.close()
        return True, "Rekod di DB berjaya ditandakan BATAL."
    except Exception as e:
        print(f"[DB Engine ERROR] cancel_patient_in_db: {e}. Queuing for sync.")
        add_to_sync_queue("CANCEL", payload)
        return False, f"Ralat pembatalan di DB: {str(e)}"

def process_sync_queue(config):
    """
    Sync semua rekod tertunggak dalam sync_queue.json ke pangkalan data apabila sambungan pulih.
    """
    queue = load_sync_queue()
    if not queue:
        return True, "Tiada rekod tertunggak untuk sync."
        
    db_config = config.get("db_config", {})
    if not db_config.get("enabled", False):
        return False, "Pangkalan data belum diaktifkan dalam Tetapan."
        
    remaining = []
    synced_count = 0
    
    for item in queue:
        action = item.get("action")
        payload = item.get("payload", {})
        
        try:
            if action == "INSERT":
                p_data = payload.get("patient_data")
                exams = payload.get("examinations")
                sync_patient_record(p_data, exams, config)
            elif action == "UPDATE":
                update_patient_in_db(payload.get("xray_no"), payload.get("updated_data"), config)
            elif action == "CANCEL":
                cancel_patient_in_db(payload.get("xray_no"), payload.get("reason"), payload.get("cancelled_by"), config)
                
            synced_count += 1
        except Exception as e:
            item["retry_count"] = item.get("retry_count", 0) + 1
            item["last_error"] = str(e)
            remaining.append(item)
            
    save_sync_queue(remaining)
    return True, f"Berjaya sync {synced_count} rekod. {len(remaining)} rekod masih menunggu."
