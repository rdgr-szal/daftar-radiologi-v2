import os
import json
import sqlite3
import datetime
import re
from core.config import PENDAFTARAN_DIR, ensure_dirs

# Path pangkalan data SQLite kekal untuk rekod MPPS & Reject Analysis
MPPS_DB_PATH = os.path.join(PENDAFTARAN_DIR, "radiologi_local.db")

# 14 Kategori Standard Analisis Penolakan Imej mengikut PHRIS (PER.SS-RA 101)
STANDARD_REJECT_CATEGORIES = [
    "OVER EXPOSURE",
    "UNDER EXPOSURE",
    "DOUBLE EXPOSURE",
    "WRONG TECHNIQUE",
    "WRONG PATIENT",
    "WRONG MARKER",
    "COLLIMATION ERROR",
    "PATIENT MOVEMENT",
    "PATIENT ARTIFACT",
    "EQUIPMENT FAULT",
    "DETECTOR FAULT",
    "IMAGE ARTIFACT",
    "PROCESSING FAULT",
    "MISCELLANEOUS"
]

def get_db_connection():
    """Membuka sambungan SQLite ke pangkalan data tempatan dengan row_factory."""
    ensure_dirs()
    conn = sqlite3.connect(MPPS_DB_PATH, timeout=15)
    conn.row_factory = sqlite3.Row
    return conn

def init_mpps_db():
    """Inisialisasi jadual mpps_records dan mpps_rejected_images jika belum wujud."""
    conn = get_db_connection()
    cur = conn.cursor()
    try:
        # 1. Jadual Utama MPPS
        cur.execute("""
        CREATE TABLE IF NOT EXISTS mpps_records (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            sop_instance_uid TEXT UNIQUE,
            sop_class_uid TEXT,
            accession_number TEXT,
            patient_id TEXT,
            patient_name TEXT,
            study_instance_uid TEXT,
            performed_procedure_step_id TEXT,
            procedure_description TEXT,
            modality TEXT,
            station_ae TEXT,
            station_name TEXT,
            status TEXT DEFAULT 'IN PROGRESS',
            start_date TEXT,
            start_time TEXT,
            end_date TEXT,
            end_time TEXT,
            status_reason TEXT,
            comments TEXT,
            total_images_count INTEGER DEFAULT 0,
            raw_dataset_json TEXT,
            created_at TEXT DEFAULT (datetime('now','localtime')),
            updated_at TEXT DEFAULT (datetime('now','localtime'))
        );
        """)

        # 2. Jadual Analisis Penolakan Imej (Reject Analysis)
        cur.execute("""
        CREATE TABLE IF NOT EXISTS mpps_rejected_images (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            mpps_record_id INTEGER REFERENCES mpps_records(id) ON DELETE CASCADE,
            sop_instance_uid TEXT,
            sop_class_uid TEXT,
            reject_reason TEXT,
            standard_category TEXT,
            rejected_by TEXT,
            image_count INTEGER DEFAULT 1,
            comments TEXT,
            created_at TEXT DEFAULT (datetime('now','localtime'))
        );
        """)

        # Indeks carian pantas
        cur.execute("CREATE INDEX IF NOT EXISTS idx_mpps_sop_instance ON mpps_records (sop_instance_uid);")
        cur.execute("CREATE INDEX IF NOT EXISTS idx_mpps_accession ON mpps_records (accession_number);")
        cur.execute("CREATE INDEX IF NOT EXISTS idx_mpps_patient_id ON mpps_records (patient_id);")
        cur.execute("CREATE INDEX IF NOT EXISTS idx_mpps_start_date ON mpps_records (start_date);")
        cur.execute("CREATE INDEX IF NOT EXISTS idx_mpps_status ON mpps_records (status);")
        cur.execute("CREATE INDEX IF NOT EXISTS idx_mpps_rej_record ON mpps_rejected_images (mpps_record_id);")
        cur.execute("CREATE INDEX IF NOT EXISTS idx_mpps_rej_cat ON mpps_rejected_images (standard_category);")

        conn.commit()
    except Exception as e:
        print(f"[MPPS DB Error] init_mpps_db: {e}")
    finally:
        cur.close()
        conn.close()

def match_reject_category(raw_reason):
    """
    Memetakan sebarang teks sebab penolakan, ulasan atau kod ke dalam salah satu
    daripada 14 Kategori Standard PHRIS.
    """
    if not raw_reason:
        return "MISCELLANEOUS"
    
    text = str(raw_reason).strip().upper()

    # 1. Exact match dahulu
    for cat in STANDARD_REJECT_CATEGORIES:
        if cat == text or cat in text:
            return cat

    # 2. Regex / Keyword matching (Prioritaskan kategori spesifik dahulu)
    if any(k in text for k in ["COLLIMAT", "KOLIMASI", "FIELD SIZE", "COLLIMATION"]):
        return "COLLIMATION ERROR"
    elif any(k in text for k in ["OVER", "DARK", "GELAP", "HIGH KV", "HIGH MAS", "TERLEBIH DOS", "OVEREXPOSURE"]):
        return "OVER EXPOSURE"
    elif any(k in text for k in ["UNDER", "LIGHT", "CERAH", "LOW KV", "LOW MAS", "KURANG DOS", "UNDEREXPOSURE"]):
        return "UNDER EXPOSURE"
    elif any(k in text for k in ["DOUBLE", "TINDIH", "BERTINDIH", "DOUBLE EXPOSURE"]):
        return "DOUBLE EXPOSURE"
    elif any(k in text for k in ["WRONG PATIENT", "SALAH PESAKIT", "WRONG PAT", "PATIENT MISMATCH"]):
        return "WRONG PATIENT"
    elif any(k in text for k in ["MARKER", "PENANDA", "WRONG MARKER", "NO MARKER", "MARKER SALAH", "TIADA MARKER"]):
        return "WRONG MARKER"
    elif any(k in text for k in ["MOVE", "MOTION", "GERAK", "PERGERAKAN", "PATIENT MOVEMENT", "BREATHING"]):
        return "PATIENT MOVEMENT"
    elif any(k in text for k in ["PATIENT ARTIFACT", "ARTEFAK PESAKIT", "BENDA ASING", "JEWELLERY", "NECKLACE", "RANTAI", "BARANG KEMAS", "ZIPPER", "BRA", "BUTTON", "BUTANG", "COIN", "PIN", "FOREIGN BODY"]):
        return "PATIENT ARTIFACT"
    elif any(k in text for k in ["EQUIPMENT", "ALAT", "XRAY TUBE", "GENERATOR", "TUBE FAULT", "MACHINE FAULT"]):
        return "EQUIPMENT FAULT"
    elif any(k in text for k in ["DETECTOR", "DETEKTOR", "DR PANEL", "PANEL FAULT", "CR PLATE", "CASSETTE FAULT", "PLATE FAULT"]):
        return "DETECTOR FAULT"
    elif any(k in text for k in ["IMAGE ARTIFACT", "ARTEFAK IMEJ", "GRID LINE", "SCRATCH", "CALIBRATION", "DEAD PIXEL", "GHOST"]):
        return "IMAGE ARTIFACT"
    elif any(k in text for k in ["PROCESSING", "PEMPROSESAN", "CR READER", "SOFTWARE FAULT", "ALGORITHM"]):
        return "PROCESSING FAULT"
    elif any(k in text for k in ["TECHNIQUE", "TEKNIK", "POSITION", "POSISI", "KEDUDUKAN", "CENTERING", "ANGLE", "CUT OFF", "GRID CUT"]):
        return "WRONG TECHNIQUE"
    else:
        return "MISCELLANEOUS"

def dataset_to_serializable_dict(ds):
    """Menukar pydicom Dataset kepada format dictionary JSON-serializable untuk audit."""
    if ds is None:
        return {}
    out = {}
    try:
        for elem in ds:
            tag_str = f"({elem.tag.group:04X},{elem.tag.element:04X})"
            key_name = elem.name or elem.keyword or tag_str
            if elem.VR == 'SQ':
                seq_list = []
                for sub_ds in elem.value:
                    seq_list.append(dataset_to_serializable_dict(sub_ds))
                out[key_name] = seq_list
            elif elem.VR in ['OB', 'OW', 'UN'] and len(elem.value) > 128:
                out[key_name] = f"<binary data {len(elem.value)} bytes>"
            else:
                val = elem.value
                if hasattr(val, 'original_string'):
                    val = str(val)
                elif isinstance(val, (bytes, bytearray)):
                    try:
                        val = val.decode('utf-8', errors='ignore')
                    except Exception:
                        val = str(val)
                else:
                    val = str(val) if val is not None else ""
                out[key_name] = val
    except Exception as e:
        out["_serialization_error"] = str(e)
    return out

def save_mpps_create_record(dataset, sop_instance_uid, sop_class_uid=None, station_ae_fallback=""):
    """
    Menyimpan rekod baru dari event N-CREATE (Status: IN PROGRESS).
    Memulangkan (record_id: int, message: str)
    """
    init_mpps_db()
    conn = get_db_connection()
    cur = conn.cursor()

    try:
        now_dt = datetime.datetime.now()
        today_str = now_dt.strftime("%Y%m%d")
        now_time_str = now_dt.strftime("%H%M%S")

        # Ekstraksi medan asas dari dataset
        patient_name = str(getattr(dataset, "PatientName", "") or "").strip().replace("^", " ")
        patient_id = str(getattr(dataset, "PatientID", "") or "").strip()
        accession_no = str(getattr(dataset, "AccessionNumber", "") or "").strip()
        study_uid = str(getattr(dataset, "StudyInstanceUID", "") or "").strip()
        proc_step_id = str(getattr(dataset, "PerformedProcedureStepID", "") or "").strip()
        proc_desc = str(getattr(dataset, "PerformedProcedureStepDescription", "") or "").strip()
        modality = str(getattr(dataset, "Modality", "CR") or "CR").strip().upper()
        station_ae = str(getattr(dataset, "PerformedStationAETitle", "") or station_ae_fallback or "").strip().upper()
        station_name = str(getattr(dataset, "PerformedStationName", "") or "").strip()
        start_date = str(getattr(dataset, "PerformedProcedureStepStartDate", "") or today_str).strip()
        start_time = str(getattr(dataset, "PerformedProcedureStepStartTime", "") or now_time_str).strip()
        status = str(getattr(dataset, "PerformedProcedureStepStatus", "IN PROGRESS") or "IN PROGRESS").strip().upper()

        # Semak ScheduledStepAttributesSequence jika wujud untuk melengkapkan Accession / Procedure
        if hasattr(dataset, "ScheduledStepAttributesSequence") and len(dataset.ScheduledStepAttributesSequence) > 0:
            sps = dataset.ScheduledStepAttributesSequence[0]
            if not accession_no:
                accession_no = str(getattr(sps, "AccessionNumber", "") or "").strip()
            if not study_uid:
                study_uid = str(getattr(sps, "StudyInstanceUID", "") or "").strip()
            if not proc_desc:
                proc_desc = str(getattr(sps, "RequestedProcedureDescription", "") or getattr(sps, "ScheduledProcedureStepDescription", "") or "").strip()

        from core.dicom_engine import dedup_laterality
        proc_desc = dedup_laterality(proc_desc)

        sop_inst = str(sop_instance_uid or getattr(dataset, "SOPInstanceUID", getattr(dataset, "AffectedSOPInstanceUID", "")) or "").strip()
        sop_cls = str(sop_class_uid or getattr(dataset, "SOPClassUID", getattr(dataset, "AffectedSOPClassUID", "1.2.840.10008.3.1.2.3.3")) or "").strip()

        raw_dict = dataset_to_serializable_dict(dataset)
        raw_dict["_mpps_event"] = "N-CREATE"
        raw_json = json.dumps(raw_dict, ensure_ascii=False)

        now_str = now_dt.strftime("%Y-%m-%d %H:%M:%S")

        # Masukkan atau ganti rekod
        cur.execute("""
        INSERT INTO mpps_records (
            sop_instance_uid, sop_class_uid, accession_number, patient_id, patient_name,
            study_instance_uid, performed_procedure_step_id, procedure_description,
            modality, station_ae, station_name, status, start_date, start_time,
            raw_dataset_json, created_at, updated_at
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        ON CONFLICT(sop_instance_uid) DO UPDATE SET
            status = excluded.status,
            accession_number = COALESCE(NULLIF(excluded.accession_number, ''), mpps_records.accession_number),
            patient_id = COALESCE(NULLIF(excluded.patient_id, ''), mpps_records.patient_id),
            patient_name = COALESCE(NULLIF(excluded.patient_name, ''), mpps_records.patient_name),
            procedure_description = COALESCE(NULLIF(excluded.procedure_description, ''), mpps_records.procedure_description),
            raw_dataset_json = excluded.raw_dataset_json,
            updated_at = excluded.updated_at;
        """, (
            sop_inst, sop_cls, accession_no, patient_id, patient_name,
            study_uid, proc_step_id, proc_desc, modality, station_ae, station_name,
            status, start_date, start_time, raw_json, now_str, now_str
        ))
        conn.commit()

        rec_id = cur.lastrowid
        print(f"[MPPS Engine] N-CREATE recorded: ID={rec_id}, SOP={sop_inst}, Patient={patient_name}, Acc={accession_no}, Status={status}")
        return rec_id, "N-CREATE successfully registered."
    except Exception as e:
        print(f"[MPPS Engine ERROR] save_mpps_create_record: {e}")
        return None, str(e)
    finally:
        cur.close()
        conn.close()

def save_mpps_set_record(dataset, sop_instance_uid):
    """
    Mengemaskini rekod MPPS dari event N-SET (Status: COMPLETED / DISCONTINUED).
    Turut merekodkan Reject Analysis sekiranya imej dibatalkan / ditolak.
    Memulangkan (success: bool, message: str)
    """
    init_mpps_db()
    conn = get_db_connection()
    cur = conn.cursor()

    try:
        sop_inst = str(sop_instance_uid or "").strip()
        if not sop_inst:
            return False, "Empty SOP Instance UID"

        cur.execute("SELECT * FROM mpps_records WHERE sop_instance_uid = ?", (sop_inst,))
        row = cur.fetchone()
        if not row:
            # Jika belum ada N-CREATE (sesetengah konsol hantar N-SET terus), cipta rekod baru
            print(f"[MPPS Engine Warning] Record not found for N-SET SOP={sop_inst}. Creating baseline record...")
            save_mpps_create_record(dataset, sop_inst)
            cur.execute("SELECT * FROM mpps_records WHERE sop_instance_uid = ?", (sop_inst,))
            row = cur.fetchone()

        now_dt = datetime.datetime.now()
        today_str = now_dt.strftime("%Y%m%d")
        now_time_str = now_dt.strftime("%H%M%S")
        now_str = now_dt.strftime("%Y-%m-%d %H:%M:%S")

        end_date = str(getattr(dataset, "PerformedProcedureStepEndDate", "") or today_str).strip()
        end_time = str(getattr(dataset, "PerformedProcedureStepEndTime", "") or now_time_str).strip()
        status = str(getattr(dataset, "PerformedProcedureStepStatus", row["status"] if row else "COMPLETED") or "COMPLETED").strip().upper()
        status_reason = str(getattr(dataset, "PerformedProcedureStepStatusReason", "") or "").strip()
        comments = str(getattr(dataset, "CommentsOnThePerformedProcedureStep", "") or "").strip()

        # Kira jumlah imej daripada PerformedSeriesSequence
        image_count = 0
        operator_name = ""
        series_sop_class = ""
        if hasattr(dataset, "PerformedSeriesSequence"):
            for series in dataset.PerformedSeriesSequence:
                if not operator_name and hasattr(series, "OperatorName"):
                    operator_name = str(series.OperatorName)
                if hasattr(series, "ReferencedImageSequence"):
                    image_count += len(series.ReferencedImageSequence)
                    if len(series.ReferencedImageSequence) > 0 and not series_sop_class:
                        ref_img = series.ReferencedImageSequence[0]
                        series_sop_class = str(getattr(ref_img, "ReferencedSOPClassUID", ""))

        if image_count == 0:
            # Fallback ke bilangan intances dalam dataset jika ada
            image_count = int(getattr(dataset, "NumberOfStudyRelatedInstances", 0) or 1 if status == "COMPLETED" else 0)

        # Gabungkan raw dataset JSON
        existing_raw = {}
        if row and row["raw_dataset_json"]:
            try:
                existing_raw = json.loads(row["raw_dataset_json"])
            except Exception:
                pass
        
        new_raw = dataset_to_serializable_dict(dataset)
        existing_raw["_mpps_n_set"] = new_raw
        raw_json_str = json.dumps(existing_raw, ensure_ascii=False)

        # Kemaskini rekod mpps_records
        cur.execute("""
        UPDATE mpps_records
        SET status = ?, end_date = ?, end_time = ?, status_reason = ?, comments = ?,
            total_images_count = ?, raw_dataset_json = ?, updated_at = ?
        WHERE sop_instance_uid = ?;
        """, (status, end_date, end_time, status_reason, comments, image_count, raw_json_str, now_str, sop_inst))
        conn.commit()

        # Semak dan proses Analisis Penolakan (Reject Analysis)
        # Jika status = DISCONTINUED atau terdapat ulasan penolakan / alasan pembatalan
        is_discontinued = "DISCONTINUED" in status
        has_reject_comment = any(k in comments.upper() for k in ["REJECT", "REPEAT", "TOLAK", "ULANG", "ERROR", "OVER", "UNDER", "MOVEMENT"])
        
        if is_discontinued or has_reject_comment or status_reason:
            raw_reason_text = f"{status_reason} {comments}".strip() or "DISCONTINUED BY OPERATOR"
            std_category = match_reject_category(raw_reason_text)
            rej_operator = operator_name or (row["station_ae"] if row else "CONSOLE")

            # Bilangan imej ditolak (minimum 1)
            rej_count = max(1, image_count if is_discontinued else 1)

            cur.execute("""
            INSERT INTO mpps_rejected_images (
                mpps_record_id, sop_instance_uid, sop_class_uid, reject_reason,
                standard_category, rejected_by, image_count, comments, created_at
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?);
            """, (
                row["id"] if row else None,
                sop_inst,
                series_sop_class or (row["sop_class_uid"] if row else "1.2.840.10008.3.1.2.3.3"),
                raw_reason_text,
                std_category,
                rej_operator,
                rej_count,
                comments,
                now_str
            ))
            conn.commit()
            print(f"[MPPS Reject Logged] Recorded reject image for SOP={sop_inst}: Category={std_category} (Reason: {raw_reason_text})")

        print(f"[MPPS Engine] N-SET updated: SOP={sop_inst}, Status={status}, Images={image_count}")
        return True, "N-SET successfully updated."
    except Exception as e:
        print(f"[MPPS Engine ERROR] save_mpps_set_record: {e}")
        return False, str(e)
    finally:
        cur.close()
        conn.close()

def get_mpps_monthly_reject_summary(year):
    """
    Mengagregat data Analisis Penolakan Imej (Reject Analysis) dan Jumlah Imej Diambil
    daripada jadual MPPS untuk 12 bulan pada tahun yang dipilih.
    Memulangkan format dictionary selari dengan seksyen 7 PHRIS.
    """
    init_mpps_db()
    conn = get_db_connection()
    cur = conn.cursor()

    summary = {
        "penolakan": {r: [0]*12 for r in STANDARD_REJECT_CATEGORIES},
        "total_images": [0]*12,
        "total_repeats": [0]*12
    }

    try:
        year_str = str(year)

        # 1. Kira Jumlah Imej Diambil (a) daripada mpps_records
        cur.execute("""
        SELECT 
            strftime('%m', created_at) as month_num,
            SUM(CASE WHEN total_images_count > 0 THEN total_images_count ELSE 1 END) as total_imgs
        FROM mpps_records
        WHERE (strftime('%Y', created_at) = ? OR start_date LIKE ?)
          AND status IN ('COMPLETED', 'DISCONTINUED', 'IN PROGRESS')
        GROUP BY strftime('%m', created_at);
        """, (year_str, f"{year_str}%"))

        for r in cur.fetchall():
            try:
                m_idx = int(r["month_num"]) - 1
                if 0 <= m_idx < 12:
                    summary["total_images"][m_idx] += int(r["total_imgs"] or 0)
            except Exception:
                pass

        # 2. Kira Pecahan Imej Ditolak mengikut 14 Kategori Standard PHRIS
        cur.execute("""
        SELECT 
            strftime('%m', created_at) as month_num,
            standard_category,
            SUM(image_count) as total_rej
        FROM mpps_rejected_images
        WHERE strftime('%Y', created_at) = ?
        GROUP BY strftime('%m', created_at), standard_category;
        """, (year_str,))

        for r in cur.fetchall():
            try:
                m_idx = int(r["month_num"]) - 1
                cat = str(r["standard_category"]).strip().upper()
                count = int(r["total_rej"] or 0)

                if 0 <= m_idx < 12:
                    if cat in summary["penolakan"]:
                        summary["penolakan"][cat][m_idx] += count
                    else:
                        summary["penolakan"]["MISCELLANEOUS"][m_idx] += count
                    summary["total_repeats"][m_idx] += count
            except Exception:
                pass

    except Exception as e:
        print(f"[MPPS Engine ERROR] get_mpps_monthly_reject_summary: {e}")
    finally:
        cur.close()
        conn.close()

    return summary

def get_mpps_records_list(limit=50, offset=0, status=None, query=None, start_date=None, end_date=None):
    """Mendapatkan senarai rekod MPPS untuk paparan UI dengan penapis dan pagination."""
    init_mpps_db()
    conn = get_db_connection()
    cur = conn.cursor()

    conditions = []
    params = []

    if status and status != "ALL":
        conditions.append("m.status = ?")
        params.append(status)

    if start_date:
        conditions.append("(m.start_date >= ? OR date(m.created_at) >= ?)")
        params.extend([start_date.replace("-", ""), start_date])

    if end_date:
        conditions.append("(m.start_date <= ? OR date(m.created_at) <= ?)")
        params.extend([end_date.replace("-", ""), end_date])

    if query:
        q_wild = f"%{query.strip()}%"
        conditions.append("(m.patient_name LIKE ? OR m.patient_id LIKE ? OR m.accession_number LIKE ? OR m.procedure_description LIKE ?)")
        params.extend([q_wild, q_wild, q_wild, q_wild])

    where_clause = f"WHERE {' AND '.join(conditions)}" if conditions else ""

    # Dapatkan jumlah rekod
    cur.execute(f"SELECT COUNT(*) as cnt FROM mpps_records m {where_clause}", params)
    total_count = cur.fetchone()["cnt"]

    # Dapatkan senarai rekod berserta maklumat reject
    sql = f"""
    SELECT 
        m.*,
        (SELECT COUNT(*) FROM mpps_rejected_images r WHERE r.mpps_record_id = m.id) as reject_count,
        (SELECT GROUP_CONCAT(r.standard_category, ', ') FROM mpps_rejected_images r WHERE r.mpps_record_id = m.id) as reject_categories
    FROM mpps_records m
    {where_clause}
    ORDER BY m.id DESC
    LIMIT ? OFFSET ?;
    """
    params.extend([limit, offset])
    cur.execute(sql, params)
    rows = [dict(r) for r in cur.fetchall()]

    cur.close()
    conn.close()

    return {
        "total": total_count,
        "limit": limit,
        "offset": offset,
        "records": rows
    }

def get_mpps_record_details(sop_instance_uid):
    """Mendapatkan rekod MPPS terperinci termasuk data audit JSON & senarai imej ditolak."""
    init_mpps_db()
    conn = get_db_connection()
    cur = conn.cursor()

    cur.execute("SELECT * FROM mpps_records WHERE sop_instance_uid = ? OR id = ?", (sop_instance_uid, sop_instance_uid))
    rec = cur.fetchone()
    if not rec:
        cur.close()
        conn.close()
        return None

    rec_dict = dict(rec)
    
    # Ambil maklumat imej ditolak
    cur.execute("SELECT * FROM mpps_rejected_images WHERE mpps_record_id = ? ORDER BY id ASC", (rec_dict["id"],))
    rec_dict["rejected_images"] = [dict(r) for r in cur.fetchall()]

    # Formatkan raw JSON
    if rec_dict.get("raw_dataset_json"):
        try:
            rec_dict["raw_dataset"] = json.loads(rec_dict["raw_dataset_json"])
        except Exception:
            rec_dict["raw_dataset"] = rec_dict["raw_dataset_json"]

    cur.close()
    conn.close()
    return rec_dict
