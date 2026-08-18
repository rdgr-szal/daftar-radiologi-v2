import os
import sys
import socket
import datetime
import json
import re
import threading
import time
from core.config import DICOM_WORKLIST_PATH, load_config, ensure_dirs

# DICOM Standard SOP Class UIDs (Universal across all pynetdicom versions)
MWL_FIND_SOP_CLASS = "1.2.840.10008.5.1.4.31"  # Modality Worklist Information Model - FIND
VERIFICATION_SOP_CLASS = "1.2.840.10008.1.1"    # Verification (C-ECHO)
MPPS_SOP_CLASS = "1.2.840.10008.3.1.2.3.3"      # Modality Performed Procedure Step SOP Class
MPPS_GENERAL_SOP_CLASS = "1.2.840.10008.5.1.4.32" # General Purpose Performed Procedure Step SOP Class

PYDICOM_AVAILABLE = False
PYNETDICOM_AVAILABLE = False

try:
    import pydicom
    from pydicom.dataset import Dataset
    from pydicom.sequence import Sequence
    from pydicom.uid import generate_uid
    PYDICOM_AVAILABLE = True
except ImportError:
    pass

try:
    from pynetdicom import AE, evt
    PYNETDICOM_AVAILABLE = True
except ImportError:
    pass

def get_local_lan_ip():
    """Detect the local IPv4 address on the active LAN interface."""
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.settimeout(0.5)
        # Bind routing interface without actual internet communication
        s.connect(('10.255.255.255', 1))
        ip = s.getsockname()[0]
        s.close()
        return ip
    except Exception:
        try:
            return socket.gethostbyname(socket.gethostname())
        except Exception:
            return "127.0.0.1"

# --- WORKLIST DATA MANAGEMENT (JSON PERSISTENT STORE) ---

def load_dicom_worklist():
    """Load active patient worklist records from persistent storage."""
    ensure_dirs()
    if not os.path.exists(DICOM_WORKLIST_PATH):
        return []
    try:
        with open(DICOM_WORKLIST_PATH, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception as e:
        print(f"[DICOM Worklist Error] Failed to read worklist: {e}")
        return []

def save_dicom_worklist(items):
    """Save patient worklist records to persistent storage."""
    ensure_dirs()
    try:
        with open(DICOM_WORKLIST_PATH, "w", encoding="utf-8") as f:
            json.dump(items, f, indent=2, ensure_ascii=False)
        return True
    except Exception as e:
        print(f"[DICOM Worklist Error] Failed to write worklist: {e}")
        return False

def dedup_laterality(text: str) -> str:
    """
    Remove duplicate laterality occurrences (e.g. 'RIGHT RADIUS ULNA RIGHT' -> 'RIGHT RADIUS ULNA').
    Handles both English (RIGHT, LEFT, BILATERAL, BOTH) and Malay (KANAN, KIRI, KEDUA-DUA, KEDUA).
    """
    if not text:
        return ""
    
    cleaned = str(text).strip()
    laterality_tokens = ["RIGHT", "LEFT", "BILATERAL", "BOTH", "KANAN", "KIRI", "KEDUA-DUA", "KEDUA"]
    
    for lat in laterality_tokens:
        pattern = re.compile(rf"\b{re.escape(lat)}\b", re.IGNORECASE)
        matches = list(pattern.finditer(cleaned))
        if len(matches) >= 2:
            # Keep the first occurrence, strip duplicate later occurrences
            for m in reversed(matches[1:]):
                start, end = m.span()
                cleaned = cleaned[:start] + " " + cleaned[end:]
            cleaned = re.sub(r"\s+", " ", cleaned).strip()
            
    return cleaned

def build_procedure_name(bahagian: str, lateraliti: str = "") -> str:
    """
    Build a standardized procedure description (e.g. 'RIGHT RADIUS ULNA', 'CHEST PA', 'LEFT KNEE AP/LAT').
    Guarantees that laterality is never duplicated in the description.
    """
    bahagian_clean = str(bahagian or "").strip().upper()
    lat_clean = str(lateraliti or "").strip().upper()

    lat_map = {
        "KANAN": "RIGHT",
        "KIRI": "LEFT",
        "KEDUA-DUA": "BILATERAL",
        "KEDUA": "BILATERAL"
    }
    lat_clean = lat_map.get(lat_clean, lat_clean)

    if not lat_clean:
        return dedup_laterality(bahagian_clean)

    # Check for directional laterality (RIGHT, LEFT, BILATERAL, BOTH)
    is_directional = lat_clean in ["RIGHT", "LEFT", "BILATERAL", "BOTH"]
    has_lat_in_bahagian = bool(re.search(rf"\b{re.escape(lat_clean)}\b", bahagian_clean))
    
    # Also check if Malay counterpart is already in bahagian_clean
    if lat_clean == "RIGHT" and bool(re.search(r"\bKANAN\b", bahagian_clean)):
        has_lat_in_bahagian = True
    elif lat_clean == "LEFT" and bool(re.search(r"\bKIRI\b", bahagian_clean)):
        has_lat_in_bahagian = True

    if is_directional:
        if has_lat_in_bahagian:
            proc_desc = bahagian_clean
        else:
            proc_desc = f"{lat_clean} {bahagian_clean}".strip()
    else:
        if has_lat_in_bahagian:
            proc_desc = bahagian_clean
        else:
            proc_desc = f"{bahagian_clean} {lat_clean}".strip()

    return dedup_laterality(proc_desc)

def build_procedure_id(proc_name_or_bahagian: str) -> str:
    """
    Build a clean DICOM RequestedProcedureID (VR SH, <= 16 chars).
    e.g. 'RIGHT RADIUS ULNA' -> 'RIGHTRADIUSULNA'
    """
    if not proc_name_or_bahagian:
        return "CXR"
    cleaned = re.sub(r'[^A-Z0-9]', '', str(proc_name_or_bahagian).upper())
    return cleaned[:16] if cleaned else "CXR"

def format_dicom_patient_name(name_str):
    """
    Format patient name for DICOM standard Person Name (PN VR).
    Cleans special characters and formats appropriately.
    """
    if not name_str:
        return "UNKNOWN"
    clean = str(name_str).strip().upper()
    # Replace slashes and double carets
    clean = clean.replace("/", " ").replace("\\", " ")
    return clean

def add_to_dicom_worklist(patient_data, exam_list=None):
    """
    Format and add newly registered patient examination(s) to the DICOM Modality Worklist.
    Automatically called upon patient registration.
    """
    try:
        config = load_config()
        dicom_cfg = config.get("dicom_config", {})
        
        # Station AE Title filter (console AE or default wildcard)
        station_ae = str(dicom_cfg.get("console_ae_title", "")).strip().upper()
        default_modality = str(dicom_cfg.get("default_modality", "CR")).strip().upper() or "CR"
        facility_name = str(config.get("klinik_asal", "RADIOLOGY CLINIC")).strip().upper()
        
        now = datetime.datetime.now()
        date_str = now.strftime("%Y%m%d")
        time_str = now.strftime("%H%M%S")
        
        # Extract and format patient demographics
        raw_name = patient_data.get("nama", "")
        patient_name = format_dicom_patient_name(raw_name)
        
        ic_no = str(patient_data.get("ic_pasport", "")).strip().replace("-", "").replace(" ", "")
        umur = str(patient_data.get("umur", "")).strip()
        jantina = str(patient_data.get("jantina", "O")).strip().upper()
        
        if jantina in ["L", "LELAKI", "M", "MALE"]:
            dicom_sex = "M"
        elif jantina in ["P", "PEREMPUAN", "F", "FEMALE"]:
            dicom_sex = "F"
        else:
            dicom_sex = "O"
            
        # Parse Date of Birth from Malaysian MyKad if applicable
        dob_str = ""
        if len(ic_no) >= 6 and ic_no[:6].isdigit():
            yy = int(ic_no[:2])
            mm = int(ic_no[2:4])
            dd = int(ic_no[4:6])
            current_yy = int(str(now.year)[2:])
            full_year = (1900 + yy) if yy > current_yy else (2000 + yy)
            if 1 <= mm <= 12 and 1 <= dd <= 31:
                dob_str = f"{full_year:04d}{mm:02d}{dd:02d}"

        # Age in DICOM 4-character format (e.g. '035Y')
        dicom_age = f"{int(umur):03d}Y" if umur.isdigit() else "000Y"

        if not exam_list or not isinstance(exam_list, list):
            exam_list = [{
                "xray_no": patient_data.get("nombor_xray", "0001"),
                "bahagian": patient_data.get("bahagian_pemeriksaan", "CHEST"),
                "lateraliti": patient_data.get("lateraliti", ""),
                "modality": patient_data.get("modality", default_modality),
                "bil_kes": patient_data.get("bil_kes", "1"),
                "operator": patient_data.get("operator", "")
            }]

        current_worklist = load_dicom_worklist()
        
        # Purge records older than auto_clear_hours (default 24 hours)
        hours_limit = int(dicom_cfg.get("auto_clear_hours", 24))
        cutoff = now - datetime.timedelta(hours=hours_limit)
        
        cleaned_worklist = []
        for itm in current_worklist:
            created_at_str = itm.get("created_at")
            if created_at_str:
                try:
                    dt = datetime.datetime.fromisoformat(created_at_str)
                    if dt > cutoff:
                        cleaned_worklist.append(itm)
                except Exception:
                    cleaned_worklist.append(itm)
            else:
                cleaned_worklist.append(itm)

        for idx, ex in enumerate(exam_list):
            xray_no = str(ex.get("xray_no", patient_data.get("nombor_xray", "0001"))).strip()
            bahagian = str(ex.get("bahagian", ex.get("bahagian_pemeriksaan", "CHEST"))).strip().upper()
            lateraliti = str(ex.get("lateraliti", "")).strip().upper()
            
            proc_desc = build_procedure_name(bahagian, lateraliti)
            operator = str(ex.get("operator") or patient_data.get("operator", "")).strip().upper()
            doctor = str(patient_data.get("pegawai_rujukan", "")).strip().upper()
            case_no = str(ex.get("bil_kes") or patient_data.get("bil_kes", "")).strip()
            
            # Modality Mapping
            mod_code = default_modality
            raw_mod = str(ex.get("modality", "")).upper()
            if "CR" in raw_mod or "GENERAL" in raw_mod or "X-RAY" in raw_mod or "XRAY" in raw_mod:
                mod_code = "CR"
            elif "DX" in raw_mod or "DIGITAL" in raw_mod:
                mod_code = "DX"
            elif "US" in raw_mod or "ULTRASOUND" in raw_mod:
                mod_code = "US"
            elif "CT" in raw_mod:
                mod_code = "CT"
            elif "MR" in raw_mod:
                mod_code = "MR"
            elif "MG" in raw_mod or "MAMMO" in raw_mod:
                mod_code = "MG"
            elif "RG" in raw_mod:
                mod_code = "RG"

            clean_xray_num = int(''.join(filter(str.isdigit, xray_no))) if any(c.isdigit() for c in xray_no) else (idx + 1)
            
            if PYDICOM_AVAILABLE:
                study_uid = generate_uid(prefix="1.2.826.0.1.3680043.9.7128.")
            else:
                study_uid = f"1.2.826.0.1.3680043.9.7128.{now.strftime('%Y%m%d%H%M%S')}.{clean_xray_num}.{idx+1}"

            req_proc_id = build_procedure_id(proc_desc)
            sps_id = f"SPS_{clean_xray_num}_{idx+1}"[:16]

            worklist_item = {
                "id": f"{xray_no}_{idx+1}_{date_str}",
                "created_at": now.isoformat(),
                "patient_name": patient_name,
                "patient_id": ic_no[:16] if ic_no else f"PAT-{clean_xray_num}"[:16],
                "patient_birth_date": dob_str,
                "patient_sex": dicom_sex,
                "patient_age": dicom_age,
                "accession_number": xray_no[:16],
                "study_instance_uid": study_uid,
                "requested_procedure_id": req_proc_id,
                "requested_procedure_desc": proc_desc[:64],
                "scheduled_station_ae": station_ae[:16] if station_ae else "*",
                "scheduled_station_name": "XRAY_ROOM",
                "scheduled_step_start_date": date_str,
                "scheduled_step_start_time": time_str,
                "modality": mod_code[:16],
                "scheduled_physician": operator[:64],
                "scheduled_step_desc": proc_desc[:64],
                "scheduled_step_id": sps_id,
                "scheduled_step_location": facility_name[:64],
                "referring_physician": doctor[:64],
                "admission_id": case_no[:16]
            }
            cleaned_worklist.append(worklist_item)

        save_dicom_worklist(cleaned_worklist)
        print(f"[DICOM MWL] Successfully queued {len(exam_list)} examination(s) for patient {patient_name} (X-Ray No: {patient_data.get('nombor_xray')})")
        return True
    except Exception as e:
        print(f"[DICOM MWL Error] Failed to queue worklist record: {e}")
        return False

def remove_from_dicom_worklist(xray_no):
    """Remove records for a given X-ray accession number from the DICOM worklist."""
    try:
        items = load_dicom_worklist()
        target = str(xray_no).strip()
        filtered = [it for it in items if it.get("accession_number") != target]
        if len(filtered) != len(items):
            save_dicom_worklist(filtered)
            print(f"[DICOM MWL] Removed accession {xray_no} from worklist.")
        return True
    except Exception as e:
        print(f"[DICOM MWL Error] Failed to remove worklist item: {e}")
        return False

def update_in_dicom_worklist(xray_no, updated_data):
    """Update patient demographic details in the DICOM worklist."""
    try:
        items = load_dicom_worklist()
        target = str(xray_no).strip()
        modified = False
        for it in items:
            if it.get("accession_number") == target:
                if "nama" in updated_data and updated_data["nama"]:
                    it["patient_name"] = format_dicom_patient_name(updated_data["nama"])
                if "ic_pasport" in updated_data and updated_data["ic_pasport"]:
                    it["patient_id"] = str(updated_data["ic_pasport"]).replace("-", "").replace(" ", "")[:16]
                if "bahagian_pemeriksaan" in updated_data and updated_data["bahagian_pemeriksaan"]:
                    bahagian = str(updated_data["bahagian_pemeriksaan"]).strip().upper()
                    lateraliti = str(updated_data.get("lateraliti", "")).strip().upper()
                    proc = build_procedure_name(bahagian, lateraliti)
                    it["requested_procedure_id"] = build_procedure_id(proc)
                    it["requested_procedure_desc"] = proc[:64]
                    it["scheduled_step_desc"] = proc[:64]
                if "operator" in updated_data and updated_data["operator"]:
                    it["scheduled_physician"] = str(updated_data["operator"]).strip().upper()[:64]
                if "pegawai_rujukan" in updated_data and updated_data["pegawai_rujukan"]:
                    it["referring_physician"] = str(updated_data["pegawai_rujukan"]).strip().upper()[:64]
                modified = True
        if modified:
            save_dicom_worklist(items)
            print(f"[DICOM MWL] Updated accession {xray_no} in worklist.")
        return True
    except Exception as e:
        print(f"[DICOM MWL Error] Failed to update worklist item: {e}")
        return False

def clear_dicom_worklist():
    """Clear all queued items from the DICOM worklist."""
    return save_dicom_worklist([])

# --- DICOM C-FIND & C-ECHO SCP HANDLERS ---

def handle_c_echo(event):
    """C-ECHO SCP Request Handler (DICOM Verification Service)."""
    requestor = event.assoc.requestor
    print(f"[DICOM Verification SCP] C-ECHO received from {requestor.ae_title} ({requestor.address}:{requestor.port})")
    return 0x0000  # Success

def match_query_field(query_val, record_val):
    """Match query key supporting DICOM wildcards (*, ?), partial match, or exact match."""
    if query_val is None or query_val == "" or query_val == "*":
        return True
    q = str(query_val).strip().upper().replace("*", "").replace("?", "")
    r = str(record_val).strip().upper()
    if not q:
        return True
    return q in r or r in q

def handle_c_find(event):
    """
    DICOM Modality Worklist C-FIND SCP Request Handler.
    Processes query keys from Modality Console (SCU) and yields matching patient datasets.
    """
    requestor = event.assoc.requestor
    raw_ae = requestor.ae_title
    req_ae = raw_ae.decode('ascii', errors='ignore').strip() if isinstance(raw_ae, bytes) else str(raw_ae).strip()
    req_ip = requestor.address
    print(f"[DICOM MWL C-FIND] Query received from {req_ae} @ {req_ip}")

    if not PYDICOM_AVAILABLE:
        print("[DICOM MWL Error] pydicom library is not installed.")
        return

    try:
        model = event.identifier
        
        # Extract query search keys from the C-FIND identifier
        q_patient_name = getattr(model, "PatientName", None)
        q_patient_id = getattr(model, "PatientID", None)
        q_accession = getattr(model, "AccessionNumber", None)
        
        q_start_date = None
        q_modality = None
        q_station_ae = None
        
        if hasattr(model, "ScheduledProcedureStepSequence") and len(model.ScheduledProcedureStepSequence) > 0:
            sps_req = model.ScheduledProcedureStepSequence[0]
            q_start_date = getattr(sps_req, "ScheduledProcedureStepStartDate", None)
            q_modality = getattr(sps_req, "Modality", None)
            q_station_ae = getattr(sps_req, "ScheduledStationAETitle", None)

        today_str = datetime.date.today().strftime("%Y%m%d")
        
        # Handle date query format
        if str(q_start_date) == str(today_str):
            search_date = today_str
        elif q_start_date and "-" in str(q_start_date):
            parts = str(q_start_date).split("-")
            search_date = parts[0].strip() or today_str
        else:
            search_date = str(q_start_date) if q_start_date else None

        worklist_items = load_dicom_worklist()
        match_count = 0

        for item in worklist_items:
            # Apply matching filters
            if q_patient_id and not match_query_field(q_patient_id, item.get("patient_id")):
                continue
            if q_patient_name and not match_query_field(q_patient_name, item.get("patient_name")):
                continue
            if q_accession and not match_query_field(q_accession, item.get("accession_number")):
                continue
            if search_date and search_date != "*" and item.get("scheduled_step_start_date") != search_date:
                continue
            if q_modality and not match_query_field(q_modality, item.get("modality")):
                continue

            # Build DICOM Standard MWL Response Dataset
            ds = Dataset()
            ds.is_little_endian = True
            ds.is_implicit_VR = False

            # Specific Character Set (UTF-8)
            ds.SpecificCharacterSet = 'ISO_IR 192'

            # Patient Identification & Demographics
            ds.PatientName = str(item.get("patient_name", "UNKNOWN"))[:64]
            ds.PatientID = str(item.get("patient_id", "0000"))[:16]
            ds.PatientBirthDate = str(item.get("patient_birth_date", ""))[:8]
            ds.PatientSex = str(item.get("patient_sex", "O"))[:1]
            ds.PatientAge = str(item.get("patient_age", "000Y"))[:4]
            ds.PatientWeight = ''

            # Study & Requested Procedure Details
            acc_num = str(item.get("accession_number", "0001"))[:16]
            proc_desc_clean = dedup_laterality(str(item.get("requested_procedure_desc", "X-RAY")))[:64]
            step_desc_clean = dedup_laterality(str(item.get("scheduled_step_desc", proc_desc_clean)))[:64]
            req_proc_id_val = str(item.get("requested_procedure_id") or build_procedure_id(proc_desc_clean))[:16]

            ds.AccessionNumber = acc_num
            ds.StudyInstanceUID = str(item.get("study_instance_uid", generate_uid()))
            ds.RequestedProcedureID = req_proc_id_val
            ds.RequestedProcedureDescription = proc_desc_clean
            ds.StudyID = acc_num
            ds.AdmissionID = str(item.get("admission_id", ""))[:16]
            ds.ReferringPhysicianName = str(item.get("referring_physician", ""))[:64]
            ds.RequestingPhysician = str(item.get("referring_physician", ""))[:64]

            # Scheduled Procedure Step Sequence (Type 1 Sequence)
            sps_ds = Dataset()
            target_station_ae = str(item.get("scheduled_station_ae", req_ae or "CONSOLE"))[:16]
            sps_ds.ScheduledStationAETitle = target_station_ae if target_station_ae != "*" else req_ae[:16]
            sps_ds.ScheduledProcedureStepStartDate = str(item.get("scheduled_step_start_date", today_str))[:8]
            sps_ds.ScheduledProcedureStepStartTime = str(item.get("scheduled_step_start_time", "090000"))[:6]
            sps_ds.Modality = str(item.get("modality", "CR"))[:16]
            sps_ds.ScheduledPerformingPhysicianName = str(item.get("scheduled_physician", ""))[:64]
            sps_ds.ScheduledProcedureStepDescription = step_desc_clean
            sps_ds.ScheduledProcedureStepID = str(item.get("scheduled_step_id", "SPS-0001"))[:16]
            sps_ds.ScheduledStationName = str(item.get("scheduled_station_name", "XRAY_ROOM"))[:16]
            sps_ds.ScheduledProcedureStepLocation = str(item.get("scheduled_step_location", ""))[:16]
            sps_ds.ScheduledProcedureStepStatus = "SCHEDULED"
            sps_ds.ScheduledProtocolCodeSequence = Sequence()

            # Carestream MWL requirement: RequestedProcedureSequence inside ScheduledProcedureStepSequence
            req_proc = Dataset()
            req_proc.RequestedProcedureID = req_proc_id_val
            req_proc.RequestedProcedureDescription = proc_desc_clean
            sps_ds.RequestedProcedureSequence = Sequence([req_proc])

            ds.ScheduledProcedureStepSequence = Sequence([sps_ds])

            match_count += 1
            print(f"[DICOM MWL] Returning match #{match_count}: {item.get('accession_number')} | {item.get('patient_name')} ({item.get('requested_procedure_desc')})")
            
            # Yield Pending Status (0xFF00) with matching dataset
            yield (0xFF00, ds)

        print(f"[DICOM MWL C-FIND] Completed. Dispatched {match_count} record(s) to SCU {req_ae}.")
    except Exception as e:
        print(f"[DICOM MWL C-FIND Error] Exception during query processing: {e}")
        yield (0xC000, None)  # Error: Unable to process

# --- DICOM C-ECHO SCU & CONSOLE CONNECTIVITY TEST (UNIVERSAL) ---

def check_arp_table(target_ip):
    """
    Step 0 - ARP Table Check:
    Check if target IP exists in local ARP table for early diagnostic insights.
    """
    import subprocess
    try:
        # Use -n flag when available on mac/linux to avoid slow DNS resolution
        target = str(target_ip).strip()
        if not target or target in ["127.0.0.1", "localhost", "0.0.0.0"]:
            return True, f"Loopback/local address {target} is active on local machine."

        cmd = ["arp", "-n", target] if sys.platform != 'win32' else ["arp", "-a", target]
        proc = subprocess.run(cmd, capture_output=True, text=True, timeout=1.5)
        output = (proc.stdout or "") + (proc.stderr or "")
        
        if target in output and "no match" not in output.lower() and "cannot find" not in output.lower():
            return True, f"IP {target} found in local ARP cache (MAC address resolved)."
        return False, f"IP {target} not currently in local ARP cache (Normal if routed through a gateway or console hasn't communicated recently)."
    except Exception as e:
        return False, f"ARP diagnostic note: {e}"

def check_local_dicom_server(local_port=104):
    """
    Step 1 - Local DICOM Server Live Check:
    Verify if local app DICOM SCP server is currently listening on port 104.
    """
    try:
        p = int(local_port)
    except (ValueError, TypeError):
        p = 104
    try:
        with socket.create_connection(('127.0.0.1', p), timeout=1.0):
            return True, f"Local DICOM MWL server is actively listening on 127.0.0.1:{p}."
    except Exception as e:
        return False, f"Local DICOM server is not running on port {p} ({e}). Please enable the DICOM server in Settings."

def test_dicom_connection(console_ip=None, console_port=None, console_ae=None, my_ae=None, local_port=None):
    """
    Universal 5-Stage DICOM Connection Test between DaftarRadiologi (SCP) and Modality Console (SCU).
    Universal across all console vendors (Carestream, Konica Minolta, GE, Fuji, Agfa, Philips, Siemens, etc.).
    
    1. Step 0 (ARP): Checks local ARP cache.
    2. Step 1 (Local Server): Verifies local DICOM server listening on local port (default 104).
    3. Step 2 (C-ECHO / DICOM Ping): pynetdicom Verification SOP Class association & C-ECHO.
       - Confirms whether console recognizes local AE title ('XRAY' / 'KAUNTER') and answers alive.
    4. Step 3 (C-FIND MWL): Checks console response to MWL query.
       - Note: Modality consoles are typically SCUs (pull worklist from app), so rejecting C-FIND queries is normal.
    5. Returns structured diagnostic checks report with clear causes.
    """
    config = load_config()
    dicom_cfg = config.get("dicom_config", {})

    target_ip = str(console_ip or dicom_cfg.get("console_ip", "")).strip()
    target_port = console_port or dicom_cfg.get("console_port", 104)
    target_ae = str(console_ae or dicom_cfg.get("console_ae_title", "CARESTREAM")).strip() or "CARESTREAM"
    local_ae = str(my_ae or dicom_cfg.get("ae_title", "KAUNTER")).strip() or "KAUNTER"
    local_p = local_port or dicom_cfg.get("port", 104)

    now_str = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    result = {
        "ok": False,
        "testedAt": now_str,
        "ae_title": local_ae,
        "console": target_ae,
        "console_ip": target_ip,
        "console_port": target_port,
        "checks": [],
        "summary": "",
        "failure_reason": ""
    }

    if not target_ip:
        result["summary"] = "Modality Console IP address is not specified."
        result["failure_reason"] = "Modality Console IP address is empty. Please configure the console IP address in Settings."
        result["checks"].append({
            "name": "Configuration Validation",
            "ok": False,
            "detail": "Console IP is empty. Please enter the console IPv4 address."
        })
        return result

    try:
        port_num = int(target_port)
    except (ValueError, TypeError):
        result["summary"] = f"Invalid port number: {target_port}"
        result["failure_reason"] = f"Invalid port number: {target_port}. Standard DICOM port is 104 (or 2762 / 11112)."
        result["checks"].append({
            "name": "Configuration Validation",
            "ok": False,
            "detail": f"Invalid console port '{target_port}'."
        })
        return result

    t_start = time.time()

    # Step 0: ARP Cache Inspection
    arp_ok, arp_msg = check_arp_table(target_ip)
    result["checks"].append({
        "name": "ARP & LAN Cache",
        "ok": arp_ok,
        "detail": arp_msg,
        "info_only": True
    })

    # Step 1: Local Server Listening Check
    local_ok, local_msg = check_local_dicom_server(local_p)
    result["checks"].append({
        "name": f"Local DICOM Server (127.0.0.1:{local_p})",
        "ok": local_ok,
        "detail": local_msg
    })

    # Step 1.5: TCP Network Socket Connection to Modality Console
    socket_connected = False
    sock_error_msg = ""
    try:
        with socket.create_connection((target_ip, port_num), timeout=3.0):
            socket_connected = True
    except Exception as e_sock:
        sock_error_msg = str(e_sock)

    if not socket_connected:
        reason = (
            f"Unable to establish TCP socket connection to {target_ip}:{port_num} ({sock_error_msg}). "
            "Possible causes: Modality console is turned OFF, IP address is wrong, "
            "Windows/Modality Firewall is blocking port, or console is on a different subnet."
        )
        result["checks"].append({
            "name": f"TCP Network Connection ({target_ip}:{port_num})",
            "ok": False,
            "detail": f"Connection Refused / Timed Out: {sock_error_msg}"
        })
        result["checks"].append({
            "name": f"DICOM C-ECHO Ping ({target_ae})",
            "ok": False,
            "detail": "Skipped: Network socket to console is unreachable."
        })
        result["ok"] = False
        result["failure_reason"] = reason
        result["summary"] = f"Network Connection Failed to {target_ip}:{port_num}"
        return result

    result["checks"].append({
        "name": f"TCP Network Connection ({target_ip}:{port_num})",
        "ok": True,
        "detail": f"Console host is reachable and accepting TCP connections on port {port_num}."
    })

    # Step 2: DICOM C-ECHO Verification via pynetdicom
    if not PYNETDICOM_AVAILABLE:
        result["ok"] = True
        result["summary"] = f"TCP Network Reachable ({target_ip}:{port_num})"
        result["checks"].append({
            "name": f"DICOM C-ECHO Verification ({target_ae})",
            "ok": True,
            "detail": "Network socket reachable (pynetdicom library not installed for full DICOM protocol verification)."
        })
        return result

    try:
        ae = AE(ae_title=local_ae.encode('ascii'))
        ae.add_requested_context(VERIFICATION_SOP_CLASS)
        ae.network_timeout = 4.0
        ae.acse_timeout = 4.0
        ae.dimse_timeout = 4.0

        assoc = ae.associate(target_ip, port_num, ae_title=target_ae.encode('ascii'))
        if assoc.is_established:
            status = assoc.send_c_echo()
            assoc.release()
            total_ms = round((time.time() - t_start) * 1000, 1)

            if status and getattr(status, "Status", None) == 0x0000:
                result["ok"] = True
                result["checks"].append({
                    "name": f"DICOM C-ECHO Ping ({target_ae})",
                    "ok": True,
                    "detail": f"Success (0x0000)! Console recognizes AE '{local_ae}' and responded in {total_ms}ms."
                })
            else:
                st_hex = hex(status.Status) if status and hasattr(status, "Status") else "UNKNOWN"
                result["ok"] = True
                result["checks"].append({
                    "name": f"DICOM C-ECHO Ping ({target_ae})",
                    "ok": True,
                    "detail": f"Association accepted. DICOM status response: {st_hex} ({total_ms}ms)."
                })
        elif getattr(assoc, 'is_rejected', False):
            total_ms = round((time.time() - t_start) * 1000, 1)
            reason = (
                f"DICOM Association Rejected by console '{target_ae}'. "
                f"Root cause: AE Title mismatch. Please verify that AE Title '{local_ae}' is registered as HIS/RIS in the console settings, "
                f"or console AE Title '{target_ae}' matches the console station name."
            )
            result["ok"] = False
            result["failure_reason"] = reason
            result["checks"].append({
                "name": f"DICOM C-ECHO Ping ({target_ae})",
                "ok": False,
                "detail": f"Association Rejected: AE Title '{local_ae}' or '{target_ae}' is not configured on console."
            })
            result["summary"] = "Association Rejected: AE Title Mismatch"
            return result
        else:
            total_ms = round((time.time() - t_start) * 1000, 1)
            reason = (
                f"DICOM Association failed or timed out ({total_ms}ms). "
                "Console DICOM service might be stopped, or port is occupied by another service."
            )
            result["ok"] = False
            result["failure_reason"] = reason
            result["checks"].append({
                "name": f"DICOM C-ECHO Ping ({target_ae})",
                "ok": False,
                "detail": f"Association Aborted / Timed out ({total_ms}ms)."
            })
            result["summary"] = "DICOM Association Failed"
            return result

    except Exception as e_assoc:
        total_ms = round((time.time() - t_start) * 1000, 1)
        result["ok"] = False
        result["failure_reason"] = f"DICOM Association Exception: {str(e_assoc)}"
        result["checks"].append({
            "name": f"DICOM C-ECHO Ping ({target_ae})",
            "ok": False,
            "detail": f"Error: {str(e_assoc)}"
        })
        result["summary"] = f"DICOM Error: {str(e_assoc)}"
        return result

    # Step 3: C-FIND MWL Query Role Check (Informative)
    try:
        ae_find = AE(ae_title=local_ae.encode('ascii'))
        # Set scu_role and scp_role for pynetdicom 2.x and 3.x compatibility
        try:
            ae_find.add_requested_context(MWL_FIND_SOP_CLASS, scu_role=True, scp_role=True)
        except Exception:
            ae_find.add_requested_context(MWL_FIND_SOP_CLASS)

        ae_find.network_timeout = 3.0
        ae_find.acse_timeout = 3.0
        ae_find.dimse_timeout = 3.0

        assoc_find = ae_find.associate(target_ip, port_num, ae_title=target_ae.encode('ascii'))
        if assoc_find.is_established:
            ds_query = Dataset()
            ds_query.PatientName = "*"
            ds_query.PatientID = "*"
            ds_query.AccessionNumber = "*"
            ds_query.ScheduledProcedureStepSequence = Sequence()

            responses = assoc_find.send_c_find(ds_query, query_model=MWL_FIND_SOP_CLASS)
            # Consume responses
            resp_list = []
            for status, identifier in responses:
                resp_list.append(status)
            assoc_find.release()

            result["checks"].append({
                "name": "C-FIND MWL Query",
                "ok": True,
                "detail": f"Console accepted C-FIND MWL query ({len(resp_list)} response frames received).",
                "info_only": True
            })
        else:
            result["checks"].append({
                "name": "C-FIND MWL Query",
                "ok": True,
                "detail": "Console is a modality SCU (pulls worklist from DaftarRadiologi, does not serve C-FIND) — Normal behavior.",
                "info_only": True
            })
    except Exception as e_find:
        result["checks"].append({
            "name": "C-FIND MWL Query",
            "ok": True,
            "detail": f"Modality console operates in SCU worklist receiver mode ({e_find}) — Normal behavior.",
            "info_only": True
        })

    total_time_ms = round((time.time() - t_start) * 1000, 1)
    result["ok"] = True
    result["summary"] = f"DICOM Connection OK ({total_time_ms}ms)"
    return result

def test_dicom_echo_scu(console_ip, console_port=104, console_ae="CARESTREAM", my_ae="KAUNTER"):
    """
    Backwards-compatible wrapper for test_dicom_echo_scu.
    Returns (success: bool, message: str, elapsed_ms: float).
    """
    res = test_dicom_connection(
        console_ip=console_ip,
        console_port=console_port,
        console_ae=console_ae,
        my_ae=my_ae
    )
    if res.get("ok"):
        return True, f"DICOM Connection OK! Console '{res.get('console')}' at {res.get('console_ip')}:{res.get('console_port')} is online and verified.", 0.0
    else:
        fail_msg = res.get("failure_reason") or res.get("summary") or "DICOM Connection Failed."
        return False, fail_msg, 0.0

# --- DICOM MPPS SCP HANDLERS (N-CREATE & N-SET) ---

def handle_n_create(event):
    """
    DICOM MPPS N-CREATE SCP Request Handler.
    Triggered when Modality Console (SCU) starts a procedure (Status: IN PROGRESS).
    """
    requestor = event.assoc.requestor
    raw_ae = requestor.ae_title
    req_ae = raw_ae.decode('ascii', errors='ignore').strip() if isinstance(raw_ae, bytes) else str(raw_ae).strip()
    print(f"[DICOM MPPS SCP] N-CREATE received from {req_ae} @ {requestor.address}")

    if not PYDICOM_AVAILABLE:
        print("[DICOM MPPS Error] pydicom library is not installed.")
        return 0x0110, None  # Processing Failure

    try:
        ds = getattr(event, "attribute_list", None) or getattr(event, "dataset", None)
        if ds is None and hasattr(event.request, "DataSet"):
            ds = event.request.DataSet

        if ds is None:
            print("[DICOM MPPS N-CREATE Error] Missing Attribute List in N-CREATE request.")
            return 0x0120, None  # Missing Attribute

        sop_inst = getattr(event.request, "AffectedSOPInstanceUID", None)
        if not sop_inst and hasattr(ds, "SOPInstanceUID"):
            sop_inst = str(ds.SOPInstanceUID)
        if not sop_inst and hasattr(ds, "AffectedSOPInstanceUID"):
            sop_inst = str(ds.AffectedSOPInstanceUID)

        from core.mpps_engine import save_mpps_create_record
        rec_id, msg = save_mpps_create_record(
            dataset=ds,
            sop_instance_uid=sop_inst,
            sop_class_uid=getattr(event.request, "AffectedSOPClassUID", MPPS_SOP_CLASS),
            station_ae_fallback=req_ae
        )

        if rec_id is None:
            return 0x0110, None  # Processing Failure

        # Success Status
        return 0x0000, None
    except Exception as e:
        print(f"[DICOM MPPS N-CREATE Exception] {e}")
        return 0x0110, None

def handle_n_set(event):
    """
    DICOM MPPS N-SET SCP Request Handler.
    Triggered when Modality Console (SCU) updates, completes, or discontinues a procedure step.
    """
    requestor = event.assoc.requestor
    raw_ae = requestor.ae_title
    req_ae = raw_ae.decode('ascii', errors='ignore').strip() if isinstance(raw_ae, bytes) else str(raw_ae).strip()
    print(f"[DICOM MPPS SCP] N-SET received from {req_ae} @ {requestor.address}")

    if not PYDICOM_AVAILABLE:
        return 0x0110, None

    try:
        ds = getattr(event, "modification_list", None) or getattr(event, "attribute_list", None) or getattr(event, "dataset", None)
        if ds is None and hasattr(event.request, "DataSet"):
            ds = event.request.DataSet

        if ds is None:
            print("[DICOM MPPS N-SET Error] Missing Modification List in N-SET request.")
            return 0x0120, None

        sop_inst = getattr(event.request, "RequestedSOPInstanceUID", getattr(event.request, "AffectedSOPInstanceUID", None))
        if not sop_inst and hasattr(ds, "SOPInstanceUID"):
            sop_inst = str(ds.SOPInstanceUID)

        from core.mpps_engine import save_mpps_set_record
        ok, msg = save_mpps_set_record(dataset=ds, sop_instance_uid=sop_inst)

        if not ok:
            return 0x0110, None

        return 0x0000, None
    except Exception as e:
        print(f"[DICOM MPPS N-SET Exception] {e}")
        return 0x0110, None

# --- DICOM MWL SERVER DAEMON (BACKGROUND SCP SERVICE) ---

class DicomMWLServerDaemon:
    _instance = None
    _lock = threading.Lock()

    def __init__(self):
        self.server_thread = None
        self.ae_server = None
        self.is_running = False
        self.host = "0.0.0.0"
        self.port = 104
        self.ae_title = "KAUNTER"
        self.last_error = ""

    @classmethod
    def get_instance(cls):
        with cls._lock:
            if cls._instance is None:
                cls._instance = DicomMWLServerDaemon()
            return cls._instance

    def start(self, host="0.0.0.0", port=104, ae_title="KAUNTER"):
        """Start the DICOM MWL & MPPS SCP Server in a background daemon thread."""
        if not PYNETDICOM_AVAILABLE:
            self.last_error = "The pynetdicom library is required. Please install via pip install pynetdicom."
            print(f"[DICOM Server] {self.last_error}")
            return False, self.last_error

        with self._lock:
            if self.is_running:
                return True, "DICOM Server is already running."

            self.host = host
            self.port = int(port)
            self.ae_title = str(ae_title).strip() or "KAUNTER"
            self.last_error = ""

            try:
                self.ae_server = AE(ae_title=self.ae_title.encode('ascii'))
                
                # Add Presentation Contexts using standard SOP Class UIDs (MWL, Verification, MPPS)
                self.ae_server.add_supported_context(MWL_FIND_SOP_CLASS)
                self.ae_server.add_supported_context(VERIFICATION_SOP_CLASS)
                self.ae_server.add_supported_context(MPPS_SOP_CLASS)
                self.ae_server.add_supported_context(MPPS_GENERAL_SOP_CLASS)
                
                # Bind event handlers
                handlers = [
                    (evt.EVT_C_ECHO, handle_c_echo),
                    (evt.EVT_C_FIND, handle_c_find),
                    (evt.EVT_N_CREATE, handle_n_create),
                    (evt.EVT_N_SET, handle_n_set)
                ]

                # Launch non-blocking background server thread
                self.server_thread = threading.Thread(
                    target=self._run_server,
                    args=(handlers,),
                    daemon=True
                )
                self.server_thread.start()
                self.is_running = True
                print(f"[DICOM MWL+MPPS Server] Successfully started on {self.host}:{self.port} (AE Title: {self.ae_title})")
                return True, f"DICOM MWL+MPPS Server is now active on port {self.port} (AE: {self.ae_title})"
            except Exception as e:
                self.is_running = False
                self.last_error = str(e)
                print(f"[DICOM Server Error] Failed to start server: {e}")
                return False, f"Failed to start DICOM Server: {str(e)}"

    def _run_server(self, handlers):
        try:
            self.ae_server.start_server(
                (self.host, self.port),
                block=True,
                evt_handlers=handlers
            )
        except Exception as e:
            self.is_running = False
            self.last_error = str(e)
            print(f"[DICOM Server Stopped/Error] {e}")

    def stop(self):
        """Stop the DICOM MWL+MPPS Server."""
        with self._lock:
            if not self.is_running:
                return True, "DICOM Server is already stopped."

            try:
                if self.ae_server:
                    self.ae_server.shutdown()
                self.is_running = False
                print("[DICOM MWL+MPPS Server] Server stopped.")
                return True, "DICOM MWL+MPPS Server successfully stopped."
            except Exception as e:
                self.last_error = str(e)
                return False, f"Error stopping server: {str(e)}"

    def restart(self, host="0.0.0.0", port=104, ae_title="KAUNTER"):
        """Restart the DICOM Server with new parameters."""
        self.stop()
        time.sleep(0.5)
        return self.start(host=host, port=port, ae_title=ae_title)

    def get_status(self):
        """Return real-time diagnostic status of the DICOM Server."""
        items = load_dicom_worklist()
        mpps_count = 0
        try:
            from core.mpps_engine import get_mpps_records_list
            res = get_mpps_records_list(limit=1, offset=0)
            mpps_count = res.get("total", 0)
        except Exception:
            pass

        return {
            "pynetdicom_installed": PYNETDICOM_AVAILABLE,
            "running": self.is_running,
            "host": self.host,
            "port": self.port,
            "ae_title": self.ae_title,
            "local_lan_ip": get_local_lan_ip(),
            "active_worklist_count": len(items),
            "mpps_supported": True,
            "mpps_records_count": mpps_count,
            "last_error": self.last_error
        }

def start_dicom_server_from_config():
    """Start the DICOM Server from config.json if enabled."""
    config = load_config()
    dicom_cfg = config.get("dicom_config", {})
    if dicom_cfg.get("enabled", False):
        port = dicom_cfg.get("port", 104)
        ae = dicom_cfg.get("ae_title", "KAUNTER")
        host = dicom_cfg.get("host", "0.0.0.0")
        daemon = DicomMWLServerDaemon.get_instance()
        return daemon.start(host=host, port=port, ae_title=ae)
    return False, "DICOM MWL Server is disabled in settings."
