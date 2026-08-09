import os
import datetime
from core.config import load_config, MONTH_MAP, ALL_MODALITIES_CATALOG
from core.excel_engine import get_patients_list

def generate_export_data(year, month):
    """
    Menjana data penuh dan statistik agregasi bulanan untuk cetakan/PDF Buku Daftar KKM.
    Menapis rekod aktif dan mengecualikan rekod berstatus BATAL dari agregasi statistik.
    """
    config = load_config()
    klinik_asal = config.get("klinik_asal", "Klinik Kesihatan")
    singkatan = str(config.get("singkatan_klinik", "RAD")).upper()
    referral_clinics = config.get("klinik_rujukan", [])
    facility_type = str(config.get("facility_type", "KK")).upper()
    hospital_scope = str(config.get("hospital_scope", "ALL")).upper()
    single_modality_code = str(config.get("single_modality", "")).upper()
    
    modality_name = ""
    if facility_type == "HOSPITAL" and hospital_scope == "SINGLE" and single_modality_code:
        for m in ALL_MODALITIES_CATALOG:
            if m["code"].upper() == single_modality_code:
                modality_name = m["name"]
                break
        if not modality_name:
            modality_name = single_modality_code
    
    all_records = get_patients_list(year, month)
    
    # Rekod aktif untuk statistik beban kerja
    active_records = [r for r in all_records if not r.get("is_cancelled", False)]
    
    # Bulan nama
    folder_name, month_name = MONTH_MAP.get(month, ("01_JAN", "JAN"))
    
    # Pengiraan Ringkasan Audit (Audit Summary Breakdown)
    total_cases = len(active_records)
    male_count = sum(1 for r in active_records if r["jantina"].upper() in ["M", "LELAKI", "L"])
    female_count = sum(1 for r in active_records if r["jantina"].upper() in ["F", "PEREMPUAN", "P"])
    
    cxr_count = sum(1 for r in active_records if r["bahagian_pemeriksaan"].upper() in ["CXR", "DADA"])
    cxr_tb_count = sum(1 for r in active_records if "TB" in r["bahagian_pemeriksaan"].upper())
    cxr_hrg_count = sum(1 for r in active_records if "HRG" in r["bahagian_pemeriksaan"].upper())
    cxr_rme_count = sum(1 for r in active_records if "RME" in r["bahagian_pemeriksaan"].upper())
    axr_count = sum(1 for r in active_records if r["jenis_pemeriksaan"].upper() in ["ABDOMEN", "AXR"])
    ext_count = sum(1 for r in active_records if r["jenis_pemeriksaan"].upper() in ["EXTREMITI", "EXT"])
    
    # Klinik Rujukan breakdown
    clinic_counts = {}
    for c in referral_clinics:
        clinic_counts[c] = sum(1 for r in active_records if str(r["klinik_rujukan"]).upper() == str(c).upper())
    if klinik_asal:
        clinic_counts[klinik_asal] = sum(1 for r in active_records if str(r["klinik_rujukan"]).upper() == str(klinik_asal).upper())
    
    return {
        "year": year,
        "month": month,
        "month_name": month_name,
        "klinik_asal": klinik_asal,
        "singkatan": singkatan,
        "facility_type": facility_type,
        "hospital_scope": hospital_scope,
        "modality_name": modality_name,
        "records": all_records,
        "active_records": active_records,
        "total_cases": total_cases,
        "cancelled_count": len(all_records) - total_cases,
        "male_count": male_count,
        "female_count": female_count,
        "cxr_count": cxr_count,
        "cxr_tb_count": cxr_tb_count,
        "cxr_hrg_count": cxr_hrg_count,
        "cxr_rme_count": cxr_rme_count,
        "axr_count": axr_count,
        "ext_count": ext_count,
        "clinic_counts": clinic_counts
    }
