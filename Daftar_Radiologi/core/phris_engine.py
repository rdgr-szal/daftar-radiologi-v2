import os
import datetime
from core.config import load_config, load_extra_phris_data, save_extra_phris_data
from core.excel_engine import get_patients_list

def get_phris_matrix_data(year, month=None):
    """
    Menjana matriks reten tahunan PHRIS (PER.SS-RA 101) 12 bulan (JAN - DIS).
    Mengecualikan rekod BATAL daripada kiraan dos/kes dan menyokong fasiliti dinamik.
    """
    config = load_config()
    klinik_asal = str(config.get("klinik_asal", "")).strip().upper()
    
    races = ["MELAYU", "CINA", "INDIA", "BUMIPUTERA", "WARGA ASING", "LAIN-LAIN"]
    bangsa = {r: [0]*12 for r in races}
    
    arrival_keys = ["TROLLEY", "WHEELCHAIR", "RUJUK TERUS", "KLINIK_OPD"]
    kedatangan = {k: [0]*12 for k in arrival_keys}
    
    exam_am_keys = ["DADA", "ABDOMEN", "EXTREMITI", "RANGKA KEPALA", "SPINA VERTEBRA", "PELVIS", "SKELETAL SURVEY", "DEXA", "OPG", "LAIN-LAIN"]
    pemeriksaan_am = {k: [0]*12 for k in exam_am_keys}
    
    exam_other_keys = ["RME", "PTB", "KES LAIN", "COVID", "DM"]
    pemeriksaan_lain = {k: [0]*12 for k in exam_other_keys}
    
    appointment_keys = ["USG", "MAMMO", "CT-SCAN", "LAIN-LAIN"]
    temujanji = {k: [0]*12 for k in appointment_keys}
    
    consumable_keys = ["CD-R", "FILEM 10X12", "FILEM 14X17"]
    consumable = {k: [0]*12 for k in consumable_keys}
    
    rejections_list = [
        "OVER EXPOSURE", "UNDER EXPOSURE", "DOUBLE EXPOSURE", "WRONG TECHNIQUE",
        "WRONG PATIENT", "WRONG MARKER", "COLLIMATION ERROR", "PATIENT MOVEMENT",
        "PATIENT ARTIFACT", "EQUIPMENT FAULT", "DETECTOR FAULT", "IMAGE ARTIFACT",
        "PROCESSING FAULT", "MISCELLANEOUS"
    ]
    penolakan = {r: [0]*12 for r in rejections_list}
    penolakan["JUMLAH IMEJ"] = [0]*12
    penolakan["PENGULANGAN"] = [0]*12

    for month_idx in range(12):
        month_num = month_idx + 1
        all_records = get_patients_list(year, month_num)
        
        # Tapis rekod aktif sahaja untuk reten
        records = [r for r in all_records if not r.get("is_cancelled", False)]
        
        for r in records:
            b_val = str(r.get("bangsa", "")).upper()
            w_val = str(r.get("warganegara", "")).upper()
            kat_val = str(r.get("kategori", "")).upper()
            cat_val = str(r.get("catatan", "")).upper()
            jenis_val = str(r.get("jenis_pemeriksaan", "")).upper()
            bahagian_val = str(r.get("bahagian_pemeriksaan", "")).upper()
            klinik_val = str(r.get("klinik_rujukan", "")).upper()
            cd_val = str(r.get("cd_filem", "")).upper()
            
            # 1. Bangsa
            if w_val not in ["YA", "YES", "Y"] or "ASING" in b_val:
                bangsa["WARGA ASING"][month_idx] += 1
            elif "MELAYU" in b_val:
                bangsa["MELAYU"][month_idx] += 1
            elif "CINA" in b_val:
                bangsa["CINA"][month_idx] += 1
            elif "INDIA" in b_val:
                bangsa["INDIA"][month_idx] += 1
            elif "BUMI" in b_val or "SABAH" in b_val or "SARAWAK" in b_val:
                bangsa["BUMIPUTERA"][month_idx] += 1
            else:
                bangsa["LAIN-LAIN"][month_idx] += 1

            # 2. Kedatangan
            if "TROLLEY" in kat_val or "TROLLEY" in cat_val:
                kedatangan["TROLLEY"][month_idx] += 1
            elif "WHEELCHAIR" in kat_val or "KERUSI RODA" in cat_val or "WHEELCHAIR" in cat_val:
                kedatangan["WHEELCHAIR"][month_idx] += 1
            elif klinik_val and klinik_asal and klinik_val != klinik_asal:
                kedatangan["RUJUK TERUS"][month_idx] += 1
            else:
                kedatangan["KLINIK_OPD"][month_idx] += 1

            # 3. Pemeriksaan AM
            if "DADA" in jenis_val or "CXR" in bahagian_val or "CHEST" in jenis_val:
                pemeriksaan_am["DADA"][month_idx] += 1
            elif "ABDOMEN" in jenis_val or "AXR" in bahagian_val or "KUB" in bahagian_val:
                pemeriksaan_am["ABDOMEN"][month_idx] += 1
            elif "EXTREMITI" in jenis_val or any(k in bahagian_val for k in ["FOOT", "ANKLE", "KNEE", "HAND", "WRIST", "ELBOW", "SHOULDER"]):
                pemeriksaan_am["EXTREMITI"][month_idx] += 1
            elif "RANGKA KEPALA" in jenis_val or "SKULL" in bahagian_val or "FACE" in bahagian_val:
                pemeriksaan_am["RANGKA KEPALA"][month_idx] += 1
            elif "SPINA" in jenis_val or any(k in bahagian_val for k in ["LUMBOSACRAL", "CERVICAL", "THORACIC"]):
                pemeriksaan_am["SPINA VERTEBRA"][month_idx] += 1
            elif "PELVIS" in jenis_val or "PELVIS" in bahagian_val:
                pemeriksaan_am["PELVIS"][month_idx] += 1
            elif "SKELETAL" in jenis_val or "SKELETAL" in bahagian_val:
                pemeriksaan_am["SKELETAL SURVEY"][month_idx] += 1
            elif "DEXA" in jenis_val or "DEXA" in bahagian_val:
                pemeriksaan_am["DEXA"][month_idx] += 1
            elif "OPG" in jenis_val or "OPG" in bahagian_val:
                pemeriksaan_am["OPG"][month_idx] += 1
            else:
                pemeriksaan_am["LAIN-LAIN"][month_idx] += 1

            # 4. Pemeriksaan Lain
            if "RME" in bahagian_val or "RME" in cat_val:
                pemeriksaan_lain["RME"][month_idx] += 1
            if "PTB" in bahagian_val or "TB" in bahagian_val or "PTB" in cat_val:
                pemeriksaan_lain["PTB"][month_idx] += 1
            if "COVID" in cat_val or "COVID" in bahagian_val:
                pemeriksaan_lain["COVID"][month_idx] += 1
            if "DM" in cat_val or "DIABETES" in cat_val:
                pemeriksaan_lain["DM"][month_idx] += 1

            # 5. Temujanji
            if "USG" in jenis_val or "ULTRASOUND" in cat_val or "USG" in cat_val:
                temujanji["USG"][month_idx] += 1
            elif "MAMMO" in jenis_val or "MAMMO" in cat_val:
                temujanji["MAMMO"][month_idx] += 1
            elif "CT" in jenis_val or "CT-SCAN" in cat_val:
                temujanji["CT-SCAN"][month_idx] += 1

            # 6. Consumables
            if "CD" in cd_val:
                consumable["CD-R"][month_idx] += 1
            if "10X12" in cd_val:
                consumable["FILEM 10X12"][month_idx] += 1
            if "14X17" in cd_val:
                consumable["FILEM 14X17"][month_idx] += 1

            # 7. Penolakan Imej
            tot_exp = 1
            try:
                tot_exp = int(r.get("total_expose", 1) or 1)
            except (ValueError, TypeError):
                tot_exp = 1

            tot_rej = 0
            try:
                tot_rej = int(r.get("total_reject", 0) or 0)
            except (ValueError, TypeError):
                tot_rej = 0

            penolakan["JUMLAH IMEJ"][month_idx] += tot_exp
            penolakan["PENGULANGAN"][month_idx] += tot_rej
            
            if tot_rej > 0:
                matched_rej = False
                for r_key in rejections_list:
                    if r_key in cat_val:
                        penolakan[r_key][month_idx] += tot_rej
                        matched_rej = True
                        break
                if not matched_rej:
                    penolakan["MISCELLANEOUS"][month_idx] += tot_rej

    return {
        "year": year,
        "bangsa": bangsa,
        "kedatangan": kedatangan,
        "pemeriksaan_am": pemeriksaan_am,
        "pemeriksaan_lain": pemeriksaan_lain,
        "temujanji": temujanji,
        "consumable": consumable,
        "penolakan": penolakan
    }
