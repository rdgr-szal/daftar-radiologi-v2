import os
import datetime
from core.config import load_config, MONTH_MAP, ALL_MODALITIES_CATALOG
from core.excel_engine import get_patients_list

def generate_export_data(year=None, month=None, period="month", selected_date_str=None, start_date_str=None, end_date_str=None):
    """
    Menjana data penuh untuk cetakan/PDF Buku Daftar KKM berdasarkan julat tempoh yang dipilih.
    """
    today = datetime.date.today()
    if year is None:
        year = today.year
    if month is None:
        month = today.month

    config = load_config()
    klinik_asal = config.get("klinik_asal", "Klinik Kesihatan")
    singkatan = str(config.get("singkatan_klinik", "RAD")).upper()
    facility_type = str(config.get("facility_type", "KK")).upper()
    hospital_scope = str(config.get("hospital_scope", "ALL")).upper()
    single_modality_code = str(config.get("single_modality", "")).upper()
    
    modality_name = ""
    if facility_type == "HOSPITAL" and single_modality_code:
        for m in ALL_MODALITIES_CATALOG:
            if m["code"].upper() == single_modality_code:
                modality_name = m["name"]
                break
        if not modality_name:
            modality_name = single_modality_code
    
    all_records = get_patients_list(
        year=year,
        month=month,
        period=period,
        selected_date_str=selected_date_str,
        start_date_str=start_date_str,
        end_date_str=end_date_str
    )
    
    folder_name, month_name = MONTH_MAP.get(month, ("01_JAN", "JAN"))

    target_date = today
    if selected_date_str:
        try:
            target_date = datetime.datetime.strptime(selected_date_str, "%Y-%m-%d").date()
        except ValueError:
            target_date = today

    if period == "day":
        period_label = f"SEHARI ({target_date.strftime('%d/%m/%Y')})"
    elif period == "week":
        w_start = target_date - datetime.timedelta(days=target_date.weekday())
        w_end = w_start + datetime.timedelta(days=6)
        period_label = f"SEMINGGU ({w_start.strftime('%d/%m/%Y')} - {w_end.strftime('%d/%m/%Y')})"
    elif period == "year":
        period_label = f"SETAHUN ({year})"
    elif period == "custom" and start_date_str and end_date_str:
        try:
            s_d = datetime.datetime.strptime(start_date_str, "%Y-%m-%d").strftime("%d/%m/%Y")
            e_d = datetime.datetime.strptime(end_date_str, "%Y-%m-%d").strftime("%d/%m/%Y")
            period_label = f"JULAT TARIKH ({s_d} - {e_d})"
        except ValueError:
            period_label = f"JULAT TARIKH ({start_date_str} - {end_date_str})"
    else:  # month
        period_label = f"SEBULAN ({month_name} {year})"
    
    active_records = [r for r in all_records if not r.get("is_cancelled", False)]
    total_cases = len(active_records)
    
    return {
        "year": year,
        "month": month,
        "month_name": month_name,
        "period": period,
        "period_label": period_label,
        "klinik_asal": klinik_asal,
        "singkatan": singkatan,
        "facility_type": facility_type,
        "hospital_scope": hospital_scope,
        "modality_name": modality_name,
        "records": all_records,
        "active_records": active_records,
        "total_cases": total_cases,
        "cancelled_count": len(all_records) - total_cases
    }
