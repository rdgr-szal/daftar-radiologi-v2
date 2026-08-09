import datetime
from flask import Blueprint, render_template, request, jsonify
from core.config import load_config, MONTH_MAP
from core.excel_engine import get_patients_list

dashboard_bp = Blueprint('dashboard', __name__)

@dashboard_bp.route('/dashboard')
def dashboard_view():
    config = load_config()
    today = datetime.date.today()
    period = request.args.get('period', 'day')  # Default to 'day' as requested
    selected_date = request.args.get('day_date', str(today))
    
    try:
        t_date = datetime.datetime.strptime(selected_date, "%Y-%m-%d").date()
    except Exception:
        t_date = today
        selected_date = str(today)

    if period in ['day', 'week']:
        selected_year = t_date.year
        selected_month = t_date.month
    else:
        selected_year = int(request.args.get('year', today.year))
        selected_month = int(request.args.get('month', today.month))

    from core.excel_engine import get_available_years
    available_years = get_available_years()
    if selected_year not in available_years:
        available_years.append(selected_year)
        available_years = sorted(list(set(available_years)))

    current_week_num = t_date.isocalendar()[1]
    
    # Calculate week range text
    week_start = t_date - datetime.timedelta(days=t_date.weekday())
    week_end = week_start + datetime.timedelta(days=6)
    week_range_text = f"{week_start.strftime('%d/%m/%Y')} - {week_end.strftime('%d/%m/%Y')}"

    return render_template(
        'dashboard.html',
        config=config,
        selected_year=selected_year,
        selected_month=selected_month,
        period=period,
        selected_date=selected_date,
        week_range_text=week_range_text,
        current_week_num=current_week_num,
        month_map=MONTH_MAP,
        current_year=today.year,
        available_years=available_years,
        current_page='dashboard'
    )

@dashboard_bp.route('/api/dashboard-data')
def api_dashboard_data():
    today = datetime.date.today()
    period = request.args.get('period', 'day')
    selected_date = request.args.get('day_date', str(today))
    year = int(request.args.get('year', today.year))
    month = int(request.args.get('month', today.month))
    
    config = load_config()
    singkatan = str(config.get("singkatan_klinik", "")).strip()
    klinik_asal = str(config.get("klinik_asal", "")).strip()
    if not singkatan:
        singkatan = klinik_asal or "KK"

    patients = get_patients_list(
        year=year,
        month=month,
        period=period,
        selected_date_str=selected_date
    )
    
    # Filter out cancelled records for calculations
    active_records = [p for p in patients if not p.get("is_cancelled", False)]
    
    # Bil Kes (active records)
    total_cases = len(active_records)
    
    # Bil Pesakit (unique active patients based on date + (ic or name))
    active_patient_keys = set()
    for p in active_records:
        t_str = str(p.get("tarikh", "")).strip()
        i_str = str(p.get("ic_pasport", "")).strip().upper()
        n_str = str(p.get("nama", "")).strip().upper()
        pk = f"{t_str}_{i_str}" if (i_str and i_str != "-") else f"{t_str}_{n_str}"
        active_patient_keys.add(pk)
    total_patients = len(active_patient_keys)
    
    # Helper to check if a clinic is "fasiliti sendiri"
    def is_sendiri(cl_val):
        cl_upper = str(cl_val or "").strip().upper()
        if not cl_upper:
            return True
        return cl_upper == klinik_asal.upper() or cl_upper == singkatan.upper() or cl_upper == "FASILITI SENDIRI"

    # Section 1: Reten mengikut jenis pemeriksaan (Fasiliti Sendiri vs Fasiliti Luar)
    # Bilangan kes mengikut jenis pemeriksaan
    sendiri_modality_counts = {}
    luar_modality_counts = {}
    
    # Let's count by jenis_pemeriksaan (modality)
    for p in active_records:
        mod = str(p.get("jenis_pemeriksaan", "CHEST")).strip().upper()
        if not mod:
            mod = "CHEST"
        clinic = p.get("klinik_rujukan", "")
        if is_sendiri(clinic):
            sendiri_modality_counts[mod] = sendiri_modality_counts.get(mod, 0) + 1
        else:
            luar_modality_counts[mod] = luar_modality_counts.get(mod, 0) + 1

    # Section 2: Bilangan pesakit mengikut keadaan (kategori kedatangan)
    # pesakit jalan kaki (kategori = 'BERJALAN')
    jalan_kaki_sendiri = 0
    jalan_kaki_luar = 0
    
    # pesakit wheelchair & trolley split by specific facility
    wheelchair_sendiri = 0
    wheelchair_luar = 0
    trolley_sendiri = 0
    trolley_luar = 0
    
    wheelchair_by_clinic = {}
    trolley_by_clinic = {}
    
    # RME (Routine Medical Examination)
    rme_sendiri = 0
    rme_luar = 0
    
    for p in active_records:
        kategori_raw = str(p.get("kategori", "")).strip().upper()
        clinic = p.get("klinik_rujukan", "")
        facility_label = singkatan if is_sendiri(clinic) else (clinic.upper() if clinic else "LUAR")
        
        # Check jalan kaki
        if kategori_raw == "BERJALAN":
            if is_sendiri(clinic):
                jalan_kaki_sendiri += 1
            else:
                jalan_kaki_luar += 1
        # Check wheelchair & trolley
        elif "WHEELCHAIR" in kategori_raw or "KERUSI RODA" in kategori_raw:
            wheelchair_by_clinic[facility_label] = wheelchair_by_clinic.get(facility_label, 0) + 1
            if is_sendiri(clinic):
                wheelchair_sendiri += 1
            else:
                wheelchair_luar += 1
        elif "TROLI" in kategori_raw or "TROLLEY" in kategori_raw:
            trolley_by_clinic[facility_label] = trolley_by_clinic.get(facility_label, 0) + 1
            if is_sendiri(clinic):
                trolley_sendiri += 1
            else:
                trolley_luar += 1
            
        # Check RME
        bahagian_val = str(p.get("bahagian_pemeriksaan", "")).strip().upper()
        if "RME" in bahagian_val:
            if is_sendiri(clinic):
                rme_sendiri += 1
            else:
                rme_luar += 1
                
    wheelchair_count = sum(wheelchair_by_clinic.values())
    trolley_count = sum(trolley_by_clinic.values())

    # Section 3:
    # 1. kiraan jumlah pesakit keseluruhan & kes mengikut fasiliti rujukan
    clinic_patients_map = {}
    clinic_cases_map = {}
    for p in active_records:
        clinic = p.get("klinik_rujukan", "")
        t_str = str(p.get("tarikh", "")).strip()
        i_str = str(p.get("ic_pasport", "")).strip().upper()
        n_str = str(p.get("nama", "")).strip().upper()
        
        # Consistent key logic: date + identity key (ignoring '-' placeholder)
        pk = f"{t_str}_{i_str}" if (i_str and i_str != "-") else f"{t_str}_{n_str}"
        
        label = singkatan if is_sendiri(clinic) else (clinic.upper() if clinic else "LUAR")
        if label not in clinic_patients_map:
            clinic_patients_map[label] = set()
        clinic_patients_map[label].add(pk)
        
        clinic_cases_map[label] = clinic_cases_map.get(label, 0) + 1
        
    clinic_patient_counts = {lbl: len(pks) for lbl, pks in clinic_patients_map.items()}

    # 2. table kiraan kes fasiliti sendiri vs luar
    kes_sendiri = sum(1 for p in active_records if is_sendiri(p.get("klinik_rujukan", "")))
    kes_luar = total_cases - kes_sendiri
    
    # 4. table jantina
    male_count = sum(1 for p in active_records if str(p.get("jantina", "")).strip().upper() in ["M", "LELAKI", "L"])
    female_count = sum(1 for p in active_records if str(p.get("jantina", "")).strip().upper() in ["F", "PEREMPUAN", "P"])
    
    # 5. table bangsa
    bangsa_counts = {}
    for p in active_records:
        b = str(p.get("bangsa", "")).strip().upper()
        if not b:
            b = "LAIN-LAIN"
        bangsa_counts[b] = bangsa_counts.get(b, 0) + 1
        
    # 6. table consumables
    consumables_counts = {}
    for p in active_records:
        c = str(p.get("cd_filem", "")).strip().upper()
        if c and c != "-":
            consumables_counts[c] = consumables_counts.get(c, 0) + 1

    # Total modality counts for Section 1 (combined sendiri + luar)
    total_modality_counts = {}
    for mod in set(list(sendiri_modality_counts.keys()) + list(luar_modality_counts.keys())):
        total_modality_counts[mod] = sendiri_modality_counts.get(mod, 0) + luar_modality_counts.get(mod, 0)

    return jsonify({
        "success": True,
        "singkatan": singkatan,
        "summary": {
            "total_cases": total_cases,
            "total_patients": total_patients
        },
        "section1": {
            "sendiri": sendiri_modality_counts,
            "luar": luar_modality_counts,
            "total": total_modality_counts
        },
        "section2": {
            "jalan_kaki_sendiri": jalan_kaki_sendiri,
            "jalan_kaki_luar": jalan_kaki_luar,
            "wheelchair": wheelchair_count,
            "trolley": trolley_count,
            "rme_sendiri": rme_sendiri,
            "rme_luar": rme_luar
        },
        "section3": {
            "clinic_patients": clinic_patient_counts,
            "clinic_cases": clinic_cases_map,
            "kes_sendiri_vs_luar": {
                "sendiri": kes_sendiri,
                "luar": kes_luar
            },
            "mobilisasi_fasiliti": {
                "wheelchair_sendiri": wheelchair_sendiri,
                "wheelchair_luar": wheelchair_luar,
                "trolley_sendiri": trolley_sendiri,
                "trolley_luar": trolley_luar
            },
            "mobilisasi_by_clinic": {
                "wheelchair": wheelchair_by_clinic,
                "trolley": trolley_by_clinic
            },
            "jantina": {
                "lelaki": male_count,
                "perempuan": female_count
            },
            "bangsa": bangsa_counts,
            "consumables": consumables_counts
        }
    })
