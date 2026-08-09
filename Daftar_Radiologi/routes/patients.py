import datetime
from flask import Blueprint, render_template, request, jsonify
from core.config import load_config, MONTH_MAP
from core.excel_engine import (
    get_patients_list,
    get_grouped_patients_list,
    get_available_years,
    edit_patient_record,
    cancel_patient_record
)

patients_bp = Blueprint('patients', __name__)

@patients_bp.route('/patient-list')
@patients_bp.route('/patient-list', methods=['GET'])
def patient_directory():
    config = load_config()
    today = datetime.date.today()
    period = request.args.get('period', 'day')
    selected_date = request.args.get('day_date', str(today))
    search_q = request.args.get('q', '')

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

    grouped_patients, total_cases, total_patients, week_range_text = get_grouped_patients_list(
        year=selected_year,
        month=selected_month,
        search_query=search_q,
        period=period,
        selected_date_str=selected_date
    )

    available_years = get_available_years()
    if selected_year not in available_years:
        available_years.append(selected_year)
        available_years = sorted(list(set(available_years)))

    current_week_num = t_date.isocalendar()[1]
    
    return render_template(
        'patient_list.html',
        config=config,
        patients=grouped_patients,
        total_cases=total_cases,
        total_patients=total_patients,
        week_range_text=week_range_text,
        current_week_num=current_week_num,
        selected_year=selected_year,
        selected_month=selected_month,
        period=period,
        selected_date=selected_date,
        search_q=search_q,
        month_map=MONTH_MAP,
        current_year=today.year,
        available_years=available_years,
        current_page='patient_list'
    )

@patients_bp.route('/api/patients')
def api_get_patients():
    today = datetime.date.today()
    year = int(request.args.get('year', today.year))
    month = int(request.args.get('month', today.month))
    period = request.args.get('period', 'day')
    selected_date = request.args.get('day_date', str(today))
    search_q = request.args.get('q', '')
    
    grouped_patients, total_cases, total_patients, week_range_text = get_grouped_patients_list(
        year=year,
        month=month,
        search_query=search_q,
        period=period,
        selected_date_str=selected_date
    )
    return jsonify({
        "success": True,
        "total_cases": total_cases,
        "total_patients": total_patients,
        "week_range_text": week_range_text,
        "patients": grouped_patients
    })

@patients_bp.route('/api/patient/edit', methods=['POST'])
def api_edit_patient():
    """
    Mengemaskini butiran pesakit dan merefleksikannya secara automatik ke fail Excel dan Database.
    """
    data = request.get_json() or {}
    tarikh = data.get('tarikh')
    xray_no = data.get('nombor_xray')
    updated_data = data.get('updated_data', {})
    
    if not xray_no:
        return jsonify({"success": False, "message": "Nombor X-Ray diperlukan."}), 400
        
    success, msg = edit_patient_record(tarikh, xray_no, updated_data)
    if success:
        return jsonify({"success": True, "message": msg})
    else:
        return jsonify({"success": False, "message": msg}), 500

@patients_bp.route('/api/patient/cancel', methods=['POST'])
def api_cancel_patient():
    """
    Pematuhan Kriteria Audit KKM:
    Membatalkan pendaftaran pesakit tanpa memadam nombor X-ray, merefleksikan ke Excel & DB.
    """
    data = request.get_json() or {}
    tarikh = data.get('tarikh')
    xray_no = data.get('nombor_xray')
    reason = data.get('reason', 'Pesakit Tidak Hadir / Batal')
    staff_name = data.get('staff_name', '')
    
    if not xray_no:
        return jsonify({"success": False, "message": "Nombor X-Ray diperlukan."}), 400
        
    config = load_config()
    if not staff_name:
        staff_name = config.get('default_staff', config.get('default_juru_xray', 'PETUGAS KAUNTER'))
        
    success, msg = cancel_patient_record(tarikh, xray_no, reason, staff_name)
    if success:
        return jsonify({"success": True, "message": msg})
    else:
        return jsonify({"success": False, "message": msg}), 500
