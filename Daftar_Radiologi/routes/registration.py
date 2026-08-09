import datetime
from flask import Blueprint, render_template, request, jsonify, redirect, url_for
from core.config import load_config, save_config, parse_mykad, build_smrp_option_map
from core.excel_engine import get_next_xray_no, get_daily_case_count, add_patient_record

registration_bp = Blueprint('registration', __name__)

@registration_bp.route('/')
def borang_index():
    config = load_config()
    # Jika aplikasi belum dikonfigurasi, halakan terus ke Setup Wizard
    if not config.get("is_configured", False):
        return redirect(url_for('settings.setup_wizard'))
        
    today = datetime.date.today()
    next_xray = get_next_xray_no(today)
    next_bil_kes = get_daily_case_count(today)
    smrp_option_map = build_smrp_option_map(config.get("active_modalities"), config.get("custom_smrp_orderables"))
    
    return render_template(
        'borang.html',
        config=config,
        smrp_option_map=smrp_option_map,
        today=today.strftime("%Y-%m-%d"),
        next_xray=next_xray,
        next_bil_kes=next_bil_kes,
        current_page='borang'
    )

@registration_bp.route('/api/next-xray')
def api_next_xray():
    tarikh_str = request.args.get('tarikh')
    if tarikh_str:
        try:
            date_obj = datetime.datetime.strptime(tarikh_str, "%Y-%m-%d").date()
        except ValueError:
            date_obj = datetime.date.today()
    else:
        date_obj = datetime.date.today()
        
    next_num = get_next_xray_no(date_obj)
    next_bil = get_daily_case_count(date_obj)
    return jsonify({"success": True, "next_xray": next_num, "next_bil_kes": next_bil})

@registration_bp.route('/api/parse-ic', methods=['POST'])
def api_parse_ic():
    data = request.get_json() or {}
    ic_number = data.get('ic_number', '')
    parsed = parse_mykad(ic_number)
    if parsed:
        return jsonify({"success": True, "data": parsed})
    return jsonify({"success": False, "message": "Format No. IC tidak sah."}), 400

@registration_bp.route('/api/set-operator', methods=['POST'])
def api_set_operator():
    """Menyimpan pilihan staff/operator bertugas semasa untuk kemudahan kaunter."""
    data = request.get_json() or {}
    operator_name = str(data.get('operator', '')).strip()
    if operator_name:
        config = load_config()
        config['default_staff'] = operator_name
        config['default_juru_xray'] = operator_name
        save_config(config)
        return jsonify({"success": True, "message": f"Staff bertugas ditetapkan kepada {operator_name}"})
    return jsonify({"success": False, "message": "Nama staff tidak sah"}), 400

@registration_bp.route('/submit', methods=['POST'])
def submit_patient():
    if request.is_json:
        patient_data = request.get_json() or {}
    else:
        patient_data = request.form.to_dict()
    
    success, message, xray_no = add_patient_record(patient_data)
    
    if success:
        return jsonify({
            "success": True,
            "message": message,
            "xray_no": xray_no
        })
    else:
        return jsonify({
            "success": False,
            "message": message
        }), 500
