import datetime
from flask import Blueprint, render_template, request, make_response
from core.export_engine import generate_export_data

export_bp = Blueprint('export', __name__)

@export_bp.route('/export/print')
def print_buku_daftar():
    today = datetime.date.today()
    period = request.args.get('period', 'month')
    selected_date = request.args.get('day_date', str(today))
    start_date = request.args.get('start_date', '')
    end_date = request.args.get('end_date', '')
    
    try:
        year = int(request.args.get('year', today.year))
    except (ValueError, TypeError):
        year = today.year

    try:
        month = int(request.args.get('month', today.month))
    except (ValueError, TypeError):
        month = today.month
    
    export_data = generate_export_data(
        year=year,
        month=month,
        period=period,
        selected_date_str=selected_date,
        start_date_str=start_date,
        end_date_str=end_date
    )
    
    return render_template(
        'print_template.html',
        data=export_data
    )

@export_bp.route('/export/pdf')
def pdf_buku_daftar():
    today = datetime.date.today()
    period = request.args.get('period', 'month')
    selected_date = request.args.get('day_date', str(today))
    start_date = request.args.get('start_date', '')
    end_date = request.args.get('end_date', '')
    
    try:
        year = int(request.args.get('year', today.year))
    except (ValueError, TypeError):
        year = today.year

    try:
        month = int(request.args.get('month', today.month))
    except (ValueError, TypeError):
        month = today.month
    
    export_data = generate_export_data(
        year=year,
        month=month,
        period=period,
        selected_date_str=selected_date,
        start_date_str=start_date,
        end_date_str=end_date
    )
    html_content = render_template('print_template.html', data=export_data, is_pdf=True)
    
    response = make_response(html_content)
    response.headers['Content-Type'] = 'text/html; charset=utf-8'
    return response
