import datetime
from flask import Blueprint, render_template, request, make_response
from core.export_engine import generate_export_data

export_bp = Blueprint('export', __name__)

@export_bp.route('/export/print')
def print_buku_daftar():
    today = datetime.date.today()
    year = int(request.args.get('year', today.year))
    month = int(request.args.get('month', today.month))
    
    export_data = generate_export_data(year, month)
    
    return render_template(
        'print_template.html',
        data=export_data
    )

@export_bp.route('/export/pdf')
def pdf_buku_daftar():
    today = datetime.date.today()
    year = int(request.args.get('year', today.year))
    month = int(request.args.get('month', today.month))
    
    export_data = generate_export_data(year, month)
    html_content = render_template('print_template.html', data=export_data, is_pdf=True)
    
    # Pulangkan HTML bersedia-cetak untuk paparan PDF penyemak imbas / download
    response = make_response(html_content)
    response.headers['Content-Type'] = 'text/html; charset=utf-8'
    return response
