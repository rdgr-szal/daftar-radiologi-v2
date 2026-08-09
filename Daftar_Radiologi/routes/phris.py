import datetime
from flask import Blueprint, render_template, request, jsonify
from core.config import load_config, MONTH_MAP
from core.phris_engine import get_phris_matrix_data

phris_bp = Blueprint('phris', __name__)

@phris_bp.route('/phris')
def phris_report():
    config = load_config()
    today = datetime.date.today()
    selected_year = int(request.args.get('year', today.year))
    
    start_year = min(2024, today.year)
    available_years = list(range(start_year, today.year + 2))
    
    return render_template(
        'reten_phris.html',
        config=config,
        selected_year=selected_year,
        available_years=available_years,
        current_year=today.year,
        current_page='phris'
    )

@phris_bp.route('/api/phris-matrix')
def api_phris_matrix():
    today = datetime.date.today()
    year = int(request.args.get('year', today.year))
    
    matrix = get_phris_matrix_data(year)
    return jsonify(matrix)
