from flask import Blueprint, request, flash, redirect, url_for, render_template
from app.services.statement_service import (
    process_advance_payment, 
    generate_statement_docs, 
    get_statement_rows, 
    update_opening_balance
)
import pythoncom

statement_bp = Blueprint('statement', __name__)

@statement_bp.route('/statement')
def statement_page():
    data = get_statement_rows()
    return render_template('statement.html', rows=data['rows'], current_bf=data['current_bf'])

@statement_bp.route('/add-advance', methods=['POST'])
def add_advance():
    date = request.form.get('advance_date')
    amount = float(request.form.get('advance_amount'))
    excel_status, msg = process_advance_payment(date, amount, "ADVANCE RECEIVED")
    flash(msg, "success" if excel_status else "error")
    return redirect(url_for('statement.statement_page'))

@statement_bp.route('/generate-statement', methods=['POST'])
def generate_statement():
    pythoncom.CoInitialize()
    try:
        success, msg = generate_statement_docs()
        flash(msg, "success" if success else "warning")
    finally:
        pythoncom.CoUninitialize()
    return redirect(url_for('statement.statement_page'))

@statement_bp.route('/update-opening', methods=['POST'])
def update_opening():
    insert_at = int(request.form.get('selected_row'))
    success, msg = update_opening_balance(insert_at)
    flash(msg, "success" if success else "error")
    return redirect(url_for('statement.statement_page'))