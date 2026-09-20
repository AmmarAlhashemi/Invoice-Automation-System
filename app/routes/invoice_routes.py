from flask import Blueprint, request, flash, redirect, url_for
from app.services.invoice_service import process_and_generate_invoice
import pythoncom

invoice_bp = Blueprint('invoice', __name__)

@invoice_bp.route('/generate', methods=['POST'])
def process_invoice():
    pythoncom.CoInitialize()
    try:
        # Call the processing function from the service
        success, message = process_and_generate_invoice(request.form)
        if success:
            flash(message, "success")
        else:
            flash(message, "error")
    except Exception as e:
        flash(f"Error: {e}", "error")
    finally:
        pythoncom.CoUninitialize()

    return redirect(url_for('main.index'))