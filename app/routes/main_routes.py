from flask import Blueprint, render_template
from app.services.invoice_service import get_next_number_by_type, get_last_invoice_info

main_bp = Blueprint('main', __name__)

@main_bp.route('/')
def index():
    next_normal = get_next_number_by_type("normal")
    next_coc = get_next_number_by_type("coc")
    
    last_normal = get_last_invoice_info("normal")[1]
    last_coc = get_last_invoice_info("coc")[1]

    return render_template('index.html', 
                           next_inv_normal=next_normal, 
                           next_inv_coc=next_coc,
                           last_inv_normal=last_normal[:-5] if last_normal != "No previous invoices" else last_normal,
                           last_inv_coc=last_coc[:-5] if last_coc != "No previous invoices" else last_coc)