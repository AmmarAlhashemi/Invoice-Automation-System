import os
from pathlib import Path

class Config:
    # Basic Flask settings
    SECRET_KEY = os.urandom(24)
    
    # Define main paths
    BASE_DIR = Path(__file__).parent
    APP_DIR = BASE_DIR / "app"
    
    INV_DIR = BASE_DIR / "invoices"
    STATEMENTS_DIR = BASE_DIR / "statements"
    
    # Data files and templates
    EXCEL_PATH = STATEMENTS_DIR / "Account_1_Statement.xlsx"
    INV_TEMPLATE_PATH = APP_DIR / "templates" / "MASTER INV Template.docx"
    WORD_STATEMENT_TEMPLATE = APP_DIR / "templates" / " Statement Template.docx"
    
    # Application constants
    DATE_FORMAT = "%m/%d/%Y"
    OPENING_TEXT = "OPENING BALANCE B/F"