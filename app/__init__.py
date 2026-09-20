from flask import Flask
from config import Config
import os

def create_app(config_class=Config):
    # Set the templates folder (pages) and static folder (static)
    app = Flask(__name__, template_folder='pages', static_folder='static')
    app.config.from_object(config_class)

    # Ensure the output directories exist
    os.makedirs(app.config['INV_DIR'], exist_ok=True)
    os.makedirs(app.config['STATEMENTS_DIR'], exist_ok=True)

    # Import and register the Blueprints
    from app.routes.main_routes import main_bp
    from app.routes.invoice_routes import invoice_bp
    from app.routes.statement_routes import statement_bp

    app.register_blueprint(main_bp)
    app.register_blueprint(invoice_bp)
    app.register_blueprint(statement_bp)

    return app