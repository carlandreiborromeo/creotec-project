# backend/app/__init__.py
from flask import Flask
from flask_cors import CORS

def create_app():
    app = Flask(__name__)

    # app.config.from_pyfile('config.py', silent=True)  

    CORS(app)

    @app.route('/')
    def index():
        return 'Backend server is running'

    # Import and register blueprints
    # from .routes.auth import bp as auth_bp
    from .routes.upload import bp as upload_bp
    # from .routes.preview import bp as preview_bp
    from .routes.generate import bp as generate_bp

    # app.register_blueprint(auth_bp)
    app.register_blueprint(upload_bp)
    # app.register_blueprint(preview_bp)
    app.register_blueprint(generate_bp)
    

    return app
