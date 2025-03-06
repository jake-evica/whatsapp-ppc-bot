from flask import Flask

from app.config import Config
from app.routes.whatsapp import whatsapp_bp


def create_app() -> Flask:
    app = Flask(__name__)
    app.config.from_object('config.Config')
    app.register_blueprint(whatsapp_bp, url_prefix="/whatsapp/")
    return app
