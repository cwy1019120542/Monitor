from flask import Flask
from .config import config
from .views.tools import tools_blueprint
from .views.errors import errors_blueprint

def create_app(config_name):
    app = Flask(__name__)
    app.config.from_object(config[config_name])
    app.register_blueprint(tools_blueprint, url_prefix='/monitor/api/tools')
    app.register_blueprint(errors_blueprint)
    return app
