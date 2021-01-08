import os
from flask_cors import CORS
from flask_script import Manager
from app.factory import create_app

config_name = os.getenv('FLASK_CONFIG', 'base')
app = create_app(config_name)
CORS(app, supports_credentials=True)
manage = Manager(app)

if __name__ == '__main__':
    manage.run()