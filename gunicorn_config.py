import os
from gevent import monkey
monkey.patch_all()
debug = False
bind = f"{os.environ.get('GUNICORN_BIND_IP', '127.0.0.1')}:{os.environ.get('GUNICORN_BIND_PORT', '8000')}"
# bind = "192.168.43.115:8000"
loglevel = 'debug'
pidfile = 'app/logs/gunicorn_pid.txt'
logfile = 'app/logs/gunicorn_log.txt'
workers = 4
worker_class = 'gevent'
timeout = 500