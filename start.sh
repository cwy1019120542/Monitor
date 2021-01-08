#!/bin/bash
service vsftpd restart
service ssh restart
source /etc/profile
nohup gunicorn /home/cwy/Monitor/manage:app -c /home/cwy/Monitor/gunicorn_config.py 2>&1 &