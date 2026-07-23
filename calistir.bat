@echo off
cd /d "%~dp0"
..\yeni_maas_projesi\sanal_ortam\Scripts\python.exe manage.py runserver 127.0.0.1:8001
pause
