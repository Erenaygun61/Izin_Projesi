@echo off
setlocal
cd /d "%~dp0"
title TISKI Izin Sistemi

if not exist "sanal_ortam\Scripts\python.exe" (
    echo Izin sistemi ilk kez hazirlaniyor. Lutfen bekleyin...
    py -3 -m venv sanal_ortam 2>nul
    if errorlevel 1 python -m venv sanal_ortam
    if errorlevel 1 (
        echo.
        echo Python bulunamadi. Once python.org adresinden Python 3 kurun.
        pause
        exit /b 1
    )
    sanal_ortam\Scripts\python.exe -m pip install --upgrade pip
    sanal_ortam\Scripts\python.exe -m pip install -r requirements.txt
    if errorlevel 1 (
        echo.
        echo Gerekli paketler yuklenemedi. Internet baglantisini kontrol edip tekrar deneyin.
        pause
        exit /b 1
    )
)

echo Izin Projesi baslatiliyor...
echo Adres: http://127.0.0.1:8001/
echo Kapatmak icin bu pencerede CTRL+C yapin, sorarsa Y yazip Enter'a basin.
echo.
sanal_ortam\Scripts\python.exe manage.py runserver 127.0.0.1:8001
pause
