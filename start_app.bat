@echo off
title TallyDataX - Tally Voucher Extractor and Excel Suite
cd /d "%~dp0"

echo ========================================================
echo   Starting TallyDataX Application...
echo   App URL: http://127.0.0.1:8000
echo   Tally Prime Port: 9000
echo ========================================================

python -c "import fastapi, uvicorn, openpyxl, requests, multipart" >nul 2>&1
if %errorlevel% neq 0 (
    echo [INFO] Installing required dependencies...
    pip install fastapi uvicorn openpyxl requests python-multipart
)

python -c "import urllib.request; urllib.request.urlopen('http://127.0.0.1:8000/', timeout=1)" >nul 2>&1
if %errorlevel% equ 0 (
    echo [INFO] TallyDataX Server is already running on port 8000.
    echo [INFO] Opening application in your browser...
    start http://127.0.0.1:8000
    exit
)

echo [INFO] Starting FastAPI / Uvicorn Server and Opening Browser...
python app.py
pause

