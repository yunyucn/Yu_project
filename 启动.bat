@echo off
setlocal
title Valuation Web - Keep this window open
cd /d "%~dp0"

where python >nul 2>nul
if errorlevel 1 goto missing_python

python -c "import flask, pdfplumber, openpyxl, olefile, waitress, psutil" >nul 2>nul
if errorlevel 1 goto install_dependencies
goto start_service

:install_dependencies
echo Installing dependencies. Please wait...
python -m pip install -r requirements.txt
if errorlevel 1 goto install_failed

:start_service
set HOST=127.0.0.1
set PORT=8000
set PYTHONIOENCODING=utf-8
set PYTHONUNBUFFERED=1
echo.
echo Open in your browser: http://127.0.0.1:8000/
echo Keep this window open while using the app.
echo To stop: press Ctrl+C in this window.
echo Closing the browser does NOT stop the service.
echo.
python app.py
echo.
echo Service stopped. If startup failed, copy the error shown above.
pause
exit /b

:missing_python
echo Python was not found. Install Python and add it to PATH.
pause
exit /b 1

:install_failed
echo Dependency installation failed. Copy the error shown above.
pause
exit /b 1
