@echo off
REM Double-click to start the Annexure A generator, then open the browser.
cd /d "%~dp0"
echo Starting Annexure A generator...
start "" http://127.0.0.1:5000
python app.py
pause
