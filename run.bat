@echo off
REM Install dependencies (first run only) then launch the dashboard.
pip install -q -r "%~dp0requirements.txt"
python "%~dp0financial_dashboard.py"
pause
