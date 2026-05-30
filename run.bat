@echo off
cd /d "%~dp0"
set PYTHON=C:\Users\admin\AppData\Local\Python\pythoncore-3.14-64\python.exe
if not exist "%PYTHON%" set PYTHON=python

echo Installing required packages...
"%PYTHON%" -m pip install gspread google-auth google-auth-oauthlib openpyxl altair --quiet

echo Starting app...
"%PYTHON%" -m streamlit run app.py --server.port 8501 --server.address 0.0.0.0 --browser.gatherUsageStats false
pause
