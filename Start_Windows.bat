@echo off
cd /d "%~dp0"
echo Preparing the installer...
python -m pip install --upgrade pip
if errorlevel 1 goto failed
echo Installing the application requirements...
python -m pip install -r requirements.txt
if errorlevel 1 goto failed
python -m streamlit run app.py
pause
exit /b 0
:failed
echo Setup failed. Check your internet connection and make sure standard Python 3.12, 3.13, or 3.14 is installed.
pause
exit /b 1
