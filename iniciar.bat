@echo off
cd /d "%~dp0"
echo.
echo   GENERADOR DE INFORMES - DISTRITO 17D08
echo.
pip install -r requirements.txt --quiet 2>nul
echo   Abriendo navegador...
python app.py
pause
