@echo off
cd /d "%~dp0"
echo Installing packages: pandas, openpyxl, matplotlib, python-pptx, pillow ...
echo.
py -m pip install -r requirements.txt
if errorlevel 1 python -m pip install -r requirements.txt
echo.
echo ===== Done. You can close this window and reopen the program. =====
pause
