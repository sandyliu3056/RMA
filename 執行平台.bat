@echo off
chcp 65001 >nul
cd /d "%~dp0"
py -3 "零件規劃平台.py" || python "零件規劃平台.py"
pause
