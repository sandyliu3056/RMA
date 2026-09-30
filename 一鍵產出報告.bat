@echo off
chcp 65001 >nul
cd /d "%~dp0"
set /p F=AIO file path: 
py -3 rma_engine.py "%F%" RMA_output || python rma_engine.py "%F%" RMA_output
pause
