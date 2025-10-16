@echo off
setlocal
cd /d "%~dp0"
title MONITOR DIARIO - 24H
python -m pip install --upgrade pip >nul
python -m pip install -r requirements.txt >nul
python monitor_diario.py
pause
endlocal
