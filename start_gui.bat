@echo off
setlocal
cd /d "%~dp0"
title SISTEMA MONITOR (GUI)
python -m pip install --upgrade pip >nul
python -m pip install -r requirements.txt >nul
python sistema_monitor_dispositivos.py
endlocal
