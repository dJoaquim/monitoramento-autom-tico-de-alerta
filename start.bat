@echo off
setlocal
cd /d "%~dp0"

title SISTEMA DE MONITORAMENTO - PIROMETRO / TEMPORIZADOR / DINOMETRO

echo.
echo ==========================================================
echo       INICIANDO SISTEMA DE MONITORAMENTO DE DISPOSITIVOS
echo ==========================================================

REM === Instala dependencias ===
echo [1/2] Verificando dependencias do Python...
python -m pip install --upgrade pip >nul
python -m pip install -r requirements.txt >nul

REM === Abre planilha principal ===
echo [2/2] Abrindo planilha de monitoramento...
start "" "monitor_multidispositivo.xlsx"

REM === Abre o sistema com interface (GUI) ===
echo Iniciando tela principal do sistema...
python sistema_monitor_dispositivos.py

echo.
echo ==========================================================
echo SISTEMA FINALIZADO.
echo ==========================================================
timeout /t 10 >nul
exit
