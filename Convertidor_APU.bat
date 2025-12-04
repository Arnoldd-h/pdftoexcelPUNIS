@echo off
chcp 65001 > nul
title Convertidor APU - Interfaz Gráfica

cd /d "%~dp0"
start "" "%~dp0.venv\Scripts\pythonw.exe" "%~dp0convertidor_gui.py" %*
