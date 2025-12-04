@echo off
chcp 65001 > nul
title Convertidor PDF a Excel - APU con VAE

echo =========================================
echo   CONVERTIDOR DE APU (PDF a Excel)
echo   Formato con VAE
echo =========================================
echo.

if "%~1"=="" (
    echo Arrastra un archivo PDF sobre este .bat
    echo O ejecuta: convertir_apu.bat archivo.pdf
    echo.
    pause
    exit /b 1
)

echo Procesando: %~1
echo.

"%~dp0.venv\Scripts\python.exe" "%~dp0pdf_to_excel_apu.py" "%~1"

echo.
echo =========================================
pause
