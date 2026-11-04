@echo off
chcp 65001 >nul
title Envio de WhatsApp - Instalador
cd /d "%~dp0"

echo ============================================================
echo   Creando EnvioWhatsApp.xlsm (Excel con macros)
echo   No abras ni cierres Excel mientras se crea el archivo.
echo ============================================================
echo.

powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0src\crear_excel.ps1"
set RESULTADO=%ERRORLEVEL%

echo.
if "%RESULTADO%"=="1" (
    echo Hubo un error al crear el archivo. Revisa el mensaje de arriba.
    pause
    exit /b 1
)

echo Abriendo EnvioWhatsApp.xlsm ...
echo Si Excel muestra "Habilitar contenido" o "Habilitar macros", presionalo.
start "" "%~dp0EnvioWhatsApp.xlsm"
echo.
pause
