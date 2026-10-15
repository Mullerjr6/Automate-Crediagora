@echo off
title Configurar credenciais ERCard
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0configurar_ercard.ps1"
echo.
pause
