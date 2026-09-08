@echo off
setlocal

title Crediagora - Automacao
cd /d "%~dp0"

set "PROJECT_DIR=%~dp0..\.."
set "VENV_PYTHON=%PROJECT_DIR%\.venv\Scripts\python.exe"

if exist "%VENV_PYTHON%" (
    "%VENV_PYTHON%" start.py %*
) else (
    where py >nul 2>nul
    if errorlevel 1 (
        python start.py %*
    ) else (
        py -3 start.py %*
    )
)

set "EXIT_CODE=%ERRORLEVEL%"
echo.

if "%EXIT_CODE%"=="0" (
    echo Automacao finalizada.
) else (
    echo Automacao finalizada com erro. Verifique a pasta de logs/erros.
)

pause
exit /b %EXIT_CODE%
