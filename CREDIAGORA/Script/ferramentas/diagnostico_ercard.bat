@echo off
rem Valida apenas a configuração ERCard, sem login ou exportação.
call "%~dp0iniciar.bat" --check --somente-ercard
