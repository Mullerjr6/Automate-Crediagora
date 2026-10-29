@echo off
setlocal
cd /d "C:\Users\CREDIAGORA\Desktop\workspace\Automacoes\crediagora\CREDIAGORA\Script"
if not exist "logs" mkdir "logs"
python -u panorama_exportacoes_carteira.py >> "logs\ercard_carteira.log" 2>&1
exit /b %ERRORLEVEL%
