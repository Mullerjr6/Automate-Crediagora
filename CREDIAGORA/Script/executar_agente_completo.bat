@echo off
setlocal
cd /d "C:\Users\CREDIAGORA\Desktop\workspace\Automacoes\crediagora\CREDIAGORA\Script"
python -u agente_crediagora.py --fluxo completo >> "logs\agente_completo.log" 2>&1
exit /b %ERRORLEVEL%
