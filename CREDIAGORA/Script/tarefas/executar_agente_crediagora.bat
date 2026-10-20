@echo off
setlocal
cd /d "C:\Users\CREDIAGORA\Desktop\workspace\Automacoes\crediagora\CREDIAGORA\Script"
python -u agente_crediagora.py --fluxo %1 --tentativas 2
exit /b %ERRORLEVEL%
