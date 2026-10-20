@echo off
setlocal
cd /d "C:\Users\CREDIAGORA\Desktop\workspace\Automacoes\crediagora\CREDIAGORA\Script"
python -u ercard_carteira.py --limpar-pasta >> "logs\limpeza_ercard_carteira.log" 2>&1
exit /b %ERRORLEVEL%
