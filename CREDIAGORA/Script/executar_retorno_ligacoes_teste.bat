@echo off
setlocal
cd /d "C:\Users\CREDIAGORA\Desktop\workspace\Automacoes\crediagora\CREDIAGORA\Script"
if not exist "logs" mkdir "logs"
python -u start.py --exportacao retorno-ligacoes --sem-ercard --sem-gestor --headless >> "logs\agendado_retorno_ligacoes.log" 2>&1
set "CODIGO=%ERRORLEVEL%"
echo [%DATE% %TIME%] Codigo de saida: %CODIGO%>> "logs\agendado_retorno_ligacoes.log"
exit /b %CODIGO%
