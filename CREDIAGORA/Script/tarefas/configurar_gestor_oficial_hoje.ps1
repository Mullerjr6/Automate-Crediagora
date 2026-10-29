$root = Split-Path -Parent $PSScriptRoot
$python = (Get-Command pythonw.exe).Source
$principal = New-ScheduledTaskPrincipal -UserId "$env:USERDOMAIN\$env:USERNAME" -LogonType Interactive -RunLevel Limited
$settings = New-ScheduledTaskSettingsSet -StartWhenAvailable -ExecutionTimeLimit (New-TimeSpan -Hours 3)
$acao = New-ScheduledTaskAction -Execute $python -Argument "-u `"$root\tarefas\executar_gestor_com_exportacoes_agendado.py`"" -WorkingDirectory $root
$gatilho = New-ScheduledTaskTrigger -Once -At (Get-Date "2026-10-29 13:12")
Register-ScheduledTask -TaskName "Crediagora - Gestor Exportacoes Hoje 1312" -Action $acao -Trigger $gatilho -Principal $principal -Settings $settings -Force | Out-Null
Write-Output "tarefa configurada"
