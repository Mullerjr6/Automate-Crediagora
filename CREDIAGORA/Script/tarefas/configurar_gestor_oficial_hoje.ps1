$root = Split-Path -Parent $PSScriptRoot
$python = (Get-Command pythonw.exe).Source
$principal = New-ScheduledTaskPrincipal -UserId "$env:USERDOMAIN\$env:USERNAME" -LogonType Interactive -RunLevel Limited
$settings = New-ScheduledTaskSettingsSet -StartWhenAvailable -ExecutionTimeLimit (New-TimeSpan -Hours 3)
$acao = New-ScheduledTaskAction -Execute $python -Argument "`"$root\completo.py`" --somente-gestor"
$gatilho = New-ScheduledTaskTrigger -Once -At (Get-Date "2026-10-28 13:20")
Register-ScheduledTask -TaskName "Crediagora - Gestor Oficial Hoje 1320" -Action $acao -Trigger $gatilho -Principal $principal -Settings $settings -Force | Out-Null
Write-Output "tarefa configurada"
