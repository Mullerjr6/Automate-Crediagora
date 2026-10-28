$root = Split-Path -Parent $PSScriptRoot
$python = (Get-Command pythonw.exe).Source
$principal = New-ScheduledTaskPrincipal -UserId "$env:USERDOMAIN\$env:USERNAME" -LogonType Interactive -RunLevel Limited
$settings = New-ScheduledTaskSettingsSet -StartWhenAvailable -ExecutionTimeLimit (New-TimeSpan -Hours 3)

$acaoSequencia = New-ScheduledTaskAction -Execute $python -Argument "`"$root\tarefas\executar_sequencia_exportacoes_com_fallback.py`""
$gatilhoSequencia = New-ScheduledTaskTrigger -Once -At (Get-Date "2026-10-28 12:50")
Register-ScheduledTask -TaskName "Crediagora - Sequencia Exportacoes Hoje 1250" -Action $acaoSequencia -Trigger $gatilhoSequencia -Principal $principal -Settings $settings -Force | Out-Null

$acaoRestaurar = New-ScheduledTaskAction -Execute $python -Argument "`"$root\tarefas\restaurar_tarefa_tabela_contratos.py`""
$gatilhoRestaurar = New-ScheduledTaskTrigger -Once -At (Get-Date "2026-10-28 18:00")
Register-ScheduledTask -TaskName "Crediagora - Restaurar Tabela Contratos Hoje" -Action $acaoRestaurar -Trigger $gatilhoRestaurar -Principal $principal -Settings $settings -Force | Out-Null

Write-Output "tarefas override configuradas"
