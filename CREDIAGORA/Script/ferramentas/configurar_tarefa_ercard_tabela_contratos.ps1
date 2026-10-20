$ErrorActionPreference = 'Stop'
$nome = 'Crediagora - ERCard Tabela Contratos'
$pasta = 'C:\Users\CREDIAGORA\Desktop\workspace\Automacoes\crediagora\CREDIAGORA\Script'
$pythonw = 'C:\Users\CREDIAGORA\AppData\Local\Microsoft\WindowsApps\pythonw.exe'
$dias = @('Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday')
$horario_1250 = [datetime]'2026-10-20 12:50:00'
$horario_1700 = [datetime]'2026-10-20 17:00:00'
$gatilho_1250 = New-ScheduledTaskTrigger -Weekly -DaysOfWeek $dias -At $horario_1250
$gatilho_1700 = New-ScheduledTaskTrigger -Weekly -DaysOfWeek $dias -At $horario_1700
$triggers = @($gatilho_1250, $gatilho_1700)
$acao = New-ScheduledTaskAction -Execute $pythonw -Argument '-u tarefas\executar_ercard_tabela_contratos_agendado.py' -WorkingDirectory $pasta
$configuracao = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -WakeToRun -StartWhenAvailable -MultipleInstances IgnoreNew -ExecutionTimeLimit (New-TimeSpan -Hours 3)
Register-ScheduledTask -TaskName $nome -Action $acao -Trigger $triggers -Settings $configuracao -Description 'Extrai o relatorio 14 - Tabela Contratos do ERCard em dias uteis, excluindo feriados nacionais.' -Force | Out-Null
Get-ScheduledTask -TaskName $nome | Select-Object TaskName, State
Get-ScheduledTaskInfo -TaskName $nome | Select-Object NextRunTime, LastRunTime, LastTaskResult
