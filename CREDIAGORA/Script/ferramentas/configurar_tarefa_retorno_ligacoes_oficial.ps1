$ErrorActionPreference = 'Stop'

$scriptDir = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
$python = (Get-Command python -ErrorAction Stop).Source
$pythonSemJanela = Join-Path (Split-Path -Parent $python) 'pythonw.exe'
if (-not (Test-Path $pythonSemJanela)) {
    throw "pythonw.exe não encontrado: $pythonSemJanela"
}
$taskName = 'Crediagora - Retorno Ligacoes Oficial'
$dias = @('Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday')
$horarios = @('07:40', '13:30', '17:05')

$action = New-ScheduledTaskAction `
    -Execute $pythonSemJanela `
    -Argument '-u tarefas\executar_retorno_ligacoes_oficial_agendado.py' `
    -WorkingDirectory $scriptDir
$triggers = @(
    $horarios | ForEach-Object {
        New-ScheduledTaskTrigger -Weekly -DaysOfWeek $dias -At $_
    }
)
$settings = New-ScheduledTaskSettingsSet `
    -AllowStartIfOnBatteries `
    -DontStopIfGoingOnBatteries `
    -StartWhenAvailable `
    -WakeToRun `
    -ExecutionTimeLimit (New-TimeSpan -Hours 3)
$principal = New-ScheduledTaskPrincipal `
    -UserId "$env:USERDOMAIN\$env:USERNAME" `
    -LogonType Interactive `
    -RunLevel Limited
$task = New-ScheduledTask -Action $action -Trigger $triggers -Settings $settings -Principal $principal
Register-ScheduledTask -TaskName $taskName -InputObject $task -Force | Out-Null

Write-Host "Tarefa criada: $taskName"
Get-ScheduledTask -TaskName $taskName | Select-Object TaskName, State
