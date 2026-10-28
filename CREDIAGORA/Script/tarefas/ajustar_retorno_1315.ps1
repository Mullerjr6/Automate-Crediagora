$nome = 'Crediagora - Retorno Ligacoes Oficial'
$tarefa = Get-ScheduledTask -TaskName $nome
$gatilhos = @($tarefa.Triggers)
foreach ($gatilho in $gatilhos) {
    if ($gatilho.StartBoundary -match 'T13:20:00') {
        $gatilho.StartBoundary = $gatilho.StartBoundary -replace 'T13:20:00', 'T13:15:00'
    }
}
Set-ScheduledTask -TaskName $nome -Trigger $gatilhos | Out-Null
Write-Output 'Retorno ajustado para 13:15'
