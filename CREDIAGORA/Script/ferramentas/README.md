# Ferramentas de Manutencao

## Tarefa do Retorno de Ligações oficial

`configurar_tarefa_retorno_ligacoes_oficial.ps1` recria a tarefa agendada de
segunda a sexta às 07:40, 13:30 e 17:05. Ela executa sem janela de terminal,
desperta o computador quando possível e ignora feriados nacionais.

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\ferramentas\configurar_tarefa_retorno_ligacoes_oficial.ps1
```

`limpar_downloads_crediagora.py` limpa somente a pasta dedicada configurada
por `CREDIAGORA_DOWNLOAD_DIR` (padrao: Downloads/crediagora).

Primeiro confira sem excluir:

```powershell
python .\CREDIAGORA\Script\ferramentas\limpar_downloads_crediagora.py --check
```

Sem `--check`, os arquivos listados sao excluidos. Nao executar a limpeza
durante uma exportacao nem quando os downloads forem necessarios para diagnostico.
