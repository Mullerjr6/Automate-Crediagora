# Ferramentas de Manutencao

`limpar_downloads_crediagora.py` limpa somente a pasta dedicada configurada
por `CREDIAGORA_DOWNLOAD_DIR` (padrao: Downloads/crediagora).

Primeiro confira sem excluir:

```powershell
python .\CREDIAGORA\Script\ferramentas\limpar_downloads_crediagora.py --check
```

Sem `--check`, os arquivos listados sao excluidos. Nao executar a limpeza
durante uma exportacao nem quando os downloads forem necessarios para diagnostico.
