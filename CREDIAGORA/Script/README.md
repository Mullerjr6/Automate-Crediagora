# Automacoes Crediagora

## Arquivos principais

- `completo.py`: processo geral de Panorama, ERCard e Gestor.
- `agente_crediagora.py`: orquestrador com tentativas, estado e relatórios.
- `retorno_ligacoes_oficial.py`: atualização autorizada somente do Retorno de Ligações oficial.
- `Manha.py`: fluxo separado de Contratos, Propostas e Carteira Mifa.
- `tabela_contratos.py`: fluxo separado somente do relatório 14 - Tabela Contratos.

## Pastas

- `tarefas/`: executores usados por tarefas agendadas e atalhos.
- `ferramentas/`: configuração, diagnóstico e manutenção.
- `logs/`: histórico, estado, painel HTML e registros de execução.
- `backups/`: cópias de segurança criadas antes das atualizações.
- `testes_manuais/`: testes de interface mantidos para diagnóstico.

## Retorno de Ligações oficial

A tarefa `Crediagora - Retorno Ligacoes Oficial` roda às 07:40, 13:30 e 17:05 nos dias úteis, ignorando feriados nacionais. Para executar manualmente:

```powershell
python -u .\agente_crediagora.py --fluxo retorno-ligacoes-oficial
```

## ERCard ? Tabela Contratos

Para extrair somente o relat?rio 14, com estado, tentativas e valida??o do CSV:

```powershell
python -u .\agente_crediagora.py --fluxo ercard-tabela-contratos
```
