# Automação Crediagora

Automação em Python para atualizar as bases do Crediagora e, em seguida, exportar `14 - TABELA CONTRATOS` do ERCard para CSV.

## Estrutura

Depois de cada atualizacao do Panorama, o arquivo
`CREDIAGORA/logs de vendas/DD-MM-AAAA.txt` registra os totais do mes vigente
exportado, conforme DATA CPC. A secao VENDAS soma `Valor Contrato` e a secao
RECEITA soma `COMISSAO PREVISTA`. Cada secao informa seu proprio horario de
atualizacao, quantidade de registros e planilha de destino.
O resumo e gravado somente depois do XLSX. Reexecutar no mesmo dia substitui
o total daquela base e preserva a outra secao; nao acumula os totais entre execucoes.
Os horarios ficam no TXT. A inclusao de uma nova planilha sera configurada separadamente.

- `CREDIAGORA/Script/start.py`: automação principal.
- `CREDIAGORA/Script/ercard.py`: fase ERCard isolada, com portal web e RemoteApp.
- `CREDIAGORA/Script/iniciar.bat`: executa a automação no Windows e usa `.venv` automaticamente quando existir.
- `CREDIAGORA/Script/iniciar_ercard.bat`: executa somente a fase ERCard.
- `CREDIAGORA/Script/configurar_ercard.bat`: solicita e registra os dois logins sem exibir as senhas.
- `CREDIAGORA/Script/diagnostico.bat`: valida configuração sem abrir Chrome.
- `CREDIAGORA/Script/diagnostico_ercard.bat`: valida apenas a configuração ERCard.
- `CREDIAGORA/Script/liberar_login.bat`: libera a proteção local após o desbloqueio da conta no portal.
- `CREDIAGORA/Script/gestor.py`: fase Gestor.
- `CREDIAGORA/Script/indicadores_fpd1.py`: importacao e validacao dos indicadores.
- `CREDIAGORA/Script/receita_cpc.py`: normalizacao de DATA CPC.
- `CREDIAGORA/Script/ferramentas/`: utilitarios de manutencao, separados da execucao normal.
- `CREDIAGORA/Script/testes_manuais/`: testes controlados historicos do ERCard; nao sao executados pelo fluxo principal.
- `tests/`: testes automatizados offline, sem exportacoes reais.
- `CREDIAGORA/backups/`: backups gerados antes de atualizar planilhas.
- `CREDIAGORA/logs/`: logs de execução.
- `CREDIAGORA/erros/`: prints e HTML salvos em falhas.

## Instalação

Os comandos de inicio continuam nos mesmos caminhos. Para conferir o codigo sem
abrir os sistemas nem alterar planilhas reais, execute na raiz do projeto:

```powershell
python -m unittest discover -s tests
```

A ferramenta de limpeza foi movida. Para apenas listar os arquivos candidatos,
sem excluir downloads:

```powershell
python .\CREDIAGORA\Script\ferramentas\limpar_downloads_crediagora.py --check
```

```powershell
cd "C:\Users\CREDIAGORA\Desktop\Automacoes\crediagora"
py -3 -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
pip install -r requirements.txt
```

## Configuração

A senha não deve ficar salva no código. Configure pelo PowerShell:

```powershell
setx CREDIAGORA_SENHA "SUA_SENHA_AQUI"
```

Depois de usar `setx`, feche e abra o terminal novamente.

Variáveis opcionais:

- `CREDIAGORA_USUARIO`: usuário do portal. Padrão: `junior.muller`.
- `CREDIAGORA_DOWNLOAD_DIR`: pasta dedicada para downloads. Padrão: `%USERPROFILE%\Downloads\crediagora`.
- `CREDIAGORA_TABELA_FAT_DIR`: pasta das planilhas `fat`.
- `CREDIAGORA_HEADLESS`: use `1` para rodar Chrome invisível.
- `CREDIAGORA_MANTER_DOWNLOADS`: use `1` para manter arquivos baixados após sucesso.
- `CREDIAGORA_MANTER_NAVEGADOR`: use `1` para manter Chrome aberto ao final.
- `CREDIAGORA_DOWNLOAD_TIMEOUT`: tempo máximo de download em segundos.
- `CREDIAGORA_EXCEL_TIMEOUT`: tempo máximo aguardando Excel liberar arquivo.

Há um exemplo em `config.example.ps1`.

Credenciais ERCard obrigatórias:

```powershell
setx ERCARD_PORTAL_USUARIO "USUARIO_DO_PORTAL"
setx ERCARD_PORTAL_SENHA "SENHA_DO_PORTAL"
setx ERCARD_SISTEMA_USUARIO "USUARIO_DO_SISTEMA"
setx ERCARD_SISTEMA_SENHA "SENHA_DO_SISTEMA"
```

As senhas nunca são exibidas nos logs. O arquivo final é salvo por padrão em `%USERPROFILE%\Desktop\Exportações`.

Como alternativa aos comandos, dê dois cliques em `CREDIAGORA/Script/configurar_ercard.bat` e preencha os quatro campos solicitados.

## Uso

### Retorno de ligações (Panorama)

O fluxo padrão inclui Vendas, Receita e Retorno de ligações, nesta ordem, antes
das fases ERCard/Indicadores e Gestor. Para executar somente a nova base, na pasta `CREDIAGORA/Script`:

```powershell
python start.py --exportacao retorno-ligacoes --sem-ercard --sem-gestor
```

Para executar todas as bases do Panorama:

```powershell
python start.py --sem-ercard --sem-gestor
```

A nova rotina seleciona exclusivamente `Gerar Arquivo exportacao contato campanha Normal`
em Exportações na tela inicial. Reutiliza a tela após o login e retorna ao início
quando necessário depois de Vendas/Receita. O ZIP novo é lido diretamente na pasta de downloads já configurada.
Arquivos existentes antes da execução são preservados. Downloads desta execução
são removidos após sucesso, exceto com `--manter-downloads`; em erro são mantidos.

O módulo `retorno_ligacoes.py` substitui A2:I da aba identificada pelos cabeçalhos
de `fat_retorno_ligacoes - teste automacao.xlsx`, dentro de `CREDIAGORA_TABELA_FAT_DIR`.
Preserva os campos como texto, inclusive CPF e telefones com zeros iniciais.
A coluna E (`datacontato` no CSV, `Data_Contrato` no modelo) mantém datas nativas
do Excel quando o modelo usa formato de data, para preservar o cálculo de AnoMes.
As fórmulas existentes de J:L são traduzidas até a última linha nova; valores e
fórmulas excedentes são limpos, preservando estilos, cabeçalhos e outras abas.
O Excel recalcula as fórmulas ao abrir o arquivo.

O backup segue a pasta e a opção `--sem-backup` existentes. A gravação é feita
em arquivo temporário no mesmo diretório: todos os registros e fórmulas são
conferidos depois de reabrir o XLSX, antes da substituição atômica do original.
ZIP inválido, CSV vazio/inconsistente, cabeçalhos incompatíveis, aba ambígua,
fórmula ausente ou alteração concorrente interrompem a atualização.
Cada etapa registra seu andamento e a duração total da nova base em minutos.

Diagnóstico sem login/download:

```powershell
py -3 .\CREDIAGORA\Script\start.py --check
```

Teste exclusivo de login, sem exportar ou alterar planilhas:

```powershell
py -3 .\CREDIAGORA\Script\start.py --check-login
```

Execução completa:

```powershell
py -3 .\CREDIAGORA\Script\start.py
```

Executar somente o ERCard:

```powershell
py -3 .\CREDIAGORA\Script\start.py --somente-ercard
```

Executar somente o fluxo antigo do Crediagora:

```powershell
py -3 .\CREDIAGORA\Script\start.py --sem-ercard
```

Rodar somente vendas:

```powershell
py -3 .\CREDIAGORA\Script\start.py --exportacao vendas
```

Rodar somente receita:

```powershell
py -3 .\CREDIAGORA\Script\start.py --exportacao receita
```

Rodar em modo invisível:

```powershell
py -3 .\CREDIAGORA\Script\start.py --headless
```

## Cuidados

- Feche as planilhas antes de executar, ou o script aguardará o Excel/OneDrive liberar os arquivos.
- Use uma pasta de download dedicada contendo `crediagora` no caminho. O script bloqueia limpeza automática em pastas amplas como `Downloads` ou `Desktop`.
- Em caso de erro, confira `CREDIAGORA/logs` e `CREDIAGORA/erros`.
- Quando o portal informar bloqueio ou última tentativa, o script bloqueia novos envios localmente. Depois que o administrador desbloquear a conta, execute `CREDIAGORA/Script/liberar_login.bat` uma única vez.
- A fase remota exige Windows com sessão de usuário desbloqueada e resolução estável. Não mova ou minimize a janela durante os diálogos RemoteApp.
