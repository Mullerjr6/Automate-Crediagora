# Automação Crediagora

Automação em Python para acessar o Crediagora, exportar as bases de **vendas** e **receita gerada** e atualizar as planilhas `fat` no OneDrive/SharePoint.

## Estrutura

- `CREDIAGORA/Script/start.py`: automação principal.
- `CREDIAGORA/Script/iniciar.bat`: executa a automação no Windows e usa `.venv` automaticamente quando existir.
- `CREDIAGORA/Script/diagnostico.bat`: valida configuração sem abrir Chrome.
- `CREDIAGORA/Script/liberar_login.bat`: libera a proteção local após o desbloqueio da conta no portal.
- `CREDIAGORA/Script/limpar_downloads_crediagora.py`: limpa a pasta de downloads da automação.
- `CREDIAGORA/backups/`: backups gerados antes de atualizar planilhas.
- `CREDIAGORA/logs/`: logs de execução.
- `CREDIAGORA/erros/`: prints e HTML salvos em falhas.

## Instalação

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

## Uso

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
