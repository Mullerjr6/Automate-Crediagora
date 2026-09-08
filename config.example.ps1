# Copie os valores necessários para o seu PowerShell antes de executar a automação.
# Para salvar permanentemente no Windows, use setx em vez de $env:.

$env:CREDIAGORA_USUARIO = "junior.muller"
$env:CREDIAGORA_SENHA = "troque-pela-senha-real"

# Opcional: ajuste apenas se os caminhos mudarem nesta máquina.
$env:CREDIAGORA_DOWNLOAD_DIR = "$HOME\Downloads\crediagora"
$env:CREDIAGORA_TABELA_FAT_DIR = "$HOME\TJI PROMOTORA DE VENDAS EIRELI\Crediagora-doc - dados\tabela fat"

# Opcionais de execução.
$env:CREDIAGORA_HEADLESS = "0"
$env:CREDIAGORA_MANTER_DOWNLOADS = "0"
$env:CREDIAGORA_MANTER_NAVEGADOR = "0"
$env:CREDIAGORA_DOWNLOAD_TIMEOUT = "300"
$env:CREDIAGORA_EXCEL_TIMEOUT = "120"

# Credenciais da fase ERCard. Nunca envie os valores reais ao Git.
$env:ERCARD_PORTAL_USUARIO = "usuario-do-portal"
$env:ERCARD_PORTAL_SENHA = "senha-do-portal"
$env:ERCARD_SISTEMA_USUARIO = "usuario-do-er-card"
$env:ERCARD_SISTEMA_SENHA = "senha-do-er-card"

# Opcionais da fase ERCard.
$env:ERCARD_EXPORT_DIR = "$HOME\Desktop\Exportações"
$env:ERCARD_TIMEOUT_NORMAL = "30"
$env:ERCARD_TIMEOUT_REMOTO = "120"
$env:ERCARD_TIMEOUT_EXPORTACAO = "180"
