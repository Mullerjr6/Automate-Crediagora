$ErrorActionPreference = "Stop"

function Ler-TextoObrigatorio([string]$rotulo) {
    do {
        $valor = Read-Host $rotulo
    } while ([string]::IsNullOrWhiteSpace($valor))
    return $valor.Trim()
}

function Ler-SenhaObrigatoria([string]$rotulo) {
    do {
        $segura = Read-Host $rotulo -AsSecureString
        $ponteiro = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($segura)
        try {
            $valor = [Runtime.InteropServices.Marshal]::PtrToStringBSTR($ponteiro)
        }
        finally {
            [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($ponteiro)
        }
    } while ([string]::IsNullOrWhiteSpace($valor))
    return $valor
}

Write-Host "Configuração segura do ERCard" -ForegroundColor Cyan
Write-Host "As senhas não serão exibidas na tela."

$portalUsuario = Ler-TextoObrigatorio "Usuário do portal ER Systems"
$portalSenha = Ler-SenhaObrigatoria "Senha do portal ER Systems"
$sistemaUsuario = Ler-TextoObrigatorio "Usuário do ER Card"
$sistemaSenha = Ler-SenhaObrigatoria "Senha do ER Card"

[Environment]::SetEnvironmentVariable("ERCARD_PORTAL_USUARIO", $portalUsuario, "User")
[Environment]::SetEnvironmentVariable("ERCARD_PORTAL_SENHA", $portalSenha, "User")
[Environment]::SetEnvironmentVariable("ERCARD_SISTEMA_USUARIO", $sistemaUsuario, "User")
[Environment]::SetEnvironmentVariable("ERCARD_SISTEMA_SENHA", $sistemaSenha, "User")

Write-Host "Credenciais ERCard registradas no perfil do Windows." -ForegroundColor Green
