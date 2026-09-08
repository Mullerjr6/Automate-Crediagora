import argparse
import hashlib
import json
import os
import sys
import threading
import time
import shutil
import re
import unicodedata
from contextlib import contextmanager
from pathlib import Path
from datetime import datetime

import pandas as pd
from openpyxl import load_workbook
from openpyxl.formula.translate import Translator
from openpyxl.utils import get_column_letter, range_boundaries
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

from ercard import ErCardConfig, executar_fase_ercard


# ============================================================
# CONFIGURAÇÕES PRINCIPAIS
# ============================================================

def ler_variavel_perfil_windows(nome):
    """Lê o valor bruto do perfil do usuário sem passar pelo interpretador .bat."""
    if os.name != "nt":
        return None

    try:
        import winreg

        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, "Environment") as chave:
            valor, _ = winreg.QueryValueEx(chave, nome)
            return str(valor)
    except (FileNotFoundError, OSError):
        return None


def carregar_credencial(nome, padrao=""):
    valor_perfil = ler_variavel_perfil_windows(nome)
    if valor_perfil is not None:
        return valor_perfil.strip(), "perfil do Windows"

    valor_ambiente = os.getenv(nome)
    if valor_ambiente is not None:
        return valor_ambiente.strip(), "ambiente do processo"

    return str(padrao).strip(), "valor padrão"

URL = os.getenv(
    "CREDIAGORA_URL",
    "https://crediagora.panoramaemprestimos.com.br/login.do?action=sistema"
).strip()

USUARIO, FONTE_USUARIO = carregar_credencial(
    "CREDIAGORA_USUARIO", "junior.muller"
)

# Coloque a senha como variável de ambiente no Windows:
# setx CREDIAGORA_SENHA "SUA_SENHA_AQUI"
#
# Depois feche e abra o CMD novamente.
SENHA, FONTE_SENHA = carregar_credencial("CREDIAGORA_SENHA")

ERCARD_PORTAL_USUARIO, _ = carregar_credencial("ERCARD_PORTAL_USUARIO")
ERCARD_PORTAL_SENHA, _ = carregar_credencial("ERCARD_PORTAL_SENHA")
ERCARD_SISTEMA_USUARIO, _ = carregar_credencial("ERCARD_SISTEMA_USUARIO")
ERCARD_SISTEMA_SENHA, _ = carregar_credencial("ERCARD_SISTEMA_SENHA")


def caminho_env(nome_variavel, padrao):
    valor = os.getenv(nome_variavel)
    return Path(valor).expanduser() if valor else Path(padrao).expanduser()


def bool_env(nome_variavel, padrao=False):
    valor = os.getenv(nome_variavel)
    if valor is None:
        return padrao

    texto = unicodedata.normalize("NFD", str(valor))
    texto = "".join(
        caractere
        for caractere in texto
        if unicodedata.category(caractere) != "Mn"
    ).lower().strip()

    return texto in {"1", "s", "sim", "true", "yes", "y", "on"}


def int_env(nome_variavel, padrao, minimo=1):
    valor = os.getenv(nome_variavel)

    if valor is None or not str(valor).strip():
        return padrao

    try:
        numero = int(valor)
    except ValueError as erro:
        raise ValueError(f"Variável {nome_variavel} precisa ser um número inteiro.") from erro

    if numero < minimo:
        raise ValueError(f"Variável {nome_variavel} precisa ser maior ou igual a {minimo}.")

    return numero

# Pasta onde os arquivos exportados serão baixados
PASTA_DOWNLOAD = caminho_env(
    "CREDIAGORA_DOWNLOAD_DIR",
    Path.home() / "Downloads" / "crediagora"
)

# Caminho base da tabela fat no SharePoint/OneDrive
PASTA_TABELA_FAT = caminho_env(
    "CREDIAGORA_TABELA_FAT_DIR",
    Path.home()
    / "TJI PROMOTORA DE VENDAS EIRELI"
    / "Crediagora-doc - dados"
    / "tabela fat"
)

# Arquivos finais
ARQUIVO_FAT_VENDAS = PASTA_TABELA_FAT / "fat_vendas_Teste.xlsx"

# O script tenta encontrar a receita em alguns caminhos possíveis:
ARQUIVOS_FAT_RECEITA_POSSIVEIS = [
    PASTA_TABELA_FAT / "fat_receita_gerada_CPC_Teste.xlsx",
    PASTA_TABELA_FAT / "fat_receita_gerada_CPC_TESTE.xlsx",
    PASTA_TABELA_FAT / "tabelas fat" / "fat_receita_gerada_CPC_Teste.xlsx",
    PASTA_TABELA_FAT / "tabelas fat" / "fat_receita_gerada_CPC_TESTE.xlsx",
]

ARQUIVO_FAT_RECEITA = None

for caminho_receita in ARQUIVOS_FAT_RECEITA_POSSIVEIS:
    if caminho_receita.exists():
        ARQUIVO_FAT_RECEITA = caminho_receita
        break

if ARQUIVO_FAT_RECEITA is None:
    ARQUIVO_FAT_RECEITA = ARQUIVOS_FAT_RECEITA_POSSIVEIS[0]

# Coluna da data cpc no arquivo final
COLUNA_DATA_CPC = "L"

# Configuração de colunas por base
# Vendas: dados de A até T; fórmulas em U e V
COLUNAS_FORMULA_VENDAS = ["U", "V"]
ULTIMA_COLUNA_DADOS_VENDAS = "T"
ULTIMA_COLUNA_TOTAL_VENDAS = "V"

# Receita gerada: dados de A até V; fórmulas em W, X e Y
COLUNAS_FORMULA_RECEITA = ["W", "X", "Y"]
ULTIMA_COLUNA_DADOS_RECEITA = "V"
ULTIMA_COLUNA_TOTAL_RECEITA = "Y"

# O script fica na subpasta Script; por padrão, os artefatos pertencem à pasta CREDIAGORA.
PASTA_SCRIPT = Path(__file__).resolve().parent
PASTA_AUTOMACAO = caminho_env("CREDIAGORA_RUNTIME_DIR", PASTA_SCRIPT.parent)
PASTA_PERFIL_CHROME = caminho_env(
    "CREDIAGORA_CHROME_PROFILE_DIR",
    PASTA_AUTOMACAO / "chrome-profile"
)
PASTA_BACKUP = PASTA_AUTOMACAO / "backups"
PASTA_LOGS = PASTA_AUTOMACAO / "logs"
PASTA_ERROS = PASTA_AUTOMACAO / "erros"
ARQUIVO_ESTADO_LOGIN = PASTA_AUTOMACAO / "estado_login.json"

# Por padrão, remove arquivos exportados depois do processo. Defina como 1/sim
# se quiser manter os arquivos baixados para conferência manual.
MANTER_DOWNLOADS = bool_env("CREDIAGORA_MANTER_DOWNLOADS", False)
MANTER_NAVEGADOR = bool_env("CREDIAGORA_MANTER_NAVEGADOR", False)
MODO_HEADLESS = bool_env("CREDIAGORA_HEADLESS", False)
TIMEOUT_DOWNLOAD = int_env("CREDIAGORA_DOWNLOAD_TIMEOUT", 300)
TIMEOUT_EXCEL = int_env("CREDIAGORA_EXCEL_TIMEOUT", 120)
ERCARD_TIMEOUT_NORMAL = int_env("ERCARD_TIMEOUT_NORMAL", 30)
ERCARD_TIMEOUT_REMOTO = int_env("ERCARD_TIMEOUT_REMOTO", 120)
ERCARD_TIMEOUT_EXPORTACAO = int_env("ERCARD_TIMEOUT_EXPORTACAO", 180)
PASTA_EXPORTACOES_ERCARD = caminho_env(
    "ERCARD_EXPORT_DIR",
    Path.home() / "Desktop" / "Exportações"
)

PADROES_DOWNLOAD = [
    "*.xlsx",
    "*.xls",
    "*.xlsm",
    "*.xlsb",
    "*.csv",
    "*.zip",
    "*.crdownload",
    "*.tmp",
    "*.download",
    "*.part",
]

EXTENSOES_TEMPORARIAS_DOWNLOAD = (".crdownload", ".tmp", ".download", ".part")
EXTENSOES_VALIDAS_EXPORTACAO = (".csv", ".xlsx", ".xlsm", ".xls", ".xlsb", ".zip")

# Lista que guarda todos os logs desta execução
LOGS_ATIVACAO = []

# ============================================================
# FUNÇÕES DE APOIO
# ============================================================

def log(msg):
    agora = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
    texto = f"[{agora}] {msg}"
    print(texto)

    try:
        LOGS_ATIVACAO.append(texto)
    except Exception:
        pass


@contextmanager
def acompanhar_operacao(descricao, intervalo=10):
    """Exibe atividade periódica durante operações lentas e bloqueantes."""
    inicio = time.monotonic()
    terminou = threading.Event()

    def informar_progresso():
        while not terminou.wait(intervalo):
            decorrido = int(time.monotonic() - inicio)
            log(f"{descricao} em andamento... {decorrido}s decorridos.")

    log(f"{descricao}...")
    monitor = threading.Thread(target=informar_progresso, daemon=True)
    monitor.start()

    try:
        yield
    finally:
        terminou.set()
        monitor.join(timeout=1)
        decorrido = time.monotonic() - inicio
        log(f"{descricao} concluído em {decorrido:.1f}s.")


def criar_pastas():
    PASTA_DOWNLOAD.mkdir(parents=True, exist_ok=True)
    PASTA_BACKUP.mkdir(parents=True, exist_ok=True)
    PASTA_LOGS.mkdir(parents=True, exist_ok=True)
    PASTA_ERROS.mkdir(parents=True, exist_ok=True)
    PASTA_PERFIL_CHROME.mkdir(parents=True, exist_ok=True)


def identificador_credencial():
    conteudo = f"{USUARIO}\0{SENHA}".encode("utf-8")
    return hashlib.sha256(conteudo).hexdigest()[:16]


def ler_estado_login():
    if not ARQUIVO_ESTADO_LOGIN.exists():
        return {}

    try:
        dados = json.loads(ARQUIVO_ESTADO_LOGIN.read_text(encoding="utf-8"))
        return dados if isinstance(dados, dict) else {}
    except (OSError, ValueError, TypeError):
        return {}


def salvar_estado_login(bloqueado, motivo, tentativas_restantes=None):
    estado = {
        "bloqueado": bool(bloqueado),
        "motivo": str(motivo),
        "tentativas_restantes": tentativas_restantes,
        "credencial_id": identificador_credencial(),
        "atualizado_em": datetime.now().isoformat(timespec="seconds"),
    }
    ARQUIVO_ESTADO_LOGIN.write_text(
        json.dumps(estado, ensure_ascii=False, indent=2),
        encoding="utf-8"
    )


def validar_bloqueio_local_login():
    estado = ler_estado_login()
    mesma_credencial = estado.get("credencial_id") == identificador_credencial()

    if estado.get("bloqueado") and mesma_credencial:
        raise RuntimeError(
            "Login automático bloqueado localmente para proteger a conta. "
            f"Motivo registrado: {estado.get('motivo', 'bloqueio informado pelo portal')}. "
            "Após o administrador desbloquear a conta, execute iniciar.bat "
            "--liberar-bloqueio-login antes de tentar novamente."
        )


def liberar_bloqueio_local_login():
    try:
        ARQUIVO_ESTADO_LOGIN.unlink(missing_ok=True)
    except TypeError:
        if ARQUIVO_ESTADO_LOGIN.exists():
            ARQUIVO_ESTADO_LOGIN.unlink()


def validar_pasta_download_segura():
    caminho = PASTA_DOWNLOAD.resolve()
    home = Path.home().resolve()
    caminhos_proibidos = [
        home,
        (home / "Downloads").resolve(),
        (home / "Desktop").resolve(),
    ]

    if PASTA_DOWNLOAD.anchor:
        caminhos_proibidos.append(Path(PASTA_DOWNLOAD.anchor).resolve())

    if caminho in caminhos_proibidos:
        raise ValueError(
            "Pasta de downloads insegura para limpeza automática: "
            f"{PASTA_DOWNLOAD}. Use uma subpasta dedicada, como Downloads\\crediagora."
        )

    if "crediagora" not in str(caminho).lower():
        raise ValueError(
            "Pasta de downloads precisa ser dedicada à automação Crediagora. "
            f"Valor atual: {PASTA_DOWNLOAD}"
        )


def mostrar_resumo_configuracao(exportacao, headless, manter_downloads, manter_navegador, criar_backups):
    log("Resumo da configuração:")
    log(f"- Exportação: {exportacao}")
    log(f"- URL: {URL}")
    log(f"- Usuário: {USUARIO}")
    log(f"- Origem das credenciais: usuário={FONTE_USUARIO}; senha={FONTE_SENHA}")
    log(f"- Senha carregada: sim (comprimento: {len(SENHA)})")
    log(f"- Downloads: {PASTA_DOWNLOAD}")
    log(f"- Tabela fat: {PASTA_TABELA_FAT}")
    log(f"- Backups/logs/erros: {PASTA_AUTOMACAO}")
    log(f"- Perfil persistente do Chrome: {PASTA_PERFIL_CHROME}")
    log(f"- Timeout download: {TIMEOUT_DOWNLOAD}s")
    log(f"- Timeout Excel: {TIMEOUT_EXCEL}s")
    log(f"- Chrome headless: {'sim' if headless else 'não'}")
    log(f"- Manter navegador ao final: {'sim' if manter_navegador else 'não'}")
    log(f"- Manter downloads ao final: {'sim' if manter_downloads else 'não'}")
    log(f"- Criar backup antes de atualizar: {'sim' if criar_backups else 'não'}")


def validar_configuracao():
    if not URL.startswith(("http://", "https://")):
        raise EnvironmentError("URL inválida. Defina CREDIAGORA_URL com http:// ou https://.")

    if not USUARIO or not str(USUARIO).strip():
        raise EnvironmentError(
            "Usuário não configurado. Defina CREDIAGORA_USUARIO ou ajuste o valor padrão no script."
        )

    if not SENHA or SENHA == "COLOQUE_SUA_SENHA_AQUI":
        raise EnvironmentError(
            "Senha não configurada. Defina a variável de ambiente "
            "CREDIAGORA_SENHA antes de executar a automação."
        )

    validar_pasta_download_segura()


def criar_configuracao_ercard():
    return ErCardConfig(
        portal_usuario=ERCARD_PORTAL_USUARIO,
        portal_senha=ERCARD_PORTAL_SENHA,
        sistema_usuario=ERCARD_SISTEMA_USUARIO,
        sistema_senha=ERCARD_SISTEMA_SENHA,
        pasta_exportacao=PASTA_EXPORTACOES_ERCARD,
        pasta_erros=PASTA_ERROS,
        timeout_normal=ERCARD_TIMEOUT_NORMAL,
        timeout_remoto=ERCARD_TIMEOUT_REMOTO,
        timeout_exportacao=ERCARD_TIMEOUT_EXPORTACAO,
    )

def validar_caminhos(exportacao="todas"):
    log("Validando caminhos...")

    if not PASTA_TABELA_FAT.exists():
        raise FileNotFoundError(f"Pasta tabela fat não encontrada: {PASTA_TABELA_FAT}")

    if exportacao in ("todas", "vendas") and not ARQUIVO_FAT_VENDAS.exists():
        raise FileNotFoundError(f"Arquivo fat_vendas não encontrado: {ARQUIVO_FAT_VENDAS}")

    if exportacao in ("todas", "receita") and not ARQUIVO_FAT_RECEITA.exists():
        caminhos_tentados = "\n".join(f"- {c}" for c in ARQUIVOS_FAT_RECEITA_POSSIVEIS)
        raise FileNotFoundError(
            "Arquivo fat_receita_gerada não encontrado. "
            f"Tentei estes caminhos:\n{caminhos_tentados}"
        )

    if exportacao in ("todas", "vendas"):
        log(f"fat_vendas encontrado: {ARQUIVO_FAT_VENDAS}")

    if exportacao in ("todas", "receita"):
        log(f"fat_receita_gerada encontrado: {ARQUIVO_FAT_RECEITA}")


def limpar_downloads():
    """
    Limpa a pasta de download antes de cada exportação.
    Isso evita o Python confundir arquivo antigo com o novo.
    """
    log("Limpando pasta de downloads...")
    validar_pasta_download_segura()

    PASTA_DOWNLOAD.mkdir(parents=True, exist_ok=True)

    apagados = 0

    for padrao in PADROES_DOWNLOAD:
        for arquivo in PASTA_DOWNLOAD.glob(padrao):
            try:
                if arquivo.is_file():
                    arquivo.unlink()
                    apagados += 1
                    log(f"Arquivo removido: {arquivo.name}")
            except Exception as erro:
                log(f"Não foi possível apagar {arquivo.name}: {erro}")

    log(f"Limpeza concluída. Arquivos apagados: {apagados}")


def iniciar_chrome(headless=False, manter_navegador=False):
    opcoes = Options()

    prefs = {
        "download.default_directory": str(PASTA_DOWNLOAD),
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True,
    }

    opcoes.add_experimental_option("prefs", prefs)
    opcoes.add_argument(f"--user-data-dir={PASTA_PERFIL_CHROME}")
    opcoes.add_argument("--profile-directory=Default")
    opcoes.add_argument("--start-maximized")
    opcoes.add_argument("--disable-notifications")

    if headless:
        opcoes.add_argument("--headless=new")
        opcoes.add_argument("--window-size=1920,1080")

    if manter_navegador and not headless:
        opcoes.add_experimental_option("detach", True)

    driver = webdriver.Chrome(options=opcoes)

    try:
        driver.execute_cdp_cmd(
            "Page.setDownloadBehavior",
            {"behavior": "allow", "downloadPath": str(PASTA_DOWNLOAD)}
        )
    except Exception as erro:
        log(f"Não foi possível configurar download via CDP; seguindo com preferências do Chrome. Detalhe: {erro}")

    return driver


def esperar(driver, segundos=20):
    return WebDriverWait(driver, segundos)


def normalizar_texto(texto):
    texto = str(texto)
    texto_sem_acentos = unicodedata.normalize("NFD", texto)
    return "".join(
        caractere
        for caractere in texto_sem_acentos
        if unicodedata.category(caractere) != "Mn"
    ).lower()


def xpath_literal(texto):
    """
    Monta texto seguro para usar dentro de XPath mesmo quando tiver aspas.
    """
    texto = str(texto)

    if "'" not in texto:
        return f"'{texto}'"

    if '"' not in texto:
        return f'"{texto}"'

    partes = texto.split("'")
    return "concat(" + ', "\"\'\"", '.join(f"'{parte}'" for parte in partes) + ")"


def clicar_elemento_seguro(driver, elemento, descricao="elemento"):
    """
    Faz scroll, tenta mover o mouse e clica.
    Se o clique normal for interceptado, força clique via JavaScript.
    """
    try:
        driver.execute_script(
            "arguments[0].scrollIntoView({block: 'center', inline: 'center'});",
            elemento
        )
        time.sleep(0.4)
    except Exception:
        pass

    try:
        ActionChains(driver).move_to_element(elemento).pause(0.2).perform()
    except Exception:
        pass

    try:
        elemento.click()
    except Exception:
        driver.execute_script("arguments[0].click();", elemento)

    log(f"Clique realizado em: {descricao}")


def clicar_por_texto(driver, texto, segundos=20):
    """
    Clica em menus/botões pelo texto ou por atributos.

    Correção principal:
    - Prioriza elementos realmente clicáveis: a, button, input, div/li/span com onclick/url/descricao.
    - Procura no conteúdo principal e também em iframes.
    - Evita clicar em containers grandes, como body/div pai, que travavam o menu de Empréstimo.
    - Mantém o Selenium no contexto onde clicou, útil quando o próximo menu abre dentro do mesmo iframe.
    """
    texto_original = str(texto)
    texto_busca = normalizar_texto(texto_original)
    xpath_literal(texto_original)

    # Para "Empréstimo", o href/ação costuma vir sem acento.
    texto_sem_acento_xpath = xpath_literal(texto_busca)

    caracteres = "ÁÀÃÂÉÊÍÓÔÕÚÇáàãâéêíóôõúçABCDEFGHIJKLMNOPQRSTUVWXYZ"
    substitutos = "AAAAEEIOOOUCaaaaeeioooucabcdefghijklmnopqrstuvwxyz"

    xpaths = [
        # 1) Links/ações por href, url, ação ou onclick
        f"//a[contains(translate(@href, 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), {texto_sem_acento_xpath})]",
        f"//*[contains(translate(@url, 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), {texto_sem_acento_xpath})]",
        f"//*[contains(translate(@acao, 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), {texto_sem_acento_xpath})]",
        f"//*[contains(translate(@onclick, 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), {texto_sem_acento_xpath})]",

        # 2) Atributos de descrição/title
        f"//*[contains(translate(@descricao, '{caracteres}', '{substitutos}'), {texto_sem_acento_xpath})]",
        f"//*[contains(translate(@title, '{caracteres}', '{substitutos}'), {texto_sem_acento_xpath})]",

        # 3) Elementos clicáveis pelo texto
        f"//a[contains(translate(normalize-space(.), '{caracteres}', '{substitutos}'), {texto_sem_acento_xpath})]",
        f"//button[contains(translate(normalize-space(.), '{caracteres}', '{substitutos}'), {texto_sem_acento_xpath})]",
        f"//input[contains(translate(@value, '{caracteres}', '{substitutos}'), {texto_sem_acento_xpath})]",
        f"//div[(contains(@class,'btn') or @onclick or @url or @descricao or @tipo_btn) and contains(translate(normalize-space(.), '{caracteres}', '{substitutos}'), {texto_sem_acento_xpath})]",
        f"//li[(@onclick or .//a) and contains(translate(normalize-space(.), '{caracteres}', '{substitutos}'), {texto_sem_acento_xpath})]",
        f"//span[(@onclick or ancestor::a or ancestor::*[@onclick]) and contains(translate(normalize-space(.), '{caracteres}', '{substitutos}'), {texto_sem_acento_xpath})]",

        # 4) Fallback mais amplo, mas ainda evitando body/html
        f"//*[not(self::html) and not(self::body) and string-length(normalize-space(.)) < 80 and contains(translate(normalize-space(.), '{caracteres}', '{substitutos}'), {texto_sem_acento_xpath})]",
    ]

    def tentar_no_contexto(nome_contexto):
        ultimo_erro_local = None

        for xpath in xpaths:
            try:
                elementos = driver.find_elements(By.XPATH, xpath)

                # Filtra somente elementos visíveis.
                elementos_visiveis = []
                for elemento in elementos:
                    try:
                        if elemento.is_displayed():
                            elementos_visiveis.append(elemento)
                    except Exception:
                        continue

                if not elementos_visiveis:
                    continue

                # Prefere elementos menores/clicáveis, evitando container pai.
                elementos_visiveis.sort(
                    key=lambda el: (
                        0 if (el.tag_name or "").lower() in ["a", "button", "input"] else 1,
                        len((el.text or el.get_attribute("value") or "").strip())
                    )
                )

                elemento = elementos_visiveis[0]
                clicar_elemento_seguro(driver, elemento, f"{texto_original} em {nome_contexto}")
                time.sleep(1)
                return True

            except Exception as erro:
                ultimo_erro_local = erro
                continue

        return False

    fim = time.time() + segundos
    ultimo_erro = None

    while time.time() < fim:
        # Conteúdo principal
        try:
            driver.switch_to.default_content()
            if tentar_no_contexto("conteúdo principal"):
                return True
        except Exception as erro:
            ultimo_erro = erro

        # Iframes
        try:
            driver.switch_to.default_content()
            iframes = driver.find_elements(By.TAG_NAME, "iframe")

            for indice, iframe in enumerate(iframes):
                try:
                    driver.switch_to.default_content()
                    driver.switch_to.frame(iframe)

                    if tentar_no_contexto(f"iframe {indice}"):
                        return True

                except Exception as erro:
                    ultimo_erro = erro
                    continue

        except Exception as erro:
            ultimo_erro = erro

        time.sleep(1)

    salvar_diagnostico_erro(driver, f"erro_clicar_{texto_original}")
    raise Exception(
        f"Não consegui clicar em '{texto_original}' dentro de {segundos} segundos. "
        f"Último erro: {ultimo_erro}"
    )


def preencher_campo_por_xpath(driver, xpath, texto, segundos=20, limpar=True):
    campo = esperar(driver, segundos).until(
        EC.element_to_be_clickable((By.XPATH, xpath))
    )

    if limpar:
        campo.clear()

    campo.send_keys(texto)
    return campo


def esperar_download_novo(pasta, inicio=None, timeout=None):
    pasta = Path(pasta)
    timeout = TIMEOUT_DOWNLOAD if timeout is None else timeout
    pasta.mkdir(parents=True, exist_ok=True)

    if inicio is None:
        inicio = time.time() - 10

    def listar_temporarios():
        temporarios = []

        for arquivo in pasta.iterdir():
            try:
                nome = arquivo.name.lower()
                if arquivo.is_file() and nome.endswith(EXTENSOES_TEMPORARIAS_DOWNLOAD):
                    temporarios.append(arquivo)
            except Exception:
                continue

        return temporarios

    def listar_validos():
        arquivos = []

        for arquivo in pasta.iterdir():
            try:
                if not arquivo.is_file():
                    continue

                nome = arquivo.name.lower()

                if nome.startswith("~$"):
                    continue

                if nome.endswith(EXTENSOES_TEMPORARIAS_DOWNLOAD):
                    continue

                if not nome.endswith(EXTENSOES_VALIDAS_EXPORTACAO):
                    continue

                # Aceita arquivos modificados depois do clique.
                # Como limpamos a pasta antes, também aceitamos qualquer arquivo válido que esteja sozinho na pasta.
                if arquivo.stat().st_mtime >= inicio or len(list(pasta.glob("*"))) <= 2:
                    arquivos.append(arquivo)

            except Exception:
                continue

        return arquivos

    def arquivo_estavel(arquivo, tentativas=2, intervalo=0.5):
        tamanhos = []

        for _ in range(tentativas):
            try:
                tamanhos.append(arquivo.stat().st_size)
            except Exception:
                return False

            time.sleep(intervalo)

        return len(set(tamanhos)) == 1 and tamanhos[-1] > 0

    fim = time.time() + timeout

    while time.time() < fim:
        temporarios = listar_temporarios()
        arquivos = listar_validos()

        if arquivos and not temporarios:
            arquivo_mais_recente = max(arquivos, key=lambda x: x.stat().st_mtime)

            if arquivo_estavel(arquivo_mais_recente):
                log(f"Download confirmado: {arquivo_mais_recente}")
                return arquivo_mais_recente

        time.sleep(1)

    log("Arquivos encontrados na pasta de download no momento do erro:")
    try:
        for arquivo in pasta.iterdir():
            if arquivo.is_file():
                mod = datetime.fromtimestamp(arquivo.stat().st_mtime).strftime("%d/%m/%Y %H:%M:%S")
                log(f"{arquivo.name} | tamanho: {arquivo.stat().st_size} | modificado: {mod}")
    except Exception:
        pass

    raise TimeoutError(
        "Download não encontrado dentro do tempo limite. "
        f"Verifique se o Chrome está baixando em {PASTA_DOWNLOAD}."
    )


def criar_backup(arquivo, identificador="manual"):
    """
    Cria backup do arquivo informado na pasta de backups da automação.
    """
    arquivo = Path(arquivo)

    esperar_arquivo_excel_liberar(arquivo, timeout=TIMEOUT_EXCEL)

    PASTA_BACKUP.mkdir(parents=True, exist_ok=True)

    data_hora = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    destino = PASTA_BACKUP / f"{arquivo.stem}_backup_{identificador}_{data_hora}{arquivo.suffix}"

    shutil.copy2(arquivo, destino)
    log(f"Backup criado: {destino}")

    return destino


def criar_backups_iniciais(exportacao="todas"):
    """
    Cria backup dos arquivos antigos antes de iniciar o processo no site.

    Arquivos copiados para a pasta backups:
    - Exportação vendas: fat_vendas_Teste.xlsx
    - Receita gerada: fat_receita_gerada_CPC_Teste.xlsx / TESTE.xlsx
    """
    log("Criando backups dos arquivos atuais antes da atualização...")

    arquivos_para_backup = []

    if exportacao in ("todas", "vendas"):
        arquivos_para_backup.append(("exportacao_vendas", ARQUIVO_FAT_VENDAS))

    if exportacao in ("todas", "receita"):
        arquivos_para_backup.append(("receita_gerada", ARQUIVO_FAT_RECEITA))

    for nome, arquivo in arquivos_para_backup:
        arquivo = Path(arquivo)

        if not arquivo.exists():
            raise FileNotFoundError(f"Arquivo para backup não encontrado: {arquivo}")

        criar_backup(arquivo, identificador=nome)

    log("Backups dos arquivos selecionados criados com sucesso.")


def salvar_log_ativacao(status="execucao"):
    """
    Salva um arquivo TXT com todo o log gerado nesta execução.
    """
    try:
        PASTA_LOGS.mkdir(parents=True, exist_ok=True)

        data_hora = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        arquivo_log = PASTA_LOGS / f"log_automacao_crediagora_{status}_{data_hora}.txt"

        with open(arquivo_log, "w", encoding="utf-8") as arquivo:
            arquivo.write("\n".join(LOGS_ATIVACAO))
            arquivo.write("\n")

        print(f"Log da execução salvo em: {arquivo_log}")
        return arquivo_log

    except Exception as erro:
        print(f"Não foi possível salvar o log da execução: {erro}")
        return None



def salvar_diagnostico_erro(driver, nome_base):
    """
    Salva print e HTML de diagnóstico na pasta de erros da automação.
    """
    PASTA_ERROS.mkdir(parents=True, exist_ok=True)

    data_hora = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    nome_limpo = re.sub(r"[^a-zA-Z0-9_\\-]+", "_", str(nome_base)).strip("_")

    screenshot = PASTA_ERROS / f"{nome_limpo}_{data_hora}.png"
    html = PASTA_ERROS / f"{nome_limpo}_{data_hora}.html"

    try:
        driver.save_screenshot(str(screenshot))
        log(f"Print do erro salvo em: {screenshot}")
    except Exception as erro:
        log(f"Não foi possível salvar print do erro: {erro}")

    try:
        html.write_text(driver.page_source, encoding="utf-8")
        log(f"HTML do erro salvo em: {html}")
    except Exception as erro:
        log(f"Não foi possível salvar HTML do erro: {erro}")

    return screenshot, html

# ============================================================
# LOGIN E NAVEGAÇÃO
# ============================================================

def fazer_login(driver):
    log("Abrindo site...")
    driver.get(URL)

    time.sleep(1)
    campos_senha = driver.find_elements(By.CSS_SELECTOR, "#formloginid input[type='password']")
    texto_inicial = normalizar_texto(
        driver.find_element(By.TAG_NAME, "body").text
    ).lower()

    if not campos_senha and "vendas" in texto_inicial:
        log("Sessão válida encontrada no perfil persistente; novo login dispensado.")
        salvar_estado_login(False, "sessão válida reutilizada")
        return

    validar_bloqueio_local_login()

    log("Preenchendo usuário...")
    campo_usuario = preencher_campo_por_xpath(
        driver,
        "//form[@id='formloginid']//input[@id='usuario_id_campo' or @name='login']",
        USUARIO
    )

    log("Preenchendo senha...")
    campo_senha = preencher_campo_por_xpath(
        driver,
        "//form[@id='formloginid']//input[@id='senha_id_campo' or @name='senha']",
        SENHA
    )

    usuario_preenchido = campo_usuario.get_attribute("value") or ""
    senha_preenchida = campo_senha.get_attribute("value") or ""

    if usuario_preenchido != USUARIO:
        raise RuntimeError(
            "O campo de usuário não recebeu o login completo. "
            "O envio foi cancelado para preservar as tentativas restantes."
        )

    if senha_preenchida != SENHA:
        raise RuntimeError(
            "O campo de senha não recebeu o valor completo. "
            "O envio foi cancelado para preservar as tentativas restantes."
        )

    log(
        "Campos de login conferidos antes do envio: "
        f"usuário com {len(usuario_preenchido)} caracteres; "
        f"senha com {len(senha_preenchida)} caracteres."
    )

    log("Tentando entrar no sistema...")

    botao_entrar = esperar(driver, 10).until(
        EC.element_to_be_clickable(
            (By.CSS_SELECTOR, "#formloginid a.entrar")
        )
    )
    clicar_elemento_seguro(driver, botao_entrar, "botão Entrar do formulário de login")
    log("Formulário de login enviado uma única vez.")

    log("Aguardando tela principal...")

    def identificar_resultado_login(navegador):
        texto_pagina = normalizar_texto(
            navegador.find_element(By.TAG_NAME, "body").text
        ).lower()

        correspondencia_tentativas = re.search(
            r"tentativas? restantes?\s*(\d+)", texto_pagina
        )
        tentativas_restantes = (
            int(correspondencia_tentativas.group(1))
            if correspondencia_tentativas else None
        )

        if "usuario bloqueado" in texto_pagina or "acesso bloqueado" in texto_pagina:
            salvar_estado_login(True, "usuário bloqueado pelo portal", tentativas_restantes)
            return "erro", "usuário bloqueado pelo portal"

        if tentativas_restantes is not None and tentativas_restantes <= 1:
            salvar_estado_login(True, "última tentativa disponível", tentativas_restantes)
            return "erro", "última tentativa disponível; novos envios foram bloqueados"

        mensagens_erro = (
            "usuario nao existe",
            "senha incorreta",
            "usuario ou senha invalido",
            "login ou senha invalido",
            "tentativas restantes",
        )

        for mensagem in mensagens_erro:
            if mensagem in texto_pagina:
                return "erro", mensagem

        if "vendas" in texto_pagina:
            salvar_estado_login(False, "login confirmado")
            return "sucesso", ""

        return False

    try:
        resultado, mensagem = esperar(driver, 60).until(identificar_resultado_login)

        if resultado == "erro":
            raise RuntimeError(
                "O portal recusou o login: "
                f"{mensagem}. Confira as credenciais antes de tentar novamente."
            )

        log("Login realizado com sucesso.")
    except Exception as erro:
        screenshot = PASTA_ERROS / f"erro_login_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.png"
        try:
            driver.save_screenshot(str(screenshot))
            log(f"Print do erro salvo em: {screenshot}")
        except Exception:
            pass

        raise Exception(
            "Não consegui confirmar o login. "
            "Os campos e o botão foram conferidos. O portal recusou a autenticação; "
            "verifique se a conta está bloqueada ou presa em outra sessão. "
            f"Erro original: {erro}"
        )


def encerrar_sessao_portal(driver):
    """Tenta encerrar a sessão no servidor antes de fechar o navegador."""
    seletores = (
        "//a[contains(translate(@href, 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'logout')]",
        "//a[contains(translate(@href, 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'sair')]",
        "//a[translate(normalize-space(.), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz')='sair']",
        "//button[translate(normalize-space(.), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz')='sair']",
    )

    try:
        driver.switch_to.default_content()
        contextos = [None] + driver.find_elements(By.TAG_NAME, "iframe")

        for contexto in contextos:
            driver.switch_to.default_content()
            if contexto is not None:
                try:
                    driver.switch_to.frame(contexto)
                except Exception:
                    continue

            for seletor in seletores:
                for elemento in driver.find_elements(By.XPATH, seletor):
                    try:
                        if elemento.is_displayed() and elemento.is_enabled():
                            clicar_elemento_seguro(driver, elemento, "Sair do portal")
                            log("Sessão do portal encerrada antes de fechar o navegador.")
                            time.sleep(1)
                            return True
                    except Exception:
                        continue
    except Exception as erro:
        log(f"Não foi possível procurar a opção de saída do portal: {erro}")
    finally:
        try:
            driver.switch_to.default_content()
        except Exception:
            pass

    log("Opção de saída não encontrada; o navegador será fechado normalmente.")
    return False


def entrar_em_vendas_e_emprestimo(driver):
    """
    Entra corretamente em:
    Vendas > Empréstimo

    Correção:
    - Clica/abre Vendas com clique seguro.
    - Tenta abrir o submenu com hover também.
    - Procura Empréstimo por texto, href, url, ação, descrição e onclick.
    - Depois confirma que a tela correta carregou procurando o botão acoes_ver.
    """
    log("Abrindo menu Vendas...")

    try:
        driver.switch_to.default_content()
    except Exception:
        pass

    # 1) Abre o menu Vendas.
    clicar_por_texto(driver, "Vendas", segundos=40)
    time.sleep(1.5)

    # Às vezes o submenu abre por hover e não por clique.
    try:
        elemento_vendas = None

        for xp in [
            "//a[contains(translate(normalize-space(.), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'vendas')]",
            "//*[contains(translate(@descricao, 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'vendas')]",
            "//*[contains(translate(@title, 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'vendas')]",
            "//*[string-length(normalize-space(.)) < 80 and contains(translate(normalize-space(.), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'vendas')]",
        ]:
            try:
                driver.switch_to.default_content()
                encontrados = driver.find_elements(By.XPATH, xp)
                encontrados = [e for e in encontrados if e.is_displayed()]

                if encontrados:
                    elemento_vendas = encontrados[0]
                    break
            except Exception:
                continue

        if elemento_vendas is not None:
            ActionChains(driver).move_to_element(elemento_vendas).pause(0.8).perform()
            log("Hover realizado sobre Vendas para abrir submenu.")
            time.sleep(1)
    except Exception as erro:
        log(f"Hover em Vendas não foi necessário ou falhou: {erro}")

    # 2) Clica em Empréstimo.
    log("Clicando em Empréstimo...")

    clicou_emprestimo = False
    ultimo_erro = None

    # Seletores mais específicos antes do clique genérico.
    seletores_emprestimo = [
        (By.XPATH, "//a[contains(translate(@href, 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'emprestimo')]"),
        (By.XPATH, "//*[contains(translate(@url, 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'emprestimo')]"),
        (By.XPATH, "//*[contains(translate(@acao, 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'emprestimo')]"),
        (By.XPATH, "//*[contains(translate(@onclick, 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'emprestimo')]"),
        (By.XPATH, "//*[contains(translate(@descricao, 'ÁÀÃÂÉÊÍÓÔÕÚÇáàãâéêíóôõúçABCDEFGHIJKLMNOPQRSTUVWXYZ', 'AAAAEEIOOOUCaaaaeeioooucabcdefghijklmnopqrstuvwxyz'), 'emprestimo')]"),
        (By.XPATH, "//a[contains(translate(normalize-space(.), 'ÁÀÃÂÉÊÍÓÔÕÚÇáàãâéêíóôõúçABCDEFGHIJKLMNOPQRSTUVWXYZ', 'AAAAEEIOOOUCaaaaeeioooucabcdefghijklmnopqrstuvwxyz'), 'emprestimo')]"),
        (By.XPATH, "//li[contains(translate(normalize-space(.), 'ÁÀÃÂÉÊÍÓÔÕÚÇáàãâéêíóôõúçABCDEFGHIJKLMNOPQRSTUVWXYZ', 'AAAAEEIOOOUCaaaaeeioooucabcdefghijklmnopqrstuvwxyz'), 'emprestimo')]"),
        (By.XPATH, "//*[string-length(normalize-space(.)) < 80 and contains(translate(normalize-space(.), 'ÁÀÃÂÉÊÍÓÔÕÚÇáàãâéêíóôõúçABCDEFGHIJKLMNOPQRSTUVWXYZ', 'AAAAEEIOOOUCaaaaeeioooucabcdefghijklmnopqrstuvwxyz'), 'emprestimo')]"),
    ]

    def tentar_clicar_emprestimo_contexto(nome_contexto):
        nonlocal ultimo_erro

        for by, seletor in seletores_emprestimo:
            try:
                elementos = driver.find_elements(by, seletor)
                elementos = [e for e in elementos if e.is_displayed()]

                if not elementos:
                    continue

                elementos.sort(
                    key=lambda el: (
                        0 if (el.tag_name or "").lower() in ["a", "button"] else 1,
                        len((el.text or "").strip())
                    )
                )

                clicar_elemento_seguro(driver, elementos[0], f"Empréstimo em {nome_contexto}")
                time.sleep(2)
                return True

            except Exception as erro:
                ultimo_erro = erro
                continue

        return False

    # Tenta no contexto atual primeiro, pois o submenu pode ter aberto dentro dele.
    try:
        if tentar_clicar_emprestimo_contexto("contexto atual"):
            clicou_emprestimo = True
    except Exception as erro:
        ultimo_erro = erro

    # Tenta no conteúdo principal e nos iframes.
    if not clicou_emprestimo:
        fim = time.time() + 35

        while time.time() < fim and not clicou_emprestimo:
            try:
                driver.switch_to.default_content()
                if tentar_clicar_emprestimo_contexto("conteúdo principal"):
                    clicou_emprestimo = True
                    break
            except Exception as erro:
                ultimo_erro = erro

            try:
                driver.switch_to.default_content()
                iframes = driver.find_elements(By.TAG_NAME, "iframe")
                log(f"Procurando Empréstimo em iframes. Total: {len(iframes)}")

                for indice, iframe in enumerate(iframes):
                    try:
                        driver.switch_to.default_content()
                        driver.switch_to.frame(iframe)

                        if tentar_clicar_emprestimo_contexto(f"iframe {indice}"):
                            clicou_emprestimo = True
                            break

                    except Exception as erro:
                        ultimo_erro = erro
                        continue

            except Exception as erro:
                ultimo_erro = erro

            if not clicou_emprestimo:
                time.sleep(1)

    # Último fallback: usa a função genérica corrigida.
    if not clicou_emprestimo:
        try:
            clicar_por_texto(driver, "Empréstimo", segundos=15)
            clicou_emprestimo = True
        except Exception as erro:
            ultimo_erro = erro
            log(f"Fallback por texto também falhou para Empréstimo: {erro}")

    if not clicou_emprestimo:
        salvar_diagnostico_erro(driver, "erro_nao_achou_emprestimo")
        raise Exception(
            "Não consegui clicar no botão/menu Empréstimo. "
            f"Último erro: {ultimo_erro}. "
            "Foi salvo print e HTML na pasta de erros."
        )

    # 3) Confirma que a tela de Empréstimo carregou buscando o botão Click para ações.
    log("Confirmando carregamento da tela de Empréstimo...")

    encontrou_acoes = False
    fim = time.time() + 60

    while time.time() < fim and not encontrou_acoes:
        try:
            driver.switch_to.default_content()
            WebDriverWait(driver, 3).until(
                EC.presence_of_element_located((By.ID, "acoes_ver"))
            )
            encontrou_acoes = True
            log("Tela de Empréstimo confirmada no conteúdo principal.")
            break
        except Exception:
            pass

        try:
            driver.switch_to.default_content()
            iframes = driver.find_elements(By.TAG_NAME, "iframe")

            for indice, iframe in enumerate(iframes):
                try:
                    driver.switch_to.default_content()
                    driver.switch_to.frame(iframe)

                    WebDriverWait(driver, 3).until(
                        EC.presence_of_element_located((By.ID, "acoes_ver"))
                    )

                    encontrou_acoes = True
                    log(f"Tela de Empréstimo confirmada no iframe {indice}.")
                    break

                except Exception:
                    continue

        except Exception:
            pass

        if not encontrou_acoes:
            time.sleep(1)

    if not encontrou_acoes:
        salvar_diagnostico_erro(driver, "erro_emprestimo_nao_carregou")
        raise Exception(
            "Cliquei em Empréstimo, mas não consegui confirmar a tela. "
            "O botão 'Click para ações' não apareceu."
        )

    log("Tela de Empréstimo carregada com sucesso.")


def voltar_para_emprestimos(driver):
    """
    Depois que termina a primeira exportação, não volta para Vendas > Empréstimo.

    O fluxo correto é:
    1. Voltar uma página, se o navegador tiver ido para a tela do arquivo/download;
    2. Confirmar que o botão "Click para ações" está disponível novamente;
    3. Seguir o mesmo processo para "Receita gerada".

    Assim evita procurar "Vendas" novamente no final.
    """
    log("Voltando para a tela anterior para iniciar a próxima exportação...")

    try:
        driver.switch_to.default_content()
    except Exception:
        pass

    # Primeiro tenta voltar uma página, pois após clicar no arquivo o site pode ter navegado.
    try:
        driver.back()
        time.sleep(3)
        log("Navegador voltou uma página.")
    except Exception as erro:
        log(f"Não consegui voltar página pelo navegador: {erro}")

    # Confirma se o botão Click para ações apareceu.
    encontrou_acoes = False

    try:
        driver.switch_to.default_content()
    except Exception:
        pass

    # Procura no conteúdo principal
    try:
        WebDriverWait(driver, 8).until(
            EC.presence_of_element_located((By.ID, "acoes_ver"))
        )
        encontrou_acoes = True
        log("Botão Click para ações encontrado no conteúdo principal.")
    except Exception:
        pass

    # Procura dentro dos iframes
    if not encontrou_acoes:
        try:
            driver.switch_to.default_content()
            iframes = driver.find_elements(By.TAG_NAME, "iframe")
            log(f"Procurando Click para ações em iframes após voltar. Total: {len(iframes)}")

            for indice, iframe in enumerate(iframes):
                try:
                    driver.switch_to.default_content()
                    driver.switch_to.frame(iframe)

                    WebDriverWait(driver, 8).until(
                        EC.presence_of_element_located((By.ID, "acoes_ver"))
                    )

                    encontrou_acoes = True
                    log(f"Botão Click para ações encontrado no iframe {indice}.")
                    break

                except Exception:
                    continue

        except Exception as erro:
            log(f"Falha ao procurar botão Click para ações nos iframes: {erro}")

    # Se não encontrou, tenta mais uma vez voltar.
    if not encontrou_acoes:
        try:
            driver.switch_to.default_content()
            driver.back()
            time.sleep(3)
            log("Segunda tentativa de voltar página realizada.")
        except Exception as erro:
            log(f"Não consegui voltar na segunda tentativa: {erro}")

        try:
            driver.switch_to.default_content()
            iframes = driver.find_elements(By.TAG_NAME, "iframe")

            for indice, iframe in enumerate(iframes):
                try:
                    driver.switch_to.default_content()
                    driver.switch_to.frame(iframe)

                    WebDriverWait(driver, 8).until(
                        EC.presence_of_element_located((By.ID, "acoes_ver"))
                    )

                    encontrou_acoes = True
                    log(f"Botão Click para ações encontrado no iframe {indice} após segunda volta.")
                    break

                except Exception:
                    continue

        except Exception:
            pass

    if not encontrou_acoes:
        screenshot = PASTA_ERROS / f"erro_voltar_para_acoes_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.png"
        html = PASTA_ERROS / f"erro_voltar_para_acoes_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.html"

        try:
            driver.save_screenshot(str(screenshot))
            log(f"Print do erro salvo em: {screenshot}")
        except Exception:
            pass

        try:
            html.write_text(driver.page_source, encoding="utf-8")
            log(f"HTML do erro salvo em: {html}")
        except Exception:
            pass

        raise Exception(
            "Não consegui voltar para a tela onde aparece o botão 'Click para ações'. "
            "Parei para evitar clicar no lugar errado."
        )

    log("Tela pronta para iniciar a próxima exportação.")


# ============================================================
# EXPORTAÇÃO
# ============================================================



def localizar_iframe_com_elemento(driver, by, seletor, timeout_por_contexto=3):
    """
    Procura um elemento no conteúdo principal e dentro dos iframes.
    Retorna True deixando o driver no contexto onde o elemento foi encontrado.
    """
    try:
        driver.switch_to.default_content()
    except Exception:
        pass

    # Primeiro procura no conteúdo principal
    try:
        WebDriverWait(driver, timeout_por_contexto).until(
            EC.presence_of_element_located((by, seletor))
        )
        return True
    except Exception:
        pass

    # Depois procura em iframes
    try:
        driver.switch_to.default_content()
        iframes = driver.find_elements(By.TAG_NAME, "iframe")
        log(f"Iframes encontrados na tela: {len(iframes)}")

        for indice, iframe in enumerate(iframes):
            try:
                driver.switch_to.default_content()
                driver.switch_to.frame(iframe)

                WebDriverWait(driver, timeout_por_contexto).until(
                    EC.presence_of_element_located((by, seletor))
                )

                log(f"Elemento localizado no iframe {indice}: {seletor}")
                return True

            except Exception:
                continue

        driver.switch_to.default_content()
        return False

    except Exception:
        try:
            driver.switch_to.default_content()
        except Exception:
            pass
        return False


def clicar_por_texto_em_qualquer_contexto(driver, texto, segundos=40):
    """
    Clica em um texto procurando no conteúdo principal e também dentro dos iframes.
    É usado para menus como Exportações e Exportar Layout Arquivo.
    """
    if "'" in texto and '"' in texto:
        partes = texto.split("'")
        partes_literal = []

        for indice, parte in enumerate(partes):
            if parte:
                partes_literal.append(f"'{parte}'")
            if indice != len(partes) - 1:
                partes_literal.append('"\'"')

        texto_xpath = "concat(" + ", ".join(partes_literal) + ")"
    elif "'" in texto:
        texto_xpath = f'"{texto}"'
    else:
        texto_xpath = f"'{texto}'"

    xpath = f"//*[contains(normalize-space(.), {texto_xpath})]"

    fim = time.time() + segundos
    ultimo_erro = None

    while time.time() < fim:
        try:
            achou = localizar_iframe_com_elemento(driver, By.XPATH, xpath, timeout_por_contexto=2)

            if achou:
                elemento = WebDriverWait(driver, 5).until(
                    EC.element_to_be_clickable((By.XPATH, xpath))
                )

                driver.execute_script("arguments[0].scrollIntoView({block: 'center', inline: 'center'});", elemento)
                time.sleep(0.3)

                try:
                    elemento.click()
                except Exception:
                    driver.execute_script("arguments[0].click();", elemento)

                log(f"Clique realizado em '{texto}'.")
                time.sleep(1)
                return elemento

        except Exception as erro:
            ultimo_erro = erro
            time.sleep(1)

    screenshot = PASTA_ERROS / f"erro_click_{texto.replace(' ', '_')}_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.png"
    html = PASTA_ERROS / f"erro_click_{texto.replace(' ', '_')}_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.html"

    try:
        driver.save_screenshot(str(screenshot))
        log(f"Print do erro salvo em: {screenshot}")
    except Exception:
        pass

    try:
        html.write_text(driver.page_source, encoding="utf-8")
        log(f"HTML do erro salvo em: {html}")
    except Exception:
        pass

    raise Exception(f"Não consegui clicar no texto '{texto}'. Último erro: {ultimo_erro}")


def clicar_click_para_acoes(driver):
    """
    Clica no botão:
    <a title="Mostrar ações da tela " id="acoes_ver" class="ctr cursor btn">
        Click para ações
    </a>

    Importante:
    Se o botão estiver dentro de iframe, esta função deixa o Selenium
    dentro do mesmo iframe, porque o menu Exportações também costuma
    abrir nesse mesmo contexto.
    """
    log("Abrindo Click para ações...")

    try:
        driver.switch_to.default_content()
    except Exception:
        pass

    seletores = [
        (By.ID, "acoes_ver"),
        (By.CSS_SELECTOR, "#acoes_ver"),
        (By.XPATH, "//a[@id='acoes_ver']"),
        (By.XPATH, "//a[contains(@title, 'Mostrar ações')]"),
        (By.XPATH, "//a[contains(normalize-space(.), 'Click para ações')]"),
        (By.XPATH, "//a[contains(normalize-space(.), 'click para ações')]"),
    ]

    def tentar_clicar_no_contexto(nome_contexto):
        ultimo_erro_local = None

        for by, seletor in seletores:
            try:
                elemento = WebDriverWait(driver, 6).until(
                    EC.presence_of_element_located((by, seletor))
                )

                driver.execute_script(
                    "arguments[0].scrollIntoView({block: 'center', inline: 'center'});",
                    elemento
                )
                time.sleep(0.5)

                try:
                    WebDriverWait(driver, 6).until(
                        EC.element_to_be_clickable((by, seletor))
                    )
                    elemento.click()
                except Exception:
                    driver.execute_script("arguments[0].click();", elemento)

                log(f"Click para ações clicado em {nome_contexto} usando seletor: {seletor}")
                time.sleep(1)
                return True

            except Exception as erro:
                ultimo_erro_local = erro
                continue

        return False

    # 1) Tenta no conteúdo principal
    if tentar_clicar_no_contexto("conteúdo principal"):
        return

    # 2) Tenta JS no conteúdo principal
    try:
        clicou = driver.execute_script("""
            const el = document.getElementById('acoes_ver');
            if (el) {
                el.scrollIntoView({block: 'center', inline: 'center'});
                el.click();
                return true;
            }
            return false;
        """)

        if clicou:
            log("Click para ações clicado via JavaScript no conteúdo principal.")
            time.sleep(1)
            return
    except Exception:
        pass

    # 3) Procura dentro de iframes
    try:
        driver.switch_to.default_content()
        iframes = driver.find_elements(By.TAG_NAME, "iframe")
        log(f"Iframes encontrados na tela: {len(iframes)}")

        for indice, iframe in enumerate(iframes):
            try:
                driver.switch_to.default_content()
                driver.switch_to.frame(iframe)

                if tentar_clicar_no_contexto(f"iframe {indice}"):
                    # NÃO volta para default_content aqui.
                    # O menu Exportações provavelmente fica no mesmo iframe.
                    return

                clicou = driver.execute_script("""
                    const el = document.getElementById('acoes_ver');
                    if (el) {
                        el.scrollIntoView({block: 'center', inline: 'center'});
                        el.click();
                        return true;
                    }
                    return false;
                """)

                if clicou:
                    log(f"Click para ações clicado via JavaScript no iframe {indice}.")
                    time.sleep(1)
                    return

            except Exception:
                continue

        driver.switch_to.default_content()

    except Exception as erro:
        try:
            driver.switch_to.default_content()
        except Exception:
            pass
        log(f"Falha ao procurar dentro de iframes: {erro}")

    screenshot = PASTA_ERROS / f"erro_click_acoes_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.png"
    html = PASTA_ERROS / f"erro_click_acoes_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.html"

    try:
        driver.save_screenshot(str(screenshot))
        log(f"Print do erro salvo em: {screenshot}")
    except Exception:
        pass

    try:
        html.write_text(driver.page_source, encoding="utf-8")
        log(f"HTML do erro salvo em: {html}")
    except Exception:
        pass

    raise Exception(
        "Não consegui clicar em 'Click para ações'. "
        "Envie o print/HTML gerado na pasta de downloads para análise."
    )



def clicar_click_para_acoes_rapido(driver):
    """
    Versão rápida do clique em "Click para ações".

    Primeiro tenta clicar diretamente pelo ID fixo acoes_ver usando JavaScript.
    Se não encontrar no conteúdo principal, procura dentro dos iframes.
    Caso falhe, usa a função completa clicar_click_para_acoes como fallback.

    Importante:
    Quando encontra dentro de iframe, mantém o Selenium dentro desse iframe,
    pois o menu Exportações normalmente abre no mesmo contexto.
    """
    log("Abrindo Click para ações em modo rápido...")

    script_click = """
        const el = document.getElementById('acoes_ver');
        if (el) {
            el.scrollIntoView({block: 'center', inline: 'center'});
            el.click();
            return true;
        }
        return false;
    """

    try:
        driver.switch_to.default_content()
    except Exception:
        pass

    # 1) Tenta direto no conteúdo principal
    try:
        clicou = driver.execute_script(script_click)

        if clicou:
            log("Click para ações clicado rapidamente no conteúdo principal.")
            time.sleep(0.4)
            return

    except Exception as erro:
        log(f"Clique rápido no conteúdo principal falhou: {erro}")

    # 2) Tenta direto dentro dos iframes
    try:
        driver.switch_to.default_content()
        iframes = driver.find_elements(By.TAG_NAME, "iframe")
        log(f"Procurando Click para ações em iframes no modo rápido. Total: {len(iframes)}")

        for indice, iframe in enumerate(iframes):
            try:
                driver.switch_to.default_content()
                driver.switch_to.frame(iframe)

                clicou = driver.execute_script(script_click)

                if clicou:
                    log(f"Click para ações clicado rapidamente no iframe {indice}.")
                    time.sleep(0.4)
                    return

            except Exception:
                continue

    except Exception as erro:
        log(f"Procura rápida em iframes falhou: {erro}")

    # 3) Fallback seguro: usa a função completa antiga, mais lenta, mas mais robusta
    log("Modo rápido não encontrou o botão. Usando fallback completo.")
    clicar_click_para_acoes(driver)


def clicar_exportacoes(driver):
    """
    Clica no botão Exportações usando o HTML informado:

    <div class="btn exportacoes btnFiltros"
         tipo_btn="acoes"
         url="emprestimoInterno.do"
         descricao="Exportações"
         permissao="exportarArquivo"
         seletor="id23"
         style="display: block;">
         Exportações
    </div>

    Essa função mantém o Selenium no iframe atual, pois o botão aparece
    depois de clicar em "Click para ações".
    """
    log("Abrindo Exportações...")

    seletores = [
        (By.CSS_SELECTOR, "div.exportacoes.btnFiltros[descricao='Exportações']"),
        (By.CSS_SELECTOR, "div.exportacoes[permissao='exportarArquivo']"),
        (By.CSS_SELECTOR, "div[descricao='Exportações'][permissao='exportarArquivo']"),
        (By.CSS_SELECTOR, "div[seletor='id23']"),
        (By.XPATH, "//div[@descricao='Exportações' and @permissao='exportarArquivo']"),
        (By.XPATH, "//div[contains(@class, 'exportacoes') and contains(@class, 'btnFiltros')]"),
        (By.XPATH, "//div[normalize-space(.)='Exportações' and @tipo_btn='acoes']"),
        (By.XPATH, "//div[normalize-space(.)='Exportações']"),
    ]

    ultimo_erro = None

    # Primeiro tenta no contexto atual, que deve ser o iframe onde foi clicado acoes_ver.
    for by, seletor in seletores:
        try:
            elemento = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((by, seletor))
            )

            driver.execute_script(
                "arguments[0].scrollIntoView({block: 'center', inline: 'center'});",
                elemento
            )
            time.sleep(0.5)

            try:
                WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((by, seletor))
                )
                elemento.click()
            except Exception:
                driver.execute_script("arguments[0].click();", elemento)

            log(f"Exportações clicado usando seletor: {seletor}")
            time.sleep(0.8)
            return

        except Exception as erro:
            ultimo_erro = erro
            continue

    # Se falhar, procura dentro de todos os iframes novamente
    try:
        driver.switch_to.default_content()
        iframes = driver.find_elements(By.TAG_NAME, "iframe")
        log(f"Procurando Exportações em iframes. Total: {len(iframes)}")

        for indice, iframe in enumerate(iframes):
            try:
                driver.switch_to.default_content()
                driver.switch_to.frame(iframe)

                for by, seletor in seletores:
                    try:
                        elemento = WebDriverWait(driver, 6).until(
                            EC.presence_of_element_located((by, seletor))
                        )

                        driver.execute_script(
                            "arguments[0].scrollIntoView({block: 'center', inline: 'center'});",
                            elemento
                        )
                        time.sleep(0.5)

                        try:
                            elemento.click()
                        except Exception:
                            driver.execute_script("arguments[0].click();", elemento)

                        log(f"Exportações clicado no iframe {indice} usando seletor: {seletor}")
                        time.sleep(0.8)
                        return

                    except Exception as erro:
                        ultimo_erro = erro
                        continue

            except Exception as erro:
                ultimo_erro = erro
                continue

    except Exception as erro:
        ultimo_erro = erro

    screenshot = PASTA_ERROS / f"erro_exportacoes_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.png"
    html = PASTA_ERROS / f"erro_exportacoes_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.html"

    try:
        driver.save_screenshot(str(screenshot))
        log(f"Print do erro salvo em: {screenshot}")
    except Exception:
        pass

    try:
        html.write_text(driver.page_source, encoding="utf-8")
        log(f"HTML do erro salvo em: {html}")
    except Exception:
        pass

    raise Exception(f"Não consegui clicar em Exportações. Último erro: {ultimo_erro}")


def clicar_exportar_layout_arquivo(driver):
    """
    Clica exatamente nesta opção:

    <a class="label ui-button ui-widget ui-corner-all acaoConsulta exportacao"
       tipo="acaoautocomplete"
       autocomplete="LayoutArquivo"
       campo="layoutArquivo"
       acao="emprestimoInterno.do?action=exportarArquivo">
       Exportar Layout Arquivo
    </a>

    Importante:
    Evita clicar em opções genéricas de exportação.
    """
    log("Abrindo Exportar Layout Arquivo...")

    seletores = [
        (By.CSS_SELECTOR, "a.acaoConsulta.exportacao[campo='layoutArquivo'][autocomplete='LayoutArquivo']"),
        (By.CSS_SELECTOR, "a[campo='layoutArquivo'][acao*='exportarArquivo']"),
        (By.CSS_SELECTOR, "a[autocomplete='LayoutArquivo']"),
        (By.XPATH, "//a[@campo='layoutArquivo' and @autocomplete='LayoutArquivo']"),
        (By.XPATH, "//a[contains(@class, 'acaoConsulta') and contains(@class, 'exportacao') and @campo='layoutArquivo']"),
        (By.XPATH, "//a[@acao='emprestimoInterno.do?action=exportarArquivo']"),
        (By.XPATH, "//a[normalize-space(.)='Exportar Layout Arquivo' and @campo='layoutArquivo']"),
    ]

    ultimo_erro = None

    def tentar_contexto(nome_contexto):
        nonlocal ultimo_erro

        for by, seletor in seletores:
            try:
                elemento = WebDriverWait(driver, 12).until(
                    EC.presence_of_element_located((by, seletor))
                )

                driver.execute_script(
                    "arguments[0].scrollIntoView({block: 'center', inline: 'center'});",
                    elemento
                )
                time.sleep(0.5)

                try:
                    WebDriverWait(driver, 12).until(
                        EC.element_to_be_clickable((by, seletor))
                    )
                    elemento.click()
                except Exception:
                    driver.execute_script("arguments[0].click();", elemento)

                log(f"Exportar Layout Arquivo clicado em {nome_contexto} usando seletor: {seletor}")
                time.sleep(0.8)
                return True

            except Exception as erro:
                ultimo_erro = erro
                continue

        return False

    # Primeiro tenta no contexto atual, normalmente o iframe correto
    if tentar_contexto("contexto atual"):
        return

    # Depois tenta nos iframes
    try:
        driver.switch_to.default_content()
        iframes = driver.find_elements(By.TAG_NAME, "iframe")
        log(f"Procurando Exportar Layout Arquivo em iframes. Total: {len(iframes)}")

        for indice, iframe in enumerate(iframes):
            try:
                driver.switch_to.default_content()
                driver.switch_to.frame(iframe)

                if tentar_contexto(f"iframe {indice}"):
                    return

            except Exception as erro:
                ultimo_erro = erro
                continue

    except Exception as erro:
        ultimo_erro = erro

    screenshot = PASTA_ERROS / f"erro_layout_arquivo_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.png"
    html = PASTA_ERROS / f"erro_layout_arquivo_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.html"

    try:
        driver.save_screenshot(str(screenshot))
        log(f"Print do erro salvo em: {screenshot}")
    except Exception:
        pass

    try:
        html.write_text(driver.page_source, encoding="utf-8")
        log(f"HTML do erro salvo em: {html}")
    except Exception:
        pass

    raise Exception(
        "Não consegui clicar no botão correto 'Exportar Layout Arquivo'. "
        "Parei aqui para evitar preencher campo errado."
    )

def abrir_tela_exportacao(driver):
    clicar_click_para_acoes_rapido(driver)
    clicar_exportacoes(driver)
    clicar_exportar_layout_arquivo(driver)


def selecionar_layout(driver, nome_exportacao):
    """
    Preenche e SELECIONA o layout correto no autocomplete.

    Campo correto:
    <input type="text" name="layoutArquivo" class="txt100 ac_field"
           id="_id_layoutArquivo" autocomplete="off">

    Correção:
    - Não usa campo.click(), pois o painel paneAv pode interceptar o clique.
    - Foca o campo via JavaScript.
    - Digita o texto com send_keys para disparar o autocomplete do site.
    - Seleciona a opção correta dentro do painel de autocomplete.
    """
    log(f"Preenchendo layout: {nome_exportacao}")

    seletores_campo = [
        (By.ID, "_id_layoutArquivo"),
        (By.CSS_SELECTOR, "input#_id_layoutArquivo[name='layoutArquivo']"),
        (By.CSS_SELECTOR, "input[name='layoutArquivo'].ac_field"),
        (By.CSS_SELECTOR, "input[name='layoutArquivo']"),
        (By.XPATH, "//input[@id='_id_layoutArquivo']"),
        (By.XPATH, "//input[@name='layoutArquivo' and contains(@class, 'ac_field')]"),
        (By.XPATH, "//input[@name='layoutArquivo']"),
    ]

    campo = None
    ultimo_erro = None

    def procurar_campo_no_contexto(nome_contexto):
        nonlocal ultimo_erro

        for by, seletor in seletores_campo:
            try:
                elemento = WebDriverWait(driver, 15).until(
                    EC.presence_of_element_located((by, seletor))
                )

                driver.execute_script(
                    "arguments[0].scrollIntoView({block: 'center', inline: 'center'});",
                    elemento
                )
                time.sleep(0.5)

                log(f"Campo layoutArquivo encontrado em {nome_contexto} usando seletor: {seletor}")
                return elemento

            except Exception as erro:
                ultimo_erro = erro
                continue

        return None

    # 1) Procura no contexto atual
    campo = procurar_campo_no_contexto("contexto atual")

    # 2) Procura dentro dos iframes, se necessário
    if campo is None:
        try:
            driver.switch_to.default_content()
            iframes = driver.find_elements(By.TAG_NAME, "iframe")
            log(f"Procurando campo layoutArquivo em iframes. Total: {len(iframes)}")

            for indice, iframe in enumerate(iframes):
                try:
                    driver.switch_to.default_content()
                    driver.switch_to.frame(iframe)

                    campo = procurar_campo_no_contexto(f"iframe {indice}")

                    if campo is not None:
                        break

                except Exception as erro:
                    ultimo_erro = erro
                    continue

        except Exception as erro:
            ultimo_erro = erro

    if campo is None:
        screenshot = PASTA_ERROS / f"erro_campo_layout_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.png"
        html = PASTA_ERROS / f"erro_campo_layout_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.html"

        try:
            driver.save_screenshot(str(screenshot))
            log(f"Print do erro salvo em: {screenshot}")
        except Exception:
            pass

        try:
            html.write_text(driver.page_source, encoding="utf-8")
            log(f"HTML do erro salvo em: {html}")
        except Exception:
            pass

        raise Exception(f"Não encontrei o campo layoutArquivo. Último erro: {ultimo_erro}")

    # 3) Limpa painéis antigos, mas NÃO remove novos painéis depois que digitar
    try:
        driver.switch_to.active_element.send_keys(Keys.ESCAPE)
        time.sleep(0.3)
    except Exception:
        pass

    # 4) Foca via JS, limpa e digita com send_keys para disparar autocomplete real do site
    try:
        driver.execute_script("""
            const campo = arguments[0];
            campo.removeAttribute('readonly');
            campo.removeAttribute('disabled');
            campo.focus();
            campo.value = '';
            campo.dispatchEvent(new Event('input', { bubbles: true }));
            campo.dispatchEvent(new Event('change', { bubbles: true }));
        """, campo)

        time.sleep(0.3)

        campo.send_keys(nome_exportacao)
        time.sleep(1.5)

    except Exception as erro:
        log(f"Falha ao digitar no campo com send_keys; tentando JS puro. Detalhe: {erro}")

        driver.execute_script("""
            const campo = arguments[0];
            const valor = arguments[1];
            campo.removeAttribute('readonly');
            campo.removeAttribute('disabled');
            campo.focus();
            campo.value = valor;
            campo.dispatchEvent(new Event('input', { bubbles: true }));
            campo.dispatchEvent(new Event('keyup', { bubbles: true }));
            campo.dispatchEvent(new Event('change', { bubbles: true }));
        """, campo, nome_exportacao)

        time.sleep(1.5)

    # 5) Seleciona a opção correta do autocomplete.
    #    Isso é importante: só preencher texto pode não configurar o valor interno do site.
    def normalizar(txt):
        return (txt or "").strip().lower()

    texto_alvo = normalizar(nome_exportacao)
    selecionou = False
    ultimo_erro_opcao = None

    seletores_opcoes = [
        ".paneAv *",
        ".ui-autocomplete *",
        ".automaticotrue *",
        "li",
        "a",
        "div",
        "span",
    ]

    for seletor in seletores_opcoes:
        try:
            opcoes = driver.find_elements(By.CSS_SELECTOR, seletor)

            for opcao in opcoes:
                try:
                    txt = normalizar(opcao.text)

                    if not txt:
                        continue

                    if texto_alvo in txt:
                        driver.execute_script(
                            "arguments[0].scrollIntoView({block: 'center', inline: 'center'});",
                            opcao
                        )
                        time.sleep(0.2)

                        try:
                            opcao.click()
                        except Exception:
                            driver.execute_script("arguments[0].click();", opcao)

                        selecionou = True
                        log(f"Opção do autocomplete selecionada: {opcao.text}")
                        time.sleep(1)
                        break

                except Exception as erro:
                    ultimo_erro_opcao = erro
                    continue

            if selecionou:
                break

        except Exception as erro:
            ultimo_erro_opcao = erro
            continue

    # 6) Se não conseguiu clicar na opção, usa teclado como fallback
    if not selecionou:
        log(f"Não encontrei opção visível do autocomplete. Tentando confirmar com teclado. Detalhe: {ultimo_erro_opcao}")

        try:
            campo.send_keys(Keys.ARROW_DOWN)
            time.sleep(0.5)
            campo.send_keys(Keys.ENTER)
            time.sleep(1)
            selecionou = True
        except Exception as erro:
            log(f"Falha ao confirmar autocomplete com teclado: {erro}")

    # 7) Validação final
    try:
        valor_atual = campo.get_attribute("value")
        log(f"Valor atual no campo layoutArquivo: {valor_atual}")

        if not valor_atual:
            raise Exception("O campo layoutArquivo ficou vazio após o preenchimento.")

    except Exception as erro:
        screenshot = PASTA_ERROS / f"erro_valor_layout_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.png"
        html = PASTA_ERROS / f"erro_valor_layout_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.html"

        try:
            driver.save_screenshot(str(screenshot))
            log(f"Print do erro salvo em: {screenshot}")
        except Exception:
            pass

        try:
            html.write_text(driver.page_source, encoding="utf-8")
            log(f"HTML do erro salvo em: {html}")
        except Exception:
            pass

        raise Exception(f"Falha ao validar valor do campo layoutArquivo: {erro}")

    log(f"Layout preenchido e selecionado: {nome_exportacao}")


def clicar_primeiro_download(driver):
    """
    Clica no arquivo mais recente da lista de exportações.
    Procura links com idDigitalizacao, exibirDigitalizacao, layoutArquivo.do ou ícone file-alt.
    """
    log("Localizando arquivo mais recente para download...")

    padrao_data = re.compile(r"(\d{2}/\d{2}/\d{4})(?:\s+(\d{2}:\d{2}(?::\d{2})?))?")
    formatos_data = ["%d/%m/%Y %H:%M:%S", "%d/%m/%Y %H:%M", "%d/%m/%Y"]

    def converter_data(texto_linha):
        if not texto_linha:
            return None
        match = padrao_data.search(texto_linha)
        if not match:
            return None
        data = match.group(1)
        hora = match.group(2)
        data_texto = f"{data} {hora}" if hora else data
        for formato in formatos_data:
            try:
                return datetime.strptime(data_texto, formato)
            except Exception:
                continue
        return None

    def coletar_links():
        xpaths = [
            "//a[contains(@href, 'idDigitalizacao')]",
            "//a[contains(@href, 'exibirDigitalizacao')]",
            "//a[contains(@href, 'layoutArquivo.do')]",
            "//a[.//i[contains(@class, 'file-alt')]]",
            "//a[contains(@title, 'usuário') or contains(@title, 'usuario')]",
        ]

        encontrados = []
        for xp in xpaths:
            try:
                encontrados.extend(driver.find_elements(By.XPATH, xp))
            except Exception:
                pass

        unicos = []
        chaves = set()

        for link in encontrados:
            try:
                href = link.get_attribute("href") or ""
                html = link.get_attribute("outerHTML") or ""
                title = link.get_attribute("title") or ""
                chave = href + "|" + title + "|" + html[:150]

                if chave in chaves:
                    continue

                if (
                    "idDigitalizacao" in href
                    or "exibirDigitalizacao" in href
                    or "layoutArquivo.do" in href
                    or "file-alt" in html
                    or "fa-file" in html
                    or "usuário" in title.lower()
                    or "usuario" in title.lower()
                ):
                    chaves.add(chave)
                    unicos.append(link)
            except Exception:
                continue

        return unicos

    def tentar_no_contexto(nome_contexto):
        links = coletar_links()
        log(f"Links de arquivo encontrados em {nome_contexto}: {len(links)}")

        if not links:
            return False

        candidatos = []

        for indice, link in enumerate(links):
            try:
                linha = link.find_element(By.XPATH, "./ancestor::tr[1]")
                texto_linha = linha.text
            except Exception:
                texto_linha = link.text or ""

            candidatos.append({
                "indice": indice,
                "link": link,
                "data": converter_data(texto_linha),
                "texto": texto_linha,
                "href": link.get_attribute("href"),
                "title": link.get_attribute("title"),
            })

        candidatos_com_data = [c for c in candidatos if c["data"] is not None]

        if candidatos_com_data:
            candidatos_com_data.sort(key=lambda c: c["data"], reverse=True)
            escolhido = candidatos_com_data[0]
            log("Arquivo mais recente definido pela data da linha: " + escolhido["data"].strftime("%d/%m/%Y %H:%M:%S"))
        else:
            escolhido = candidatos[0]
            log("Não consegui identificar a data das linhas. Vou clicar no primeiro arquivo encontrado.")

        log(f"Href escolhido: {escolhido.get('href')}")
        log(f"Title escolhido: {escolhido.get('title')}")
        log(f"Texto da linha escolhida: {escolhido.get('texto')}")

        elemento = escolhido["link"]
        driver.execute_script("arguments[0].scrollIntoView({block: 'center', inline: 'center'});", elemento)
        time.sleep(0.5)

        try:
            elemento.click()
        except Exception:
            driver.execute_script("arguments[0].click();", elemento)

        log("Clique no arquivo mais recente realizado.")
        time.sleep(1)
        return True

    def tentar_todos_contextos():
        try:
            if tentar_no_contexto("contexto atual"):
                return True
        except Exception as erro:
            log(f"Falha ao buscar download no contexto atual: {erro}")

        try:
            driver.switch_to.default_content()
            if tentar_no_contexto("conteúdo principal"):
                return True
        except Exception as erro:
            log(f"Falha ao buscar download no conteúdo principal: {erro}")

        try:
            driver.switch_to.default_content()
            iframes = driver.find_elements(By.TAG_NAME, "iframe")
            log(f"Procurando links de download em iframes. Total: {len(iframes)}")

            for indice, iframe in enumerate(iframes):
                try:
                    driver.switch_to.default_content()
                    driver.switch_to.frame(iframe)
                    if tentar_no_contexto(f"iframe {indice}"):
                        return True
                except Exception as erro:
                    log(f"Não encontrei download no iframe {indice}: {erro}")
                    continue
        except Exception as erro:
            log(f"Falha ao procurar download em iframes: {erro}")

        return False

    for tentativa in range(1, 6):
        log(f"Tentativa {tentativa} de localizar link de download...")
        if tentar_todos_contextos():
            return
        time.sleep(3)

    salvar_diagnostico_erro(driver, "erro_download_arquivo")
    raise Exception(
        "Não encontrei o link correto de download. "
        "Verifique o HTML salvo na pasta erros para confirmar se a página gerou a lista de arquivos."
    )



def clicar_executar_exportacao(driver):
    """
    Clica no botão Executar da tela de exportação.
    Tenta seletores específicos antes de usar texto genérico.
    """
    log("Clicando em Executar.")

    seletores = [
        (By.XPATH, "//button[normalize-space(.)='Executar']"),
        (By.XPATH, "//input[@type='button' and translate(@value, 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz')='executar']"),
        (By.XPATH, "//input[@type='submit' and translate(@value, 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz')='executar']"),
        (By.XPATH, "//a[normalize-space(.)='Executar']"),
        (By.XPATH, "//*[normalize-space(.)='Executar' and (self::button or self::input or self::a or contains(@class,'btn'))]"),
        (By.XPATH, "//*[contains(translate(normalize-space(.), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'executar')]"),
    ]

    ultimo_erro = None

    def tentar_contexto(nome_contexto):
        nonlocal ultimo_erro

        for by, seletor in seletores:
            try:
                elemento = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((by, seletor))
                )

                driver.execute_script(
                    "arguments[0].scrollIntoView({block: 'center', inline: 'center'});",
                    elemento
                )
                time.sleep(0.3)

                try:
                    elemento.click()
                except Exception:
                    driver.execute_script("arguments[0].click();", elemento)

                log(f"Executar clicado em {nome_contexto} usando seletor: {seletor}")
                time.sleep(2)
                return True

            except Exception as erro:
                ultimo_erro = erro
                continue

        return False

    if tentar_contexto("contexto atual"):
        return

    try:
        driver.switch_to.default_content()
        iframes = driver.find_elements(By.TAG_NAME, "iframe")

        for indice, iframe in enumerate(iframes):
            try:
                driver.switch_to.default_content()
                driver.switch_to.frame(iframe)

                if tentar_contexto(f"iframe {indice}"):
                    return

            except Exception as erro:
                ultimo_erro = erro
                continue

    except Exception as erro:
        ultimo_erro = erro

    screenshot = PASTA_ERROS / f"erro_executar_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.png"
    html = PASTA_ERROS / f"erro_executar_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.html"

    try:
        driver.save_screenshot(str(screenshot))
        log(f"Print do erro salvo em: {screenshot}")
    except Exception:
        pass

    try:
        html.write_text(driver.page_source, encoding="utf-8")
        log(f"HTML do erro salvo em: {html}")
    except Exception:
        pass

    raise Exception(f"Não consegui clicar no botão Executar. Último erro: {ultimo_erro}")


def baixar_exportacao(driver, nome_exportacao, limpar_antes=True):
    if limpar_antes:
        limpar_downloads()
    else:
        log("Downloads existentes preservados antes da exportação.")

    abrir_tela_exportacao(driver)
    selecionar_layout(driver, nome_exportacao)

    clicar_executar_exportacao(driver)
    time.sleep(3)

    log("Aguardando lista de arquivos...")
    esperar(driver, 120).until(
        EC.presence_of_element_located(
            (
                By.XPATH,
                "//*[contains(text(), 'Data') or contains(text(), 'Código') or contains(text(), 'Codigo')]"
                " | //a[contains(@href, 'layoutArquivo.do')]"
                " | //a[contains(@href, 'exibirDigitalizacao')]"
                " | //a[contains(@href, 'idDigitalizacao')]"
                " | //a[.//i[contains(@class, 'file-alt')]]"
            )
        )
    )

    inicio_download = time.time() - 5

    log("Clicando no download mais recente...")
    clicar_primeiro_download(driver)

    log("Aguardando download finalizar...")
    arquivo_baixado = esperar_download_novo(PASTA_DOWNLOAD, inicio=inicio_download)

    log(f"Arquivo baixado para {nome_exportacao}: {arquivo_baixado}")

    return arquivo_baixado


# ============================================================
# ATUALIZAÇÃO DO EXCEL
# ============================================================

def remover_linhas_mes_vigente(ws, coluna_data="L"):
    hoje = datetime.today()
    data_ini = datetime(hoje.year, hoje.month, 1)

    if hoje.month == 12:
        data_fim = datetime(hoje.year + 1, 1, 1)
    else:
        data_fim = datetime(hoje.year, hoje.month + 1, 1)

    ultima_linha = ws.max_row
    removidas = 0

    for linha in range(ultima_linha, 1, -1):
        valor = ws[f"{coluna_data}{linha}"].value

        if valor is None:
            continue

        if isinstance(valor, datetime):
            data = valor
        else:
            try:
                data = pd.to_datetime(valor, dayfirst=True, errors="coerce")
                if pd.isna(data):
                    continue
                data = data.to_pydatetime()
            except Exception:
                continue

        if data_ini <= data < data_fim:
            ws.delete_rows(linha)
            removidas += 1

    return removidas





def limpar_linhas_mes_vigente_preservando_formulas(ws, coluna_data="L", colunas_formula=None, ultima_coluna_dados="T"):
    """
    Remove linhas do mês vigente baseado na coluna Data CPC.

    Mesmo removendo a linha inteira, as fórmulas das colunas U e V são preservadas
    porque primeiro capturamos os modelos das fórmulas e depois reaplicamos nas linhas novas.
    """
    return remover_linhas_mes_vigente(ws, coluna_data)





def coluna_para_numero(coluna):
    coluna = coluna.upper()
    numero = 0

    for caractere in coluna:
        numero = numero * 26 + (ord(caractere) - ord("A") + 1)

    return numero


def obter_modelos_formulas(ws, colunas_formula):
    """
    Captura as fórmulas modelo das colunas U e V antes de mexer nos dados.
    """
    modelos = {}

    for coluna in colunas_formula:
        formula_encontrada = None
        linha_formula = None

        for linha in range(2, ws.max_row + 1):
            valor = ws[f"{coluna}{linha}"].value

            if isinstance(valor, str) and valor.startswith("="):
                formula_encontrada = valor
                linha_formula = linha
                break

        if formula_encontrada:
            modelos[coluna] = {
                "formula": formula_encontrada,
                "linha": linha_formula,
                "celula": f"{coluna}{linha_formula}",
            }
            log(f"Fórmula modelo encontrada em {coluna}{linha_formula}: {formula_encontrada}")
        else:
            modelos[coluna] = None
            log(f"Atenção: nenhuma fórmula modelo encontrada na coluna {coluna}.")

    return modelos


def aplicar_formulas(ws, modelos_formula, linha_inicial, linha_final):
    """
    Replica fórmulas para as linhas informadas.
    """
    if linha_final < linha_inicial:
        log("Nenhuma linha nova para aplicar fórmulas.")
        return

    for coluna, modelo in modelos_formula.items():
        if not modelo:
            log(f"Sem fórmula modelo para a coluna {coluna}; pulando.")
            continue

        formula_base = modelo["formula"]
        celula_origem = modelo["celula"]

        for linha in range(linha_inicial, linha_final + 1):
            destino = f"{coluna}{linha}"

            try:
                ws[destino] = Translator(
                    formula_base,
                    origin=celula_origem
                ).translate_formula(destino)

            except Exception:
                ws[destino] = formula_base

        log(f"Fórmulas aplicadas na coluna {coluna}, da linha {linha_inicial} até {linha_final}.")


def obter_ultima_linha_com_dados(ws, ultima_coluna_total):
    """
    Retorna a última linha realmente preenchida, ignorando linhas que ficaram
    vazias depois da reconstrução da base.
    """
    ultima_coluna_total_num = coluna_para_numero(ultima_coluna_total)

    for linha in range(ws.max_row, 1, -1):
        for coluna in range(1, ultima_coluna_total_num + 1):
            valor = ws.cell(row=linha, column=coluna).value
            if valor is not None and str(valor).strip() != "":
                return linha

    return 1


def ajustar_tabelas_excel(ws, ultima_linha=None):
    """
    Expande Tabelas do Excel para incluir as linhas atuais.
    """
    if not ws.tables:
        return

    if ultima_linha is None:
        ultima_linha = ws.max_row

    ultima_linha = max(int(ultima_linha), 1)

    for tabela in ws.tables.values():
        min_col, min_row, max_col, _ = range_boundaries(tabela.ref)
        linha_final = max(ultima_linha, min_row)
        nova_ref = (
            f"{get_column_letter(min_col)}{min_row}:"
            f"{get_column_letter(max_col)}{linha_final}"
        )
        tabela.ref = nova_ref
        log(f"Tabela Excel ajustada: {tabela.name} -> {nova_ref}")


def valor_para_data(valor):
    """
    Converte valores variados para data.
    """
    if valor is None:
        return None

    if isinstance(valor, datetime):
        return valor

    try:
        data = pd.to_datetime(valor, dayfirst=True, errors="coerce")
        if pd.isna(data):
            return None
        return data.to_pydatetime()
    except Exception:
        return None


def linha_eh_mes_vigente(valor_data):
    hoje = datetime.today()
    data = valor_para_data(valor_data)

    if data is None:
        return False

    return data.month == hoje.month and data.year == hoje.year


def limpar_e_reconstruir_base(
    ws,
    modelos_formula,
    coluna_data="L",
    ultima_coluna_total="V"
):
    """
    Limpeza otimizada do mês vigente.

    A última coluna total muda conforme a base:
    - Vendas: até V
    - Receita: até Y
    """
    log("Iniciando limpeza otimizada do mês vigente...")

    col_data_num = coluna_para_numero(coluna_data)
    ultima_coluna_total_num = coluna_para_numero(ultima_coluna_total)

    linhas_mantidas = []
    removidas = 0

    for linha_idx in range(2, ws.max_row + 1):
        valor_data = ws.cell(row=linha_idx, column=col_data_num).value

        valores_linha = [
            ws.cell(row=linha_idx, column=col).value
            for col in range(1, ultima_coluna_total_num + 1)
        ]

        linha_vazia = all(valor is None or str(valor).strip() == "" for valor in valores_linha)

        if linha_vazia:
            continue

        if linha_eh_mes_vigente(valor_data):
            removidas += 1
        else:
            linhas_mantidas.append(valores_linha)

    log(f"Linhas do mês vigente encontradas para remoção: {removidas}")
    log(f"Linhas antigas mantidas: {len(linhas_mantidas)}")

    # Limpa conteúdo da linha 2 para baixo, até a última coluna total da base.
    for row in ws.iter_rows(
        min_row=2,
        max_row=ws.max_row,
        min_col=1,
        max_col=ultima_coluna_total_num
    ):
        for cell in row:
            cell.value = None

    # Reescreve as linhas mantidas.
    linha_destino = 2

    for valores in linhas_mantidas:
        for coluna, valor in enumerate(valores, start=1):
            ws.cell(row=linha_destino, column=coluna, value=valor)

        linha_destino += 1

    ultima_linha_mantida = linha_destino - 1

    # Reaplica fórmulas nas linhas mantidas.
    if ultima_linha_mantida >= 2:
        aplicar_formulas(ws, modelos_formula, 2, ultima_linha_mantida)

    return removidas, linha_destino

def ler_linhas_exportacao(arquivo_origem):
    """
    Lê o arquivo exportado pelo site.

    Aceita:
    - .csv
    - .xlsx
    - .xlsm

    Retorna lista de linhas sem o cabeçalho.
    """
    arquivo_origem = Path(arquivo_origem)
    extensao = arquivo_origem.suffix.lower()

    if extensao == ".csv":
        log("Arquivo de origem é CSV. Lendo com pandas...")

        tentativas = [
            {"sep": None, "engine": "python", "encoding": "utf-8-sig"},
            {"sep": ";", "encoding": "utf-8-sig"},
            {"sep": ",", "encoding": "utf-8-sig"},
            {"sep": None, "engine": "python", "encoding": "latin1"},
            {"sep": ";", "encoding": "latin1"},
            {"sep": ",", "encoding": "latin1"},
            {"sep": None, "engine": "python", "encoding": "cp1252"},
            {"sep": ";", "encoding": "cp1252"},
            {"sep": ",", "encoding": "cp1252"},
        ]

        ultimo_erro = None

        for params in tentativas:
            try:
                df = pd.read_csv(arquivo_origem, dtype=object, **params)
                df = df.where(pd.notnull(df), None)

                log(f"CSV lido com sucesso. Linhas: {len(df)} | Colunas: {len(df.columns)}")

                return df.values.tolist()

            except Exception as erro:
                ultimo_erro = erro
                continue

        raise Exception(f"Não consegui ler o CSV exportado. Último erro: {ultimo_erro}")

    if extensao in [".xlsx", ".xlsm"]:
        log("Arquivo de origem é Excel. Lendo com openpyxl...")

        wb_origem = load_workbook(arquivo_origem, data_only=True, read_only=True, keep_links=False)
        ws_origem = wb_origem.active

        linhas = []

        for linha in ws_origem.iter_rows(min_row=2, values_only=True):
            if all(valor is None for valor in linha):
                continue

            linhas.append(list(linha))

        wb_origem.close()

        log(f"Excel lido com sucesso. Linhas: {len(linhas)}")

        return linhas

    raise ValueError(
        f"Formato de origem não suportado: {extensao}. "
        "Use .csv, .xlsx ou .xlsm."
    )



def esperar_arquivo_excel_liberar(arquivo, timeout=None):
    """
    Aguarda o arquivo Excel ficar livre para edição.

    O erro [Errno 13] Permission denied geralmente acontece quando:
    - o arquivo está aberto no Excel;
    - o arquivo está bloqueado pelo OneDrive/SharePoint;
    - outro processo está sincronizando/salvando a planilha.

    Esta função tenta abrir o arquivo em modo leitura/escrita.
    Se estiver bloqueado, aguarda e tenta novamente.
    """
    arquivo = Path(arquivo)
    timeout = TIMEOUT_EXCEL if timeout is None else timeout
    fim = time.time() + timeout

    while time.time() < fim:
        try:
            with open(arquivo, "r+b"):
                log(f"Arquivo liberado para edição: {arquivo}")
                return True

        except PermissionError:
            log(
                "Arquivo está bloqueado. Feche a planilha no Excel "
                "e aguarde a sincronização do OneDrive/SharePoint..."
            )
            time.sleep(5)

        except FileNotFoundError:
            raise

        except Exception as erro:
            log(f"Aguardando liberação do arquivo. Detalhe: {erro}")
            time.sleep(5)

    raise PermissionError(
        f"Arquivo continua bloqueado após {timeout} segundos: {arquivo}. "
        "Feche o Excel, confira se ninguém está com a planilha aberta "
        "e aguarde o OneDrive/SharePoint terminar a sincronização."
    )


def copiar_dados_excel_origem_para_destino(
    arquivo_origem,
    arquivo_destino,
    coluna_data="L",
    colunas_formula=None,
    ultima_coluna_dados="T",
    ultima_coluna_total="V"
):
    """
    Processo Excel otimizado:

    1. Lê o CSV baixado a partir da linha 2.
    2. Abre o arquivo destino.
    3. Captura fórmulas conforme a base:
       - Vendas: U e V
       - Receita: W, X e Y
    4. Remove mês vigente sem usar delete_rows linha por linha.
    5. Cola dados novos até a última coluna de dados.
    6. Reaplica fórmulas até a última linha.
    """
    if colunas_formula is None:
        colunas_formula = ["U", "V"]

    arquivo_origem = Path(arquivo_origem)
    arquivo_destino = Path(arquivo_destino)

    if not arquivo_origem.exists():
        raise FileNotFoundError(f"Arquivo baixado não encontrado: {arquivo_origem}")

    if not arquivo_destino.exists():
        raise FileNotFoundError(f"Arquivo de destino não encontrado: {arquivo_destino}")

    log(f"Abrindo origem: {arquivo_origem}")
    log(f"Abrindo destino: {arquivo_destino}")
    log(f"Colunas de fórmula: {colunas_formula}")
    log(f"Colar dados até a coluna: {ultima_coluna_dados}")
    log(f"Limpar/reescrever até a coluna: {ultima_coluna_total}")

    esperar_arquivo_excel_liberar(arquivo_destino)

    linhas_origem = ler_linhas_exportacao(arquivo_origem)
    if not linhas_origem:
        raise ValueError(
            f"O arquivo exportado não possui linhas de dados: {arquivo_origem}. "
            "A atualização foi interrompida para evitar apagar o mês vigente sem repor os dados."
        )

    wb_destino = None

    try:
        with acompanhar_operacao("Abrindo a planilha de destino", intervalo=10):
            wb_destino = load_workbook(arquivo_destino, keep_links=False)
        ws_destino = wb_destino.active

        modelos_formula = obter_modelos_formulas(ws_destino, colunas_formula)

        with acompanhar_operacao("Analisando e reconstruindo a base", intervalo=10):
            removidas, primeira_linha_colagem = limpar_e_reconstruir_base(
                ws_destino,
                modelos_formula,
                coluna_data=coluna_data,
                ultima_coluna_total=ultima_coluna_total
            )

        log(f"Linhas removidas do mês vigente: {removidas}")
        log(f"Primeira linha para colagem: {primeira_linha_colagem}")

        linha_destino = primeira_linha_colagem
        linhas_coladas = 0

        limite_coluna_dados = coluna_para_numero(ultima_coluna_dados)

        for linha in linhas_origem:
            if all(valor is None or str(valor).strip() == "" for valor in linha):
                continue

            # Cola somente até a última coluna de dados,
            # preservando as colunas de fórmula.
            for coluna, valor in enumerate(linha[:limite_coluna_dados], start=1):
                ws_destino.cell(row=linha_destino, column=coluna, value=valor)

            linha_destino += 1
            linhas_coladas += 1

        ultima_linha_colada = linha_destino - 1
        

        log(f"Linhas novas coladas: {linhas_coladas}")
        if linhas_coladas == 0:
            raise ValueError(
                "Nenhuma linha válida foi colada no arquivo destino. "
                "Verifique se o arquivo exportado veio vazio ou em formato inesperado."
            )

        aplicar_formulas(
            ws_destino,
            modelos_formula,
            primeira_linha_colagem,
            ultima_linha_colada
        )

        ultima_linha_real = obter_ultima_linha_com_dados(ws_destino, ultima_coluna_total)
        ajustar_tabelas_excel(ws_destino, ultima_linha=ultima_linha_real)

        esperar_arquivo_excel_liberar(arquivo_destino)

        with acompanhar_operacao("Salvando a planilha atualizada", intervalo=10):
            wb_destino.save(arquivo_destino)
        log(f"Arquivo salvo com sucesso: {arquivo_destino}")

    finally:
        if wb_destino:
            wb_destino.close()


# ============================================================
# PROCESSO PRINCIPAL
# ============================================================

def processar_exportacao_vendas(driver):
    log("========== INICIANDO EXPORTAÇÃO VENDAS ==========")

    arquivo_vendas = baixar_exportacao(driver, "Exportação vendas")

    copiar_dados_excel_origem_para_destino(
        arquivo_origem=arquivo_vendas,
        arquivo_destino=ARQUIVO_FAT_VENDAS,
        coluna_data=COLUNA_DATA_CPC,
        colunas_formula=COLUNAS_FORMULA_VENDAS,
        ultima_coluna_dados=ULTIMA_COLUNA_DADOS_VENDAS,
        ultima_coluna_total=ULTIMA_COLUNA_TOTAL_VENDAS
    )

    log("========== EXPORTAÇÃO VENDAS CONCLUÍDA ==========")
    return ARQUIVO_FAT_VENDAS


def processar_receita_gerada(driver):
    log("========== INICIANDO RECEITA GERADA ==========")

    arquivo_receita = baixar_exportacao(driver, "Receita gerada")

    copiar_dados_excel_origem_para_destino(
        arquivo_origem=arquivo_receita,
        arquivo_destino=ARQUIVO_FAT_RECEITA,
        coluna_data=COLUNA_DATA_CPC,
        colunas_formula=COLUNAS_FORMULA_RECEITA,
        ultima_coluna_dados=ULTIMA_COLUNA_DADOS_RECEITA,
        ultima_coluna_total=ULTIMA_COLUNA_TOTAL_RECEITA
    )

    log("========== RECEITA GERADA CONCLUÍDA ==========")
    return ARQUIVO_FAT_RECEITA


def parse_args(argv=None):
    parser = argparse.ArgumentParser(
        description="Automação Crediagora: exporta vendas/receita e atualiza as planilhas fat."
    )
    parser.add_argument(
        "--check",
        action="store_true",
        help="Valida configuração, credenciais e caminhos sem abrir Chrome nem alterar planilhas."
    )
    parser.add_argument(
        "--check-login",
        action="store_true",
        help="Testa somente o login e fecha o Chrome, sem baixar ou alterar planilhas."
    )
    parser.add_argument(
        "--liberar-bloqueio-login",
        action="store_true",
        help="Libera a proteção local somente após a conta ser desbloqueada no portal."
    )
    parser.add_argument(
        "--exportacao",
        choices=["todas", "vendas", "receita"],
        default="todas",
        help="Escolhe qual base atualizar. Padrão: todas."
    )
    parser.add_argument(
        "--headless",
        action="store_true",
        help="Executa o Chrome em modo invisível."
    )
    parser.add_argument(
        "--manter-downloads",
        action="store_true",
        help="Mantém os arquivos baixados ao final mesmo quando a execução terminar com sucesso."
    )
    parser.add_argument(
        "--manter-navegador",
        action="store_true",
        help="Mantém o Chrome aberto ao final para inspeção manual. Ignorado em modo headless."
    )
    parser.add_argument(
        "--sem-backup",
        action="store_true",
        help="Não cria backup antes de atualizar as planilhas."
    )
    grupo_ercard = parser.add_mutually_exclusive_group()
    grupo_ercard.add_argument(
        "--sem-ercard",
        action="store_true",
        help="Executa somente o fluxo Crediagora já existente."
    )
    grupo_ercard.add_argument(
        "--somente-ercard",
        action="store_true",
        help="Executa somente a nova fase ERCard."
    )
    return parser.parse_args(argv)


def main(argv=None):
    args = parse_args(argv)
    driver = None
    status_log = "erro"
    headless = MODO_HEADLESS or args.headless
    manter_downloads = MANTER_DOWNLOADS or args.manter_downloads
    manter_navegador = (MANTER_NAVEGADOR or args.manter_navegador) and not headless
    criar_backups = not args.sem_backup and not args.check_login
    executar_crediagora = not args.somente_ercard
    executar_ercard = not args.sem_ercard and not args.check_login
    config_ercard = criar_configuracao_ercard()
    LOGS_ATIVACAO.clear()

    try:
        criar_pastas()
        log(
            "======= Iniciando processo de exportação: "
            f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} ======="
        )
        if args.liberar_bloqueio_login:
            liberar_bloqueio_local_login()
            log("Proteção local de login liberada. Nenhuma autenticação foi realizada.")
            status_log = "check"
            return 0

        if executar_crediagora:
            validar_configuracao()
        if executar_ercard:
            config_ercard.validar()
        if executar_crediagora and not args.check_login:
            validar_caminhos(args.exportacao)
        mostrar_resumo_configuracao(
            exportacao=args.exportacao,
            headless=headless,
            manter_downloads=manter_downloads,
            manter_navegador=manter_navegador,
            criar_backups=criar_backups
        )

        if args.check:
            log("Diagnóstico concluído. Nenhum login, download ou alteração foi executado.")
            status_log = "check"
            return

        if args.check_login:
            log("Backups dispensados no diagnóstico exclusivo de login.")

        driver = iniciar_chrome(headless=headless, manter_navegador=manter_navegador)

        if executar_crediagora:
            fazer_login(driver)

        if args.check_login:
            log("Diagnóstico de login concluído com sucesso; nenhuma exportação foi iniciada.")
            status_log = "check_login"
            return

        arquivos_atualizados = []

        if executar_crediagora:
            if criar_backups:
                criar_backups_iniciais(args.exportacao)
            else:
                log("Backups desativados por --sem-backup.")

            entrar_em_vendas_e_emprestimo(driver)

            if args.exportacao in ("todas", "vendas"):
                arquivos_atualizados.append(processar_exportacao_vendas(driver))

            if args.exportacao == "todas":
                voltar_para_emprestimos(driver)

            if args.exportacao in ("todas", "receita"):
                arquivos_atualizados.append(processar_receita_gerada(driver))

            log("Confirmando que os arquivos finais estão liberados para edição...")
            for arquivo in arquivos_atualizados:
                esperar_arquivo_excel_liberar(arquivo)

            log("Arquivos atualizados e liberados:")
            for arquivo in arquivos_atualizados:
                log(f"- {arquivo}")

        if executar_ercard:
            executar_fase_ercard(driver, config_ercard, log)

        log("========== PROCESSO FINALIZADO COM SUCESSO ==========")
        status_log = "sucesso"

    except KeyboardInterrupt:
        status_log = "interrompido"
        log("========== PROCESSO INTERROMPIDO PELO USUÁRIO ==========")
        log("A execução recebeu Ctrl+C ou foi cancelada antes de terminar.")
        return 130

    except Exception as erro:
        log("========== ERRO NO PROCESSO ==========")
        log(str(erro))

        if driver is not None:
            try:
                salvar_diagnostico_erro(driver, "erro_processo")
            except Exception as erro_diagnostico:
                log(f"Não foi possível salvar o diagnóstico do erro: {erro_diagnostico}")

        return 1

    finally:
        try:
            if driver is not None:
                if status_log in ("sucesso", "check_login"):
                    encerrar_sessao_portal(driver)

                if not manter_navegador:
                    log("Fechando navegador...")
                    driver.quit()
                else:
                    log("Navegador mantido aberto conforme configuração.")
        except Exception as erro:
            log(f"Falha ao finalizar o navegador: {erro}")

        try:
            if status_log in ("check", "check_login"):
                log("Diagnóstico encerrado; downloads preservados.")
            elif manter_downloads:
                log("Arquivos baixados mantidos conforme configuração.")
            elif status_log != "sucesso":
                log("Arquivos baixados mantidos para diagnóstico porque o processo terminou com erro.")
            else:
                log("Limpando arquivos baixados...")
                limpar_downloads()
        except Exception as erro:
            log(f"Não foi possível limpar os downloads finais: {erro}")

        try:
            log("Limpando recursos finais da automação...")
            time.sleep(2)
            log(
                "======= Processo de exportação finalizado: "
                f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} ======="
            )
        except Exception:
            pass

        try:
            salvar_log_ativacao(status=status_log)
        except Exception as erro:
            print(f"Não foi possível salvar o log final: {erro}")

    return 0


if __name__ == "__main__":
    sys.exit(main())
