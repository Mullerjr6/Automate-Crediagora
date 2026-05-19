import os
import time
import shutil
import csv
import re
from pathlib import Path
from datetime import datetime

import pandas as pd
from openpyxl import load_workbook
from openpyxl.formula.translate import Translator
from openpyxl.utils import get_column_letter, range_boundaries

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC


# ============================================================
# CONFIGURAÇÕES PRINCIPAIS
# ============================================================

URL = "https://crediagora.panoramaemprestimos.com.br/login.do?action=sistema"

USUARIO = "junior.muller"

# Coloque a senha como variável de ambiente no Windows:
# setx CREDIAGORA_SENHA "SUA_SENHA_AQUI"
#
# Depois feche e abra o CMD novamente.
SENHA = os.getenv("CREDIAGORA_SENHA", "COLOQUE_SUA_SENHA_AQUI")

# Pasta onde os arquivos exportados serão baixados
PASTA_DOWNLOAD = Path(r"C:\Users\CREDIAGORA\Downloads\crediagora")

# Caminho base informado por você
PASTA_TABELA_FAT = Path(
    r"C:\Users\CREDIAGORA\TJI PROMOTORA DE VENDAS EIRELI\Crediagora-doc - dados\tabela fat"
)

# Arquivos finais
ARQUIVO_FAT_VENDAS = PASTA_TABELA_FAT / "fat_vendas_Teste.xlsx"

# O script tenta encontrar a receita em alguns caminhos possíveis:
# 1) dentro da própria pasta tabela fat
# 2) dentro de uma subpasta chamada tabelas fat
# 3) aceita tanto TESTE quanto Teste no nome
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
    # Mantém o caminho padrão para exibir erro claro depois
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


# ============================================================
# FUNÇÕES DE APOIO
# ============================================================

def log(msg):
    agora = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
    print(f"[{agora}] {msg}")


def criar_pastas():
    PASTA_DOWNLOAD.mkdir(parents=True, exist_ok=True)


def validar_caminhos():
    log("Validando caminhos...")

    if not PASTA_TABELA_FAT.exists():
        raise FileNotFoundError(f"Pasta tabela fat não encontrada: {PASTA_TABELA_FAT}")

    if not ARQUIVO_FAT_VENDAS.exists():
        raise FileNotFoundError(f"Arquivo fat_vendas não encontrado: {ARQUIVO_FAT_VENDAS}")

    if not ARQUIVO_FAT_RECEITA.exists():
        caminhos_tentados = "\n".join(f"- {c}" for c in ARQUIVOS_FAT_RECEITA_POSSIVEIS)
        raise FileNotFoundError(
            "Arquivo fat_receita_gerada não encontrado. "
            f"Tentei estes caminhos:\n{caminhos_tentados}"
        )
    log(f"fat_vendas encontrado: {ARQUIVO_FAT_VENDAS}")
    log(f"fat_receita_gerada encontrado: {ARQUIVO_FAT_RECEITA}")



def limpar_downloads():
    """
    Limpa a pasta de download antes de cada exportação.
    Isso evita o Python confundir arquivo antigo com o novo.
    """
    log("Limpando pasta de downloads...")

    PASTA_DOWNLOAD.mkdir(parents=True, exist_ok=True)

    padroes = [
        "*.xlsx",
        "*.xls",
        "*.xlsm",
        "*.xlsb",
        "*.csv",
        "*.zip",
        "*.crdownload",
        "*.tmp",
        "*.download",
    ]

    apagados = 0

    for padrao in padroes:
        for arquivo in PASTA_DOWNLOAD.glob(padrao):
            try:
                if arquivo.is_file():
                    arquivo.unlink()
                    apagados += 1
                    log(f"Arquivo removido: {arquivo.name}")
            except Exception as erro:
                log(f"Não foi possível apagar {arquivo.name}: {erro}")

    log(f"Limpeza concluída. Arquivos apagados: {apagados}")


def iniciar_chrome():
    opcoes = Options()

    prefs = {
        "download.default_directory": str(PASTA_DOWNLOAD),
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True,
    }

    opcoes.add_experimental_option("prefs", prefs)
    opcoes.add_argument("--start-maximized")

    driver = webdriver.Chrome(options=opcoes)
    return driver


def esperar(driver, segundos=20):
    return WebDriverWait(driver, segundos)


def clicar_por_texto(driver, texto, segundos=20):
    xpath = f"//*[contains(translate(normalize-space(text()), 'ÁÀÃÂÉÊÍÓÔÕÚÇABCDEFGHIJKLMNOPQRSTUVWXYZ', 'AAAAEEIOOOUCabcdefghijklmnopqrstuvwxyz'), '{texto.lower()}')]"

    elemento = esperar(driver, segundos).until(
        EC.element_to_be_clickable((By.XPATH, xpath))
    )

    elemento.click()
    return elemento


def preencher_campo_por_xpath(driver, xpath, texto, segundos=20, limpar=True):
    campo = esperar(driver, segundos).until(
        EC.element_to_be_clickable((By.XPATH, xpath))
    )

    if limpar:
        campo.clear()

    campo.send_keys(texto)
    return campo


def esperar_download_novo(pasta, inicio=None, timeout=300):
    """
    Espera o download terminar de forma robusta.

    Problema corrigido:
    O arquivo já podia estar baixado, mas o script não reconhecia.
    Agora ele:
    - aceita CSV, XLSX, XLSM, XLS, ZIP;
    - ignora .crdownload e temporários;
    - usa horário do clique como referência;
    - confirma se o tamanho do arquivo parou de mudar;
    - se a pasta foi limpa antes, aceita o arquivo mais recente mesmo sem bater 100% o horário.
    """
    pasta = Path(pasta)
    pasta.mkdir(parents=True, exist_ok=True)

    if inicio is None:
        inicio = time.time() - 10

    extensoes_temporarias = (".crdownload", ".tmp", ".download", ".part")
    extensoes_validas = (".csv", ".xlsx", ".xlsm", ".xls", ".xlsb", ".zip")

    def listar_temporarios():
        temporarios = []

        for arquivo in pasta.iterdir():
            try:
                nome = arquivo.name.lower()
                if arquivo.is_file() and nome.endswith(extensoes_temporarias):
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

                if nome.endswith(extensoes_temporarias):
                    continue

                if not nome.endswith(extensoes_validas):
                    continue

                # Aceita arquivos modificados depois do clique.
                # Como limpamos a pasta antes, também aceitamos qualquer arquivo válido que esteja sozinho na pasta.
                if arquivo.stat().st_mtime >= inicio or len(list(pasta.glob("*"))) <= 2:
                    arquivos.append(arquivo)

            except Exception:
                continue

        return arquivos

    def arquivo_estavel(arquivo, tentativas=3, intervalo=1):
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
        "Verifique se o Chrome está baixando em C:\\Users\\CREDIAGORA\\Downloads\\crediagora."
    )


def criar_backup(arquivo):
    arquivo = Path(arquivo)

    esperar_arquivo_excel_liberar(arquivo, timeout=120)

    pasta_backup = arquivo.parent / "backup"
    pasta_backup.mkdir(exist_ok=True)

    data_hora = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    destino = pasta_backup / f"{arquivo.stem}_backup_{data_hora}{arquivo.suffix}"

    shutil.copy2(arquivo, destino)

    return destino

# ============================================================
# LOGIN E NAVEGAÇÃO
# ============================================================

def fazer_login(driver):
    log("Abrindo site...")
    driver.get(URL)

    log("Preenchendo usuário...")
    preencher_campo_por_xpath(
        driver,
        (
            "//input["
            "@type='text' "
            "or contains(@name, 'login') "
            "or contains(@name, 'usuario') "
            "or contains(@name, 'user') "
            "or contains(@id, 'login') "
            "or contains(@id, 'usuario') "
            "or contains(@id, 'user')"
            "]"
        ),
        USUARIO
    )

    log("Preenchendo senha...")
    campo_senha = preencher_campo_por_xpath(
        driver,
        "//input[@type='password']",
        SENHA
    )

    log("Tentando entrar no sistema...")
    entrou = False

    try:
        clicar_por_texto(driver, "Entrar", segundos=8)
        entrou = True
        log("Clique em 'Entrar' realizado pelo texto.")
    except Exception as erro:
        log(f"Não encontrei botão pelo texto 'Entrar': {erro}")

    if not entrou:
        xpaths_botoes = [
            "//button[contains(translate(normalize-space(.), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'entrar')]",
            "//button[contains(translate(normalize-space(.), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'acessar')]",
            "//button[contains(translate(normalize-space(.), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'login')]",
            "//input[@type='submit']",
            "//input[@type='button' and contains(translate(@value, 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'entrar')]",
            "//a[contains(translate(normalize-space(.), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'entrar')]",
            "//a[contains(translate(normalize-space(.), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'acessar')]",
        ]

        for xpath in xpaths_botoes:
            try:
                botao = WebDriverWait(driver, 5).until(
                    EC.element_to_be_clickable((By.XPATH, xpath))
                )
                driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", botao)
                time.sleep(0.3)
                botao.click()
                entrou = True
                log(f"Clique no botão de login realizado usando XPath: {xpath}")
                break
            except Exception:
                continue

    if not entrou:
        try:
            campo_senha.send_keys(Keys.ENTER)
            entrou = True
            log("Login enviado com ENTER no campo de senha.")
        except Exception as erro:
            log(f"Não consegui enviar ENTER no campo de senha: {erro}")

    if not entrou:
        try:
            form = driver.find_element(By.XPATH, "//input[@type='password']/ancestor::form")
            driver.execute_script("arguments[0].submit();", form)
            entrou = True
            log("Formulário de login enviado via JavaScript.")
        except Exception as erro:
            log(f"Não consegui submeter formulário via JavaScript: {erro}")

    log("Aguardando tela principal...")
    try:
        esperar(driver, 60).until(
            EC.presence_of_element_located(
                (By.XPATH, "//*[contains(translate(normalize-space(.), 'ÉÊABCDEFGHIJKLMNOPQRSTUVWXYZ', 'EEabcdefghijklmnopqrstuvwxyz'), 'vendas')]")
            )
        )
        log("Login realizado com sucesso.")
    except Exception as erro:
        screenshot = PASTA_DOWNLOAD / f"erro_login_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.png"
        try:
            driver.save_screenshot(str(screenshot))
            log(f"Print do erro salvo em: {screenshot}")
        except Exception:
            pass

        raise Exception(
            "Não consegui confirmar o login. "
            "Pode ser que o botão Entrar não tenha sido clicado, "
            "que o login/senha estejam incorretos, ou que exista captcha/validação. "
            f"Erro original: {erro}"
        )


def entrar_em_vendas_e_emprestimo(driver):
    """
    Entra corretamente em:
    Vendas > Empréstimo

    Esta função só libera o fluxo depois de confirmar que a tela de Empréstimo
    carregou e que o botão Click para ações está disponível.
    """
    log("Clicando em Vendas...")

    try:
        driver.switch_to.default_content()
    except Exception:
        pass

    clicar_por_texto(driver, "Vendas", segundos=40)
    time.sleep(2)

    log("Clicando em Empréstimo...")

    try:
        driver.switch_to.default_content()
    except Exception:
        pass

    # Tenta clicar em Empréstimo no contexto principal
    clicou_emprestimo = False
    try:
        clicar_por_texto(driver, "Empréstimo", segundos=25)
        clicou_emprestimo = True
    except Exception as erro:
        log(f"Não consegui clicar em Empréstimo no conteúdo principal: {erro}")

    # Se não achou, tenta dentro dos iframes
    if not clicou_emprestimo:
        try:
            driver.switch_to.default_content()
            iframes = driver.find_elements(By.TAG_NAME, "iframe")
            log(f"Procurando Empréstimo em iframes. Total: {len(iframes)}")

            for indice, iframe in enumerate(iframes):
                try:
                    driver.switch_to.default_content()
                    driver.switch_to.frame(iframe)

                    clicar_por_texto(driver, "Empréstimo", segundos=8)
                    clicou_emprestimo = True
                    log(f"Empréstimo clicado no iframe {indice}.")
                    break

                except Exception:
                    continue

        except Exception as erro:
            log(f"Falha procurando Empréstimo em iframes: {erro}")

    if not clicou_emprestimo:
        raise Exception("Não consegui clicar no botão/menu Empréstimo.")

    time.sleep(3)

    # Depois de clicar em Empréstimo, confirma que a tela carregou buscando o botão acoes_ver.
    log("Confirmando carregamento da tela de Empréstimo...")

    encontrou_acoes = False

    try:
        driver.switch_to.default_content()
    except Exception:
        pass

    # Procura o botão Click para ações no conteúdo principal
    try:
        WebDriverWait(driver, 8).until(
            EC.presence_of_element_located((By.ID, "acoes_ver"))
        )
        encontrou_acoes = True
    except Exception:
        pass

    # Procura o botão dentro dos iframes
    if not encontrou_acoes:
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

                    log(f"Tela de Empréstimo confirmada no iframe {indice}.")
                    encontrou_acoes = True
                    break

                except Exception:
                    continue

        except Exception:
            pass

    if not encontrou_acoes:
        screenshot = PASTA_DOWNLOAD / f"erro_emprestimo_nao_carregou_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.png"
        try:
            driver.save_screenshot(str(screenshot))
            log(f"Print do erro salvo em: {screenshot}")
        except Exception:
            pass

        raise Exception(
            "Cliquei em Vendas, mas não consegui confirmar a tela de Empréstimo. "
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
        screenshot = PASTA_DOWNLOAD / f"erro_voltar_para_acoes_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.png"
        html = PASTA_DOWNLOAD / f"erro_voltar_para_acoes_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.html"

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

    screenshot = PASTA_DOWNLOAD / f"erro_click_{texto.replace(' ', '_')}_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.png"
    html = PASTA_DOWNLOAD / f"erro_click_{texto.replace(' ', '_')}_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.html"

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

    screenshot = PASTA_DOWNLOAD / f"erro_click_acoes_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.png"
    html = PASTA_DOWNLOAD / f"erro_click_acoes_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.html"

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
            time.sleep(2)
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
                        time.sleep(2)
                        return

                    except Exception as erro:
                        ultimo_erro = erro
                        continue

            except Exception as erro:
                ultimo_erro = erro
                continue

    except Exception as erro:
        ultimo_erro = erro

    screenshot = PASTA_DOWNLOAD / f"erro_exportacoes_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.png"
    html = PASTA_DOWNLOAD / f"erro_exportacoes_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.html"

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
                time.sleep(2)
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

    screenshot = PASTA_DOWNLOAD / f"erro_layout_arquivo_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.png"
    html = PASTA_DOWNLOAD / f"erro_layout_arquivo_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.html"

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
    clicar_click_para_acoes(driver)
    clicar_exportacoes(driver)
    clicar_exportar_layout_arquivo(driver)


def selecionar_layout(driver, nome_exportacao):
    """
    Preenche e seleciona o layout correto no autocomplete.

    Campo correto:
    <input type="text" name="layoutArquivo" class="txt100 ac_field"
           id="_id_layoutArquivo" autocomplete="off">

    A função evita campo.click(), pois o painel do autocomplete pode interceptar
    o clique. Ela foca via JavaScript, digita com send_keys e seleciona a opção
    correta do autocomplete.
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

    campo = procurar_campo_no_contexto("contexto atual")

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
        screenshot = PASTA_DOWNLOAD / f"erro_campo_layout_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.png"
        html = PASTA_DOWNLOAD / f"erro_campo_layout_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.html"

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

    try:
        driver.switch_to.active_element.send_keys(Keys.ESCAPE)
        time.sleep(0.3)
    except Exception:
        pass

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

    try:
        valor_atual = campo.get_attribute("value")
        log(f"Valor atual no campo layoutArquivo: {valor_atual}")

        if not valor_atual:
            raise Exception("O campo layoutArquivo ficou vazio após o preenchimento.")

    except Exception as erro:
        screenshot = PASTA_DOWNLOAD / f"erro_valor_layout_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.png"
        html = PASTA_DOWNLOAD / f"erro_valor_layout_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.html"

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

    Procura links de arquivo que contenham idDigitalizacao ou exibirDigitalizacao.
    Se conseguir ler datas nas linhas, escolhe a maior data. Caso contrário,
    clica no primeiro link encontrado.
    """
    log("Localizando arquivo mais recente para download...")

    padrao_data = re.compile(
        r"(\d{2}/\d{2}/\d{4})(?:\s+(\d{2}:\d{2}(?::\d{2})?))?"
    )

    formatos_data = [
        "%d/%m/%Y %H:%M:%S",
        "%d/%m/%Y %H:%M",
        "%d/%m/%Y",
    ]

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

    def tentar_no_contexto(nome_contexto):
        links = driver.find_elements(
            By.XPATH,
            "//a[contains(@href, 'idDigitalizacao') "
            "or contains(@href, 'exibirDigitalizacao') "
            "or contains(@href, 'layoutArquivo.do')]"
        )

        links_filtrados = []
        for link in links:
            try:
                href = link.get_attribute("href") or ""
                html = link.get_attribute("outerHTML") or ""

                if (
                    "idDigitalizacao" in href
                    or "exibirDigitalizacao" in href
                    or "fa-file-alt" in html
                    or "file-alt" in html
                ):
                    links_filtrados.append(link)
            except Exception:
                continue

        links = links_filtrados

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

            data_linha = converter_data(texto_linha)

            candidatos.append({
                "indice": indice,
                "link": link,
                "data": data_linha,
                "texto": texto_linha,
                "href": link.get_attribute("href"),
                "title": link.get_attribute("title"),
            })

        candidatos_com_data = [c for c in candidatos if c["data"] is not None]

        if candidatos_com_data:
            candidatos_com_data.sort(key=lambda c: c["data"], reverse=True)
            escolhido = candidatos_com_data[0]
            log(
                "Arquivo mais recente definido pela data da linha: "
                f"{escolhido['data'].strftime('%d/%m/%Y %H:%M:%S')}"
            )
        else:
            escolhido = candidatos[0]
            log(
                "Não consegui identificar a data das linhas. "
                "Vou clicar no primeiro arquivo encontrado."
            )

        log(f"Href escolhido: {escolhido.get('href')}")
        log(f"Title escolhido: {escolhido.get('title')}")
        log(f"Texto da linha escolhida: {escolhido.get('texto')}")

        elemento = escolhido["link"]

        driver.execute_script(
            "arguments[0].scrollIntoView({block: 'center', inline: 'center'});",
            elemento
        )
        time.sleep(0.5)

        try:
            elemento.click()
        except Exception:
            driver.execute_script("arguments[0].click();", elemento)

        log("Clique no arquivo mais recente realizado.")
        time.sleep(1)
        return True

    try:
        if tentar_no_contexto("contexto atual"):
            return
    except Exception as erro:
        log(f"Falha ao buscar download no contexto atual: {erro}")

    try:
        driver.switch_to.default_content()
        iframes = driver.find_elements(By.TAG_NAME, "iframe")
        log(f"Procurando links de download em iframes. Total: {len(iframes)}")

        for indice, iframe in enumerate(iframes):
            try:
                driver.switch_to.default_content()
                driver.switch_to.frame(iframe)

                if tentar_no_contexto(f"iframe {indice}"):
                    return

            except Exception as erro:
                log(f"Não encontrei download no iframe {indice}: {erro}")
                continue

    except Exception as erro:
        log(f"Falha ao procurar download em iframes: {erro}")

    screenshot = PASTA_DOWNLOAD / f"erro_download_arquivo_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.png"
    html = PASTA_DOWNLOAD / f"erro_download_arquivo_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.html"

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
        "Não encontrei o link correto de download com idDigitalizacao/exibirDigitalizacao."
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

    screenshot = PASTA_DOWNLOAD / f"erro_executar_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.png"
    html = PASTA_DOWNLOAD / f"erro_executar_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.html"

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


def baixar_exportacao(driver, nome_exportacao):
    limpar_downloads()

    abrir_tela_exportacao(driver)
    selecionar_layout(driver, nome_exportacao)

    clicar_executar_exportacao(driver)
    time.sleep(3)

    log("Aguardando lista de arquivos...")
    esperar(driver, 90).until(
        EC.presence_of_element_located(
            (
                By.XPATH,
                "//*[contains(text(), 'Data') or contains(text(), 'Código') or contains(text(), 'Codigo')]"
                " | //a[contains(@href, 'layoutArquivo.do') and contains(@href, 'exibirDigitalizacao')]"
                " | //a[contains(@href, 'idDigitalizacao')]"
            )
        )
    )

    inicio_download = time.time() - 5

    log("Clicando no download mais recente...")
    clicar_primeiro_download(driver)

    log("Aguardando download finalizar...")
    arquivo_baixado = esperar_download_novo(PASTA_DOWNLOAD, inicio=inicio_download, timeout=300)

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


def ajustar_tabelas_excel(ws):
    """
    Expande Tabelas do Excel para incluir as linhas atuais.
    """
    if not ws.tables:
        return

    for tabela in ws.tables.values():
        min_col, min_row, max_col, _ = range_boundaries(tabela.ref)
        nova_ref = (
            f"{get_column_letter(min_col)}{min_row}:"
            f"{get_column_letter(max_col)}{ws.max_row}"
        )
        tabela.ref = nova_ref


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

        wb_origem = load_workbook(arquivo_origem, data_only=True)
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



def esperar_arquivo_excel_liberar(arquivo, timeout=120):
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

    if "esperar_arquivo_excel_liberar" in globals():
        esperar_arquivo_excel_liberar(arquivo_destino, timeout=120)

    backup = criar_backup(arquivo_destino)
    log(f"Backup criado: {backup}")

    linhas_origem = ler_linhas_exportacao(arquivo_origem)

    wb_destino = None

    try:
        wb_destino = load_workbook(arquivo_destino)
        ws_destino = wb_destino.active

        modelos_formula = obter_modelos_formulas(ws_destino, colunas_formula)

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

        aplicar_formulas(
            ws_destino,
            modelos_formula,
            primeira_linha_colagem,
            ultima_linha_colada
        )

        ajustar_tabelas_excel(ws_destino)

        if "esperar_arquivo_excel_liberar" in globals():
            esperar_arquivo_excel_liberar(arquivo_destino, timeout=120)

        wb_destino.save(arquivo_destino)
        log(f"Arquivo salvo com sucesso: {arquivo_destino}")

    finally:
        if wb_destino:
            wb_destino.close()


# ============================================================
# PROCESSO PRINCIPAL
# ============================================================

def main():
    criar_pastas()
    validar_caminhos()

    driver = iniciar_chrome()

    try:
        fazer_login(driver)
        entrar_em_vendas_e_emprestimo(driver)

        # ====================================================
        # PARTE 1 - EXPORTAÇÃO VENDAS
        # ====================================================
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
        if "esperar_arquivo_excel_liberar" in globals():
            log("Aguardando arquivo de vendas ser liberado para edição...")
            esperar_arquivo_excel_liberar(ARQUIVO_FAT_VENDAS, timeout=120)
            log("Arquivo de vendas liberado. Você pode abrir a planilha agora.")

        log("========== EXPORTAÇÃO VENDAS CONCLUÍDA ==========")

        # ====================================================
        # PREPARAR TELA PARA A RECEITA
        # ====================================================
        voltar_para_emprestimos(driver)

        # ====================================================
        # PARTE 2 - RECEITA GERADA
        # ====================================================
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

        if "esperar_arquivo_excel_liberar" in globals():
            log("Aguardando arquivo de receita ser liberado para edição...")
            esperar_arquivo_excel_liberar(ARQUIVO_FAT_RECEITA, timeout=120)
            log("Arquivo de receita liberado. Você pode abrir a planilha agora.")

        log("========== RECEITA GERADA CONCLUÍDA ==========")

        log("========== PROCESSO FINALIZADO COM SUCESSO ==========")

    except Exception as erro:
        log("========== ERRO NO PROCESSO ==========")
        log(str(erro))
        raise

    else:
        if "esperar_arquivo_excel_liberar" in globals():
            log("Aguardando arquivos Excel serem liberados para edição...")
            esperar_arquivo_excel_liberar(ARQUIVO_FAT_VENDAS, timeout=120)
            esperar_arquivo_excel_liberar(ARQUIVO_FAT_RECEITA, timeout=120)
            log("Arquivos liberados. Você pode abrir as planilhas agora.")

    finally:
        try:
            time.sleep(3)
            driver.quit()
        except Exception:
            pass
    try:
        log("Processo concluído. Verifique os arquivos atualizados:")
        log(f"- Vendas: {ARQUIVO_FAT_VENDAS}")
        log(f"- Receita: {ARQUIVO_FAT_RECEITA}")
    except Exception:
        pass

if __name__ == "__main__":
    main()

log("Script finalizado.")

log("Observação: Se os arquivos Excel não abrirem ou apresentarem erro de leitura, "
    "certifique-se de fechar as planilhas no Excel e aguardar alguns segundos para que o OneDrive/SharePoint sincronize as alterações. "
    "Se necessário, reinicie o processo para garantir que os arquivos sejam atualizados corretamente.") 
criar_backup(ARQUIVO_FAT_VENDAS)
criar_backup(ARQUIVO_FAT_RECEITA)
log("Backups dos arquivos de destino criados para segurança.")
