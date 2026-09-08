import csv
import os
import re
import time
import unicodedata
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path

from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait


URL_ERCARD = "https://app.ercard.com.br/"
DATA_INICIAL = "01/01/2024"
CONSULTA_CONTRATOS = "14 - TABELA CONTRATOS"
FORMATO_CSV = "ARQUIVO CSV (EXCEL)"


class ErCardError(RuntimeError):
    def __init__(self, etapa, mensagem):
        self.etapa = etapa
        super().__init__(f"[ERCARD][ERRO][{etapa}] {mensagem}")


@dataclass(frozen=True)
class ErCardConfig:
    portal_usuario: str
    portal_senha: str
    sistema_usuario: str
    sistema_senha: str
    pasta_exportacao: Path
    pasta_erros: Path
    timeout_normal: int = 30
    timeout_remoto: int = 120
    timeout_exportacao: int = 180

    def validar(self):
        ausentes = []
        campos = {
            "ERCARD_PORTAL_USUARIO": self.portal_usuario,
            "ERCARD_PORTAL_SENHA": self.portal_senha,
            "ERCARD_SISTEMA_USUARIO": self.sistema_usuario,
            "ERCARD_SISTEMA_SENHA": self.sistema_senha,
        }
        for nome, valor in campos.items():
            if not str(valor).strip():
                ausentes.append(nome)

        if ausentes:
            raise EnvironmentError(
                "Credenciais ERCard não configuradas: " + ", ".join(ausentes)
            )


def normalizar(texto):
    texto = unicodedata.normalize("NFD", str(texto or ""))
    return "".join(
        caractere for caractere in texto
        if unicodedata.category(caractere) != "Mn"
    ).casefold().strip()


def log_ercard(log, mensagem):
    log(f"[ERCARD] {mensagem}")


def gerar_nome_tabela_contratos(agora=None):
    agora = agora or datetime.now()
    return f"Tabela contratos {agora.strftime('%d-%m-%Y %H-%M')}.csv"


def caminho_exportacao_unico(pasta, nome):
    pasta = Path(pasta)
    candidato = pasta / nome
    contador = 2

    while candidato.exists():
        candidato = pasta / f"{Path(nome).stem} ({contador}).csv"
        contador += 1

    return candidato


def _elementos_em_contextos(driver, by, seletor):
    driver.switch_to.default_content()
    contextos = [("conteúdo principal", None)]
    contextos.extend(
        (f"iframe {indice}", iframe)
        for indice, iframe in enumerate(driver.find_elements(By.TAG_NAME, "iframe"))
    )

    for nome, iframe in contextos:
        try:
            driver.switch_to.default_content()
            if iframe is not None:
                driver.switch_to.frame(iframe)
            for elemento in driver.find_elements(by, seletor):
                if elemento.is_displayed():
                    yield nome, elemento
        except Exception:
            continue

    driver.switch_to.default_content()


def _primeiro_elemento(driver, seletores, timeout, etapa):
    fim = time.monotonic() + timeout
    while time.monotonic() < fim:
        for by, seletor in seletores:
            encontrados = list(_elementos_em_contextos(driver, by, seletor))
            if encontrados:
                return encontrados[0][1]
        time.sleep(0.4)
    raise ErCardError(etapa, "Elemento esperado não foi encontrado.")


def _clicar_texto_web(driver, texto, timeout, etapa):
    alvo = normalizar(texto)
    fim = time.monotonic() + timeout
    tags = "a,button,[role='button'],div,span"

    while time.monotonic() < fim:
        candidatos = list(_elementos_em_contextos(driver, By.CSS_SELECTOR, tags))
        candidatos_validos = []
        for _, elemento in candidatos:
            try:
                conteudo = normalizar(elemento.text)
                if conteudo == alvo and elemento.is_enabled():
                    candidatos_validos.append(elemento)
            except Exception:
                continue

        if candidatos_validos:
            candidatos_validos.sort(
                key=lambda elemento: elemento.size["width"] * elemento.size["height"]
            )
            elemento = candidatos_validos[0]
            driver.execute_script(
                "arguments[0].scrollIntoView({block:'center'});", elemento
            )
            elemento.click()
            return
        time.sleep(0.5)

    raise ErCardError(etapa, f"Texto não encontrado: {texto}")


def abrir_ercard(driver, config, log):
    log_ercard(log, "Iniciando fase ERCard...")
    log_ercard(log, "Abrindo portal...")
    driver.switch_to.default_content()
    driver.get(URL_ERCARD)
    WebDriverWait(driver, config.timeout_normal).until(
        lambda navegador: navegador.execute_script("return document.readyState") == "complete"
    )


def realizar_login_portal(driver, config, log):
    log_ercard(log, "Verificando autenticação...")
    campos_usuario = list(
        _elementos_em_contextos(
            driver,
            By.XPATH,
            "//input[@type='text' or @type='email']"
        )
    )
    campos_senha = list(
        _elementos_em_contextos(driver, By.CSS_SELECTOR, "input[type='password']")
    )

    if not campos_usuario and not campos_senha:
        log_ercard(log, "Login 1 já autenticado; formulário não exibido.")
        return

    if not campos_usuario or not campos_senha:
        raise ErCardError("Login 1", "Tela de login incompleta.")

    campo_usuario = campos_usuario[0][1]
    campo_senha = campos_senha[0][1]
    valor_usuario = (campo_usuario.get_attribute("value") or "").strip()
    if valor_usuario != config.portal_usuario:
        campo_usuario.clear()
        campo_usuario.send_keys(config.portal_usuario)

    if not (campo_senha.get_attribute("value") or ""):
        campo_senha.send_keys(config.portal_senha)

    if (campo_usuario.get_attribute("value") or "").strip() != config.portal_usuario:
        raise ErCardError("Login 1", "O usuário não foi preenchido integralmente.")
    if not (campo_senha.get_attribute("value") or ""):
        raise ErCardError("Login 1", "A senha não foi preenchida.")

    botao = _primeiro_elemento(
        driver,
        [
            (By.XPATH, "//button[normalize-space(.)='Entrar']"),
            (By.CSS_SELECTOR, "button[type='submit']"),
        ],
        config.timeout_normal,
        "Login 1"
    )
    botao.click()

    def resultado_login(navegador):
        if not navegador.find_elements(By.CSS_SELECTOR, "input[type='password']"):
            return "sucesso"
        texto = normalizar(navegador.find_element(By.TAG_NAME, "body").text)
        if any(termo in texto for termo in ("invalido", "incorreto", "bloqueado", "erro")):
            return "erro"
        return False

    resultado = WebDriverWait(driver, config.timeout_normal).until(resultado_login)
    if resultado != "sucesso":
        raise ErCardError("Login 1", "O portal recusou a autenticação.")
    log_ercard(log, "Login 1 realizado.")


def aguardar_conexao_ercard(driver, config, log):
    log_ercard(log, "Conexão sendo estabelecida...")
    log_ercard(log, "Aguardando carregamento do ambiente...")
    fim = time.monotonic() + config.timeout_remoto

    while time.monotonic() < fim:
        try:
            texto = normalizar(driver.find_element(By.TAG_NAME, "body").text)
            conectando = "aguarde enquanto a conexao" in texto
            pronto = (
                "crediagora" in texto
                or "er cartao crediagora" in texto
                or "canvas" in driver.page_source.casefold()
                or "html5.html" in driver.current_url.casefold()
            )
            if pronto and not conectando:
                return
        except Exception:
            pass
        time.sleep(0.5)

    raise ErCardError("Conexão", "Timeout de 120 segundos durante a conexão.")


def selecionar_ambiente_crediagora(driver, config, log):
    log_ercard(log, "Selecionando ambiente CrediAgora...")
    try:
        _clicar_texto_web(
            driver, "CrediAgora", config.timeout_normal, "Ambiente CrediAgora"
        )
        return True
    except ErCardError:
        return False


def abrir_er_cartao(driver, config, log):
    log_ercard(log, "Abrindo ER Cartão CrediAgora...")
    janelas_antes = set(driver.window_handles)
    try:
        _clicar_texto_web(
            driver,
            "ER Cartão CrediAgora",
            config.timeout_normal,
            "Módulo ER Cartão"
        )
        fim = time.monotonic() + config.timeout_normal
        while time.monotonic() < fim:
            novas = [janela for janela in driver.window_handles if janela not in janelas_antes]
            if novas:
                driver.switch_to.window(novas[-1])
                break
            time.sleep(0.3)
        return True
    except ErCardError:
        return False


class InterfaceWindowsErCard:
    def __init__(self, config, log):
        try:
            from pywinauto import Desktop
        except ImportError as erro:
            raise ErCardError(
                "Dependências",
                "pywinauto não está instalado. Execute pip install -r requirements.txt."
            ) from erro

        self.desktop = Desktop(backend="uia")
        self.config = config
        self.log = log

    def aguardar_janela(self, padrao_titulo, timeout=None):
        timeout = timeout or self.config.timeout_remoto
        regex = re.compile(padrao_titulo, re.IGNORECASE)
        fim = time.monotonic() + timeout

        while time.monotonic() < fim:
            for janela in self.desktop.windows():
                try:
                    if regex.search(janela.window_text() or "") and janela.is_visible():
                        return janela
                except Exception:
                    continue
            time.sleep(0.5)

        raise ErCardError("Interface remota", f"Janela não encontrada: {padrao_titulo}")

    @staticmethod
    def _controles(janela, tipos):
        controles = []
        for tipo in tipos:
            try:
                controles.extend(janela.descendants(control_type=tipo))
            except Exception:
                continue
        return controles

    def controle_texto(self, janela, texto, tipos=("Button", "Text", "ListItem", "MenuItem")):
        alvo = normalizar(texto)
        exatos = []
        parciais = []
        for controle in self._controles(janela, tipos):
            try:
                atual = normalizar(controle.window_text())
                if atual == alvo:
                    exatos.append(controle)
                elif alvo in atual:
                    parciais.append(controle)
            except Exception:
                continue

        encontrados = exatos or parciais
        if not encontrados:
            raise ErCardError("Interface remota", f"Controle não encontrado: {texto}")
        return encontrados[0]

    def campo_ao_lado(self, janela, rotulo, tipos=("Edit", "ComboBox")):
        label = self.controle_texto(janela, rotulo, tipos=("Text",))
        caixa_label = label.rectangle()
        candidatos = []
        for controle in self._controles(janela, tipos):
            try:
                caixa = controle.rectangle()
                distancia_y = abs(caixa.mid_point().y - caixa_label.mid_point().y)
                if caixa.left >= caixa_label.left and distancia_y <= max(35, caixa.height()):
                    candidatos.append((distancia_y, caixa.left, controle))
            except Exception:
                continue
        if not candidatos:
            raise ErCardError("Interface remota", f"Campo não encontrado: {rotulo}")
        candidatos.sort(key=lambda item: (item[0], item[1]))
        return candidatos[0][2]

    @staticmethod
    def preencher(controle, valor):
        controle.set_focus()
        try:
            controle.set_edit_text(valor)
        except Exception:
            controle.type_keys("^a{BACKSPACE}", set_foreground=True)
            controle.type_keys(valor, with_spaces=True, set_foreground=True)

    def clicar(self, controle):
        controle.set_focus()
        controle.click_input()

    def selecionar_opcao(self, controle, texto):
        try:
            controle.select(texto)
        except Exception:
            controle.set_focus()
            controle.type_keys("^a" + texto + "{ENTER}", with_spaces=True)

    def selecionar_texto_em_janela(self, titulo, texto, timeout=None, duplo=False):
        janela = self.aguardar_janela(titulo, timeout)
        controle = self.controle_texto(janela, texto)
        controle.set_focus()
        if duplo:
            controle.double_click_input()
        else:
            controle.click_input()
        return janela


def realizar_login_sistema(ui, config, log):
    log_ercard(log, "Realizando Login 2...")
    janela = ui.aguardar_janela(r"Logon|ER Card.*Gestão de Cartões")
    empresa = ui.campo_ao_lado(janela, "Empresa")
    ui.selecionar_opcao(empresa, "01 - FARE")
    ui.preencher(ui.campo_ao_lado(janela, "Usuário", tipos=("Edit",)), config.sistema_usuario)
    ui.preencher(ui.campo_ao_lado(janela, "Senha", tipos=("Edit",)), config.sistema_senha)
    ui.clicar(ui.controle_texto(janela, "Ok", tipos=("Button",)))
    ui.aguardar_janela(r"ER Cred.*Gestão de Operações.*FARE.*01", config.timeout_remoto)
    log_ercard(log, "ER Cred carregado.")


def abrir_consulta_gerencial(ui, config, log):
    log_ercard(log, "Abrindo Consultas > Gerencial...")
    janela = ui.aguardar_janela(r"ER Cred.*Gestão de Operações")
    try:
        janela.menu_select("Consultas->Gerencial")
        return
    except Exception:
        pass

    ui.clicar(ui.controle_texto(janela, "Consultas", tipos=("MenuItem", "Text")))
    gerencial = ui.controle_texto(janela, "Gerencial", tipos=("MenuItem", "Text"))
    ui.clicar(gerencial)


def selecionar_tabela_contratos(ui, config, log):
    log_ercard(log, f"Selecionando {CONSULTA_CONTRATOS}...")
    janela = ui.aguardar_janela(r"Geração de Dados Gerenciais")
    campo = ui.campo_ao_lado(janela, "Pesquisa Padrão para Seleção dos Dados")
    caixa = campo.rectangle()
    botoes = []
    for botao in ui._controles(janela, ("Button",)):
        try:
            ret = botao.rectangle()
            if ret.left >= caixa.right and abs(ret.mid_point().y - caixa.mid_point().y) <= 30:
                botoes.append((ret.left, botao))
        except Exception:
            continue
    if not botoes:
        raise ErCardError("Consulta", "Botão de seleção da pesquisa não encontrado.")
    botoes.sort(key=lambda item: item[0])
    ui.clicar(botoes[0][1])

    relacao = ui.aguardar_janela(r"Relação de Consultas", config.timeout_normal)
    consulta = ui.controle_texto(relacao, CONSULTA_CONTRATOS, tipos=("ListItem", "Text"))
    consulta.double_click_input()
    time.sleep(0.5)
    valor = normalizar(ui.campo_ao_lado(janela, "Pesquisa Padrão para Seleção dos Dados").window_text())
    if normalizar(CONSULTA_CONTRATOS) not in valor:
        raise ErCardError("Consulta", "A consulta 14 não ficou selecionada.")


def configurar_exportacao_csv(ui, config, arquivo_webfile, log):
    log_ercard(log, f"Selecionando {FORMATO_CSV}...")
    janela = ui.aguardar_janela(r"Geração de Dados Gerenciais")
    saida = ui.campo_ao_lado(janela, "Gerar saída dos dados para")
    ui.selecionar_opcao(saida, FORMATO_CSV)

    campo_arquivo = ui.campo_ao_lado(janela, "Arquivo para Exportação", tipos=("Edit",))
    caixa = campo_arquivo.rectangle()
    botoes = []
    for botao in ui._controles(janela, ("Button",)):
        try:
            ret = botao.rectangle()
            if ret.left >= caixa.right and abs(ret.mid_point().y - caixa.mid_point().y) <= 30:
                botoes.append((ret.left, botao))
        except Exception:
            continue
    if not botoes:
        raise ErCardError("Arquivo para Exportação", "Botão seletor não encontrado.")
    ui.clicar(sorted(botoes, key=lambda item: item[0])[0][1])

    for numero in (1, 2):
        log_ercard(log, f"Confirmando Restrictions {numero}/2...")
        aviso = ui.aguardar_janela(r"Restrictions", config.timeout_normal)
        ui.clicar(ui.controle_texto(aviso, "OK", tipos=("Button",)))
        time.sleep(0.5)

    janela_open = ui.aguardar_janela(r"^Open$|^Abrir$", config.timeout_normal)
    this_pc = ui.controle_texto(janela_open, "This PC", tipos=("TreeItem",))
    try:
        this_pc.collapse()
    except Exception:
        pass
    time.sleep(0.5)
    try:
        this_pc.expand()
    except Exception as erro:
        raise ErCardError("WebFile", "Não foi possível expandir This PC.") from erro

    webfiles = []
    for item in ui._controles(janela_open, ("TreeItem", "ListItem")):
        try:
            if normalizar(item.window_text()).startswith("webfile"):
                webfiles.append(item)
        except Exception:
            continue
    if not webfiles:
        raise ErCardError("WebFile", "Unidade iniciada por WebFile não encontrada.")
    ui.clicar(webfiles[0])
    log_ercard(log, "WebFile selecionado.")

    nome = Path(arquivo_webfile).name
    campo_nome = ui.campo_ao_lado(janela_open, "File name", tipos=("Edit",))
    ui.preencher(campo_nome, nome)
    ui.clicar(ui.controle_texto(janela_open, "Open", tipos=("Button",)))

    fim = time.monotonic() + config.timeout_normal
    while time.monotonic() < fim:
        valor = campo_arquivo.window_text().strip()
        if valor:
            log_ercard(log, f"Nome do arquivo: {nome}")
            return
        time.sleep(0.5)
    raise ErCardError("Arquivo para Exportação", "O caminho retornou vazio.")


def preencher_periodo(ui, config, log):
    janela = ui.aguardar_janela(r"Geração de Dados Gerenciais")
    data_final = datetime.now().strftime("%d/%m/%Y")
    ui.preencher(ui.campo_ao_lado(janela, "DATA INICIAL", tipos=("Edit",)), DATA_INICIAL)
    ui.preencher(ui.campo_ao_lado(janela, "DATA FINAL", tipos=("Edit",)), data_final)
    log_ercard(log, f"Data inicial: {DATA_INICIAL}")
    log_ercard(log, f"Data final: {data_final}")


def exportar_tabela_contratos(ui, config, arquivo_final, log):
    janela = ui.aguardar_janela(r"Geração de Dados Gerenciais")
    log_ercard(log, "Iniciando exportação...")
    ui.clicar(ui.controle_texto(janela, "Exportar", tipos=("Button",)))
    log_ercard(log, "Exportando Tabela de Contratos...")
    log_ercard(log, "Aguardando geração do CSV...")

    salvar = ui.aguardar_janela(r"Save As|Salvar Como|Salvar como", config.timeout_exportacao)
    campo_nome = None
    try:
        campo_nome = ui.campo_ao_lado(salvar, "File name", tipos=("Edit",))
    except ErCardError:
        campo_nome = ui.campo_ao_lado(salvar, "Nome", tipos=("Edit",))
    ui.preencher(campo_nome, str(arquivo_final))
    try:
        botao = ui.controle_texto(salvar, "Save", tipos=("Button",))
    except ErCardError:
        botao = ui.controle_texto(salvar, "Salvar", tipos=("Button",))
    ui.clicar(botao)


def validar_csv_exportado(arquivo, pasta_esperada, timeout, log):
    arquivo = Path(arquivo)
    pasta_esperada = Path(pasta_esperada).resolve()
    fim = time.monotonic() + timeout
    ultimo_tamanho = -1
    estavel = 0

    while time.monotonic() < fim:
        if arquivo.exists() and arquivo.suffix.casefold() == ".csv":
            tamanho = arquivo.stat().st_size
            if tamanho > 0 and tamanho == ultimo_tamanho:
                estavel += 1
                if estavel >= 3:
                    break
            else:
                estavel = 0
            ultimo_tamanho = tamanho
        time.sleep(1)
    else:
        raise ErCardError("Validação", "Arquivo final não foi localizado ou estabilizado.")

    if arquivo.parent.resolve() != pasta_esperada:
        raise ErCardError("Validação", "Arquivo salvo fora da pasta configurada.")
    if not arquivo.name.casefold().startswith("tabela contratos"):
        raise ErCardError("Validação", "Nome final inesperado.")

    try:
        with arquivo.open("r", encoding="utf-8-sig", errors="replace", newline="") as csv_file:
            next(csv.reader(csv_file), None)
    except OSError as erro:
        raise ErCardError("Validação", "CSV não pôde ser lido.") from erro

    log_ercard(log, "Exportação concluída com sucesso.")
    log_ercard(log, f"Arquivo validado: {arquivo}")
    return arquivo


def salvar_diagnostico_ercard(driver, config, etapa, log):
    config.pasta_erros.mkdir(parents=True, exist_ok=True)
    nome = re.sub(r"[^a-zA-Z0-9_-]+", "_", etapa).strip("_") or "ercard"
    caminho = config.pasta_erros / f"ercard_{nome}_{datetime.now():%Y-%m-%d_%H-%M-%S}.png"
    try:
        driver.save_screenshot(str(caminho))
        log_ercard(log, f"Screenshot de diagnóstico: {caminho}")
    except Exception as erro:
        log_ercard(log, f"Não foi possível salvar screenshot: {erro}")


def executar_fase_ercard(driver, config, log):
    config.validar()
    config.pasta_exportacao.mkdir(parents=True, exist_ok=True)
    config.pasta_erros.mkdir(parents=True, exist_ok=True)
    nome = gerar_nome_tabela_contratos()
    arquivo_final = caminho_exportacao_unico(config.pasta_exportacao, nome)

    try:
        abrir_ercard(driver, config, log)
        realizar_login_portal(driver, config, log)
        aguardar_conexao_ercard(driver, config, log)

        ambiente_web = selecionar_ambiente_crediagora(driver, config, log)
        modulo_web = abrir_er_cartao(driver, config, log) if ambiente_web else False

        ui = InterfaceWindowsErCard(config, log)
        if not ambiente_web:
            ui.selecionar_texto_em_janela(r"ER Systems", "CrediAgora")
        if not modulo_web:
            ui.selecionar_texto_em_janela(r"ER Systems", "ER Cartão CrediAgora")

        realizar_login_sistema(ui, config, log)
        abrir_consulta_gerencial(ui, config, log)
        selecionar_tabela_contratos(ui, config, log)
        configurar_exportacao_csv(ui, config, arquivo_final, log)
        preencher_periodo(ui, config, log)
        exportar_tabela_contratos(ui, config, arquivo_final, log)
        return validar_csv_exportado(
            arquivo_final,
            config.pasta_exportacao,
            config.timeout_exportacao,
            log
        )
    except Exception as erro:
        etapa = erro.etapa if isinstance(erro, ErCardError) else "fase_ercard"
        salvar_diagnostico_ercard(driver, config, etapa, log)
        if isinstance(erro, ErCardError):
            raise
        raise ErCardError(etapa, str(erro)) from erro
