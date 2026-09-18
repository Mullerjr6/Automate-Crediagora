import csv
import os
import re
import shutil
import time
import unicodedata
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path

from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait


URL_ERCARD = "https://app.ercard.com.br/"
DATA_INICIAL = "010124"
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
    pasta_download: Path | None = None
    timeout_normal: int = 30
    timeout_remoto: int = 20
    timeout_exportacao: int = 300

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


def _estado_janela_windows(hwnd):
    import ctypes
    from ctypes import wintypes

    user32 = ctypes.windll.user32

    class RECT(ctypes.Structure):
        _fields_ = [
            ("left", wintypes.LONG),
            ("top", wintypes.LONG),
            ("right", wintypes.LONG),
            ("bottom", wintypes.LONG),
        ]

    class MONITORINFOEXW(ctypes.Structure):
        _fields_ = [
            ("cbSize", wintypes.DWORD),
            ("rcMonitor", RECT),
            ("rcWork", RECT),
            ("dwFlags", wintypes.DWORD),
            ("szDevice", wintypes.WCHAR * 32),
        ]

    def tupla(retangulo):
        return (
            retangulo.left,
            retangulo.top,
            retangulo.right,
            retangulo.bottom,
        )

    retangulo = RECT()
    if not user32.GetWindowRect(hwnd, ctypes.byref(retangulo)):
        raise ErCardError("Janela", "GetWindowRect falhou para a janela do ER Card.")

    visivel = RECT()
    try:
        dwmapi = ctypes.windll.dwmapi
        DWMWA_EXTENDED_FRAME_BOUNDS = 9
        resultado = dwmapi.DwmGetWindowAttribute(
            hwnd,
            DWMWA_EXTENDED_FRAME_BOUNDS,
            ctypes.byref(visivel),
            ctypes.sizeof(visivel),
        )
        if resultado != 0:
            visivel = retangulo
    except Exception:
        visivel = retangulo

    monitores = []
    callback_tipo = ctypes.WINFUNCTYPE(
        wintypes.BOOL, wintypes.HMONITOR, wintypes.HDC, ctypes.POINTER(RECT), wintypes.LPARAM
    )

    @callback_tipo
    def callback_monitor(hmonitor, _hdc, _rect, _dados):
        info = MONITORINFOEXW()
        info.cbSize = ctypes.sizeof(info)
        if user32.GetMonitorInfoW(hmonitor, ctypes.byref(info)):
            monitor = tupla(info.rcMonitor)
            area_trabalho = tupla(info.rcWork)
            janela = tupla(visivel)
            intersecta = not (
                janela[2] <= monitor[0]
                or janela[0] >= monitor[2]
                or janela[3] <= monitor[1]
                or janela[1] >= monitor[3]
            )
            if intersecta:
                monitores.append(
                    {
                        "handle": int(hmonitor),
                        "dispositivo": info.szDevice,
                        "principal": bool(info.dwFlags & 1),
                        "retangulo": monitor,
                        "area_trabalho": area_trabalho,
                    }
                )
        return True

    user32.EnumDisplayMonitors(None, None, callback_monitor, 0)
    return {
        "retangulo": tupla(retangulo),
        "retangulo_visivel": tupla(visivel),
        "monitores": monitores,
        "maximizada": bool(user32.IsZoomed(hwnd)),
    }


def localizar_janela_ercard_windows(titulo_esperado=""):
    import ctypes
    from ctypes import wintypes

    user32 = ctypes.windll.user32
    candidatos = []
    callback_tipo = ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)

    @callback_tipo
    def callback(hwnd, _dados):
        if not user32.IsWindowVisible(hwnd):
            return True
        tamanho = user32.GetWindowTextLengthW(hwnd)
        if tamanho <= 0:
            return True
        titulo = ctypes.create_unicode_buffer(tamanho + 1)
        classe = ctypes.create_unicode_buffer(256)
        user32.GetWindowTextW(hwnd, titulo, len(titulo))
        user32.GetClassNameW(hwnd, classe, len(classe))
        texto = titulo.value
        nome_classe = classe.value
        pontos = 0
        if titulo_esperado and normalizar(titulo_esperado) in normalizar(texto):
            pontos += 100
        if "apps er systems" in normalizar(texto):
            pontos += 50
        if nome_classe == "Chrome_WidgetWin_1":
            pontos += 20
        if pontos:
            candidatos.append((pontos, int(hwnd), texto, nome_classe))
        return True

    user32.EnumWindows(callback, 0)
    if not candidatos:
        raise ErCardError("Janela", "Janela visivel do ER Card nao encontrada.")
    candidatos.sort(reverse=True)
    _, hwnd, titulo, classe = candidatos[0]
    return hwnd, titulo, classe


def normalizar_janela_ercard_windows(titulo_esperado, log):
    import ctypes

    user32 = ctypes.windll.user32
    try:
        user32.SetProcessDPIAware()
    except Exception:
        pass

    hwnd, titulo, classe = localizar_janela_ercard_windows(titulo_esperado)
    antes = _estado_janela_windows(hwnd)

    todos_monitores = []
    # The window state helper exposes only touched monitors. Restore first, then
    # ask MonitorFromPoint for the primary monitor's work area.
    SW_RESTORE = 9
    SW_MAXIMIZE = 3
    MONITOR_DEFAULTTOPRIMARY = 1
    user32.ShowWindow(hwnd, SW_RESTORE)
    time.sleep(0.5)

    from ctypes import wintypes

    class POINT(ctypes.Structure):
        _fields_ = [("x", wintypes.LONG), ("y", wintypes.LONG)]

    class RECT(ctypes.Structure):
        _fields_ = [
            ("left", wintypes.LONG),
            ("top", wintypes.LONG),
            ("right", wintypes.LONG),
            ("bottom", wintypes.LONG),
        ]

    class MONITORINFOEXW(ctypes.Structure):
        _fields_ = [
            ("cbSize", wintypes.DWORD),
            ("rcMonitor", RECT),
            ("rcWork", RECT),
            ("dwFlags", wintypes.DWORD),
            ("szDevice", wintypes.WCHAR * 32),
        ]

    hmonitor = user32.MonitorFromPoint(POINT(0, 0), MONITOR_DEFAULTTOPRIMARY)
    info = MONITORINFOEXW()
    info.cbSize = ctypes.sizeof(info)
    if not user32.GetMonitorInfoW(hmonitor, ctypes.byref(info)):
        raise ErCardError("Janela", "Nao foi possivel consultar o monitor principal.")
    largura = info.rcWork.right - info.rcWork.left
    altura = info.rcWork.bottom - info.rcWork.top
    if not user32.MoveWindow(
        hwnd, info.rcWork.left, info.rcWork.top, largura, altura, True
    ):
        raise ErCardError("Janela", "MoveWindow falhou ao mover para o monitor principal.")
    time.sleep(0.5)
    user32.ShowWindow(hwnd, SW_MAXIMIZE)
    user32.SetForegroundWindow(hwnd)
    time.sleep(1)

    depois = _estado_janela_windows(hwnd)
    visivel = depois["retangulo_visivel"]
    monitor = (
        info.rcMonitor.left,
        info.rcMonitor.top,
        info.rcMonitor.right,
        info.rcMonitor.bottom,
    )
    contida = (
        visivel[0] >= monitor[0]
        and visivel[1] >= monitor[1]
        and visivel[2] <= monitor[2]
        and visivel[3] <= monitor[3]
    )
    valida = (
        depois["maximizada"]
        and contida
        and len(depois["monitores"]) == 1
        and depois["monitores"][0]["principal"]
    )
    relatorio = {
        "hwnd": hwnd,
        "titulo": titulo,
        "classe": classe,
        "antes": antes,
        "depois": depois,
        "valida": valida,
    }
    if not valida:
        raise ErCardError("Janela", f"Validacao da janela maximizada falhou: {relatorio}")
    log_ercard(log, "Janela ER Card movida e maximizada no monitor principal.")
    return relatorio


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


def _localizar_painel_de_opcoes(imagem_png):
    try:
        import cv2
        import numpy as np

        dados = np.frombuffer(imagem_png, dtype=np.uint8)
        imagem = cv2.imdecode(dados, cv2.IMREAD_COLOR)
        if imagem is None:
            return None

        mascara = cv2.inRange(
            imagem,
            np.array([225, 225, 225], dtype=np.uint8),
            np.array([255, 255, 255], dtype=np.uint8),
        )
        contornos, _ = cv2.findContours(
            mascara, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE
        )
        altura_imagem, largura_imagem = imagem.shape[:2]
        candidatos = []
        for contorno in contornos:
            x, y, largura, altura = cv2.boundingRect(contorno)
            proporcao_area = (largura * altura) / (largura_imagem * altura_imagem)
            proporcao_formato = largura / max(altura, 1)
            if (
                largura >= largura_imagem * 0.20
                and altura >= altura_imagem * 0.20
                and proporcao_area >= 0.04
                and 1.15 <= proporcao_formato <= 2.5
            ):
                candidatos.append((largura * altura, x, y, largura, altura))

        if not candidatos:
            return None

        _, x, y, largura, altura = max(candidatos)
        regiao_opcao = imagem[
            y + int(altura * 0.35):y + int(altura * 0.82),
            x + int(largura * 0.02):x + int(largura * 0.48),
        ]
        if regiao_opcao.size == 0:
            return None

        cinza = cv2.cvtColor(regiao_opcao, cv2.COLOR_BGR2GRAY)
        pixels_escuros = int((cinza < 100).sum())
        if pixels_escuros < 120:
            return None

        return x, y, largura, altura, largura_imagem, altura_imagem
    except Exception:
        return None


def _localizar_dialogo_login(imagem_png):
    try:
        import cv2
        import numpy as np

        imagem = cv2.imdecode(
            np.frombuffer(imagem_png, dtype=np.uint8), cv2.IMREAD_COLOR
        )
        if imagem is None:
            return None
        altura_imagem, largura_imagem = imagem.shape[:2]
        mascara = cv2.inRange(
            imagem,
            np.array([225, 225, 225], dtype=np.uint8),
            np.array([255, 255, 255], dtype=np.uint8),
        )
        contornos, _ = cv2.findContours(
            mascara, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE
        )
        candidatos = []
        for contorno in contornos:
            x, y, largura, altura = cv2.boundingRect(contorno)
            proporcao_area = (largura * altura) / (largura_imagem * altura_imagem)
            proporcao_formato = largura / max(altura, 1)
            if (
                largura >= largura_imagem * 0.20
                and altura >= altura_imagem * 0.35
                and proporcao_area >= 0.08
                and 0.80 <= proporcao_formato <= 1.50
            ):
                candidatos.append((largura * altura, x, y, largura, altura))
        if not candidatos:
            return None
        _, x, y, largura, altura = max(candidatos)
        return x, y, largura, altura, largura_imagem, altura_imagem
    except Exception:
        return None


def _localizar_janela_gerencial(imagem_png):
    try:
        import cv2
        import numpy as np

        imagem = cv2.imdecode(np.frombuffer(imagem_png, dtype=np.uint8), cv2.IMREAD_COLOR)
        if imagem is None:
            return None
        altura_imagem, largura_imagem = imagem.shape[:2]
        cinza = cv2.cvtColor(imagem, cv2.COLOR_BGR2GRAY)
        _, mascara = cv2.threshold(cinza, 180, 255, cv2.THRESH_BINARY)
        contornos, _ = cv2.findContours(
            mascara, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE
        )
        candidatos = []
        for contorno in contornos:
            x, y, largura, altura = cv2.boundingRect(contorno)
            proporcao = largura / max(altura, 1)
            if (
                largura_imagem * 0.25 <= largura <= largura_imagem * 0.55
                and altura_imagem * 0.30 <= altura <= altura_imagem * 0.65
                and 1.20 <= proporcao <= 2.30
            ):
                candidatos.append((largura * altura, x, y, largura, altura))
        if not candidatos:
            return None
        _, x, y, largura, altura = max(candidatos)
        return x, y, largura, altura, largura_imagem, altura_imagem
    except Exception:
        return None


def _canvas_principal(driver):
    canvases = [
        elemento
        for elemento in driver.find_elements(By.TAG_NAME, "canvas")
        if elemento.is_displayed()
    ]
    if not canvases:
        raise ErCardError("Canvas", "Canvas da sessão remota não encontrado.")
    return max(
        canvases,
        key=lambda elemento: elemento.size["width"] * elemento.size["height"],
    )


def _ativar_janela_com_canvas(driver):
    try:
        handles = list(driver.window_handles)
    except Exception:
        return False

    for handle in reversed(handles):
        try:
            driver.switch_to.window(handle)
            if any(
                elemento.is_displayed()
                for elemento in driver.find_elements(By.TAG_NAME, "canvas")
            ):
                return True
        except Exception:
            continue
    return False


def _acao_canvas(
    driver,
    canvas,
    x,
    y,
    largura_imagem,
    altura_imagem,
    texto=None,
    cliques=1,
):
    deslocamento = driver.execute_script(
        """
        const rect = arguments[0].getBoundingClientRect();
        const clientX = (arguments[1] / arguments[3]) * window.innerWidth;
        const clientY = (arguments[2] / arguments[4]) * window.innerHeight;
        return {
            x: Math.round(clientX - (rect.left + rect.width / 2)),
            y: Math.round(clientY - (rect.top + rect.height / 2)),
        };
        """,
        canvas,
        x,
        y,
        largura_imagem,
        altura_imagem,
    )
    acao = ActionChains(driver).move_to_element_with_offset(
        canvas, deslocamento["x"], deslocamento["y"]
    )
    acao = acao.double_click() if cliques == 2 else acao.click()
    if texto is not None:
        acao = acao.key_down(Keys.CONTROL).send_keys("a").key_up(Keys.CONTROL)
        acao = acao.send_keys(texto)
    acao.perform()


def _geometria_canvas(driver):
    from PIL import Image
    import io

    if not _ativar_janela_com_canvas(driver):
        raise ErCardError("Canvas", "A aba da sessão remota foi fechada.")
    imagem_png = driver.get_screenshot_as_png()
    largura, altura = Image.open(io.BytesIO(imagem_png)).size
    return _canvas_principal(driver), largura, altura


def _clicar_relativo_canvas(driver, x, y, texto=None, cliques=1):
    canvas, largura, altura = _geometria_canvas(driver)
    _acao_canvas(
        driver,
        canvas,
        int(largura * x),
        int(altura * y),
        largura,
        altura,
        texto=texto,
        cliques=cliques,
    )


def _clicar_janela_gerencial(
    driver, x_relativo, y_relativo, texto=None, timeout=20
):
    fim = time.monotonic() + timeout
    painel = None
    while time.monotonic() < fim:
        imagem_png = driver.get_screenshot_as_png()
        painel = _localizar_janela_gerencial(imagem_png)
        if painel is not None:
            break
        time.sleep(0.4)
    if painel is None:
        raise ErCardError(
            "Consulta",
            "Janela Geração de Dados Gerenciais não apareceu em 20 segundos.",
        )
    x, y, largura, altura, largura_imagem, altura_imagem = painel
    _acao_canvas(
        driver,
        _canvas_principal(driver),
        x + int(largura * x_relativo),
        y + int(altura * y_relativo),
        largura_imagem,
        altura_imagem,
        texto=texto,
    )


def _mudanca_visual_significativa(
    imagem_antes,
    imagem_depois,
    regiao=None,
    proporcao_minima=0.02,
):
    try:
        import cv2
        import numpy as np

        antes = cv2.imdecode(
            np.frombuffer(imagem_antes, dtype=np.uint8), cv2.IMREAD_GRAYSCALE
        )
        depois = cv2.imdecode(
            np.frombuffer(imagem_depois, dtype=np.uint8), cv2.IMREAD_GRAYSCALE
        )
        if antes is None or depois is None or antes.shape != depois.shape:
            return True
        if regiao is not None:
            x, y, largura, altura = regiao
            limite_x = min(antes.shape[1], x + largura)
            limite_y = min(antes.shape[0], y + altura)
            antes = antes[max(0, y):limite_y, max(0, x):limite_x]
            depois = depois[max(0, y):limite_y, max(0, x):limite_x]
            if antes.size == 0 or depois.size == 0:
                return True
        diferenca = cv2.absdiff(antes, depois)
        proporcao = float((diferenca > 20).sum()) / diferenca.size
        return proporcao >= proporcao_minima
    except Exception:
        return imagem_antes != imagem_depois


def _clicar_opcao_canvas(
    driver,
    timeout,
    etapa,
    log,
    posicao_horizontal,
    nome_opcao,
    cliques=1,
    confirmar_mudanca=False,
    detector_sucesso=None,
):
    log_ercard(log, f"Aguardando a tela de {etapa} terminar de carregar...")
    fim = time.monotonic() + timeout
    proximo_aviso = time.monotonic() + 10
    pronto_anterior = None

    while time.monotonic() < fim:
        canvases = [
            elemento
            for elemento in driver.find_elements(By.TAG_NAME, "canvas")
            if elemento.is_displayed()
        ]
        canvases.sort(
            key=lambda elemento: elemento.size["width"] * elemento.size["height"],
            reverse=True,
        )

        for canvas in canvases:
            imagem_png = driver.get_screenshot_as_png()
            painel = _localizar_painel_de_opcoes(imagem_png)
            if painel is None:
                continue

            if pronto_anterior != painel[:4]:
                pronto_anterior = painel[:4]
                time.sleep(0.8)
                break

            x, y, largura, altura, largura_imagem, altura_imagem = painel
            clique_x = x + int(largura * posicao_horizontal)
            clique_y = y + int(altura * 0.58)
            deslocamento = driver.execute_script(
                """
                const rect = arguments[0].getBoundingClientRect();
                const clientX = (arguments[1] / arguments[3]) * window.innerWidth;
                const clientY = (arguments[2] / arguments[4]) * window.innerHeight;
                return {
                    x: Math.round(clientX - (rect.left + rect.width / 2)),
                    y: Math.round(clientY - (rect.top + rect.height / 2)),
                };
                """,
                canvas,
                clique_x,
                clique_y,
                largura_imagem,
                altura_imagem,
            )
            try:
                acao = ActionChains(driver).move_to_element_with_offset(
                    canvas,
                    deslocamento["x"],
                    deslocamento["y"],
                )
                acao = acao.double_click() if cliques == 2 else acao.click()
                acao.perform()
            except Exception:
                tipos_evento = (
                    ["mousemove", "mousedown", "mouseup", "click"] * cliques
                )
                driver.execute_script(
                    """
                    const canvas = arguments[0];
                    const rect = canvas.getBoundingClientRect();
                    const clientX = rect.left + rect.width / 2 + arguments[1];
                    const clientY = rect.top + rect.height / 2 + arguments[2];
                    for (const type of arguments[3]) {
                        canvas.dispatchEvent(new MouseEvent(type, {
                            bubbles: true, cancelable: true, view: window,
                            clientX, clientY, button: 0,
                        }));
                    }
                    """,
                    canvas,
                    deslocamento["x"],
                    deslocamento["y"],
                    tipos_evento,
                )
            tipo_clique = "Duplo clique" if cliques == 2 else "Clique"
            log_ercard(log, f"{tipo_clique} enviado para {nome_opcao} no canvas.")

            if not confirmar_mudanca:
                return True

            fim_confirmacao = min(fim, time.monotonic() + 6)
            while time.monotonic() < fim_confirmacao:
                imagem_atual = driver.get_screenshot_as_png()
                if detector_sucesso is not None:
                    mudou = detector_sucesso(imagem_atual) is not None
                else:
                    mudou = _mudanca_visual_significativa(
                        imagem_png,
                        imagem_atual,
                        regiao=(x, y, largura, altura),
                        proporcao_minima=0.01,
                    )
                if mudou:
                    log_ercard(
                        log,
                        f"Abertura de {nome_opcao} confirmada por mudança da tela.",
                    )
                    return True
                time.sleep(0.4)

            log_ercard(
                log,
                f"{nome_opcao} não abriu após o clique; repetindo no mesmo ícone.",
            )
            pronto_anterior = None

        if time.monotonic() >= proximo_aviso:
            restante = max(0, int(fim - time.monotonic()))
            log_ercard(log, f"Tela ainda carregando... até {restante}s restantes.")
            proximo_aviso = time.monotonic() + 10
        time.sleep(0.5)

    if confirmar_mudanca:
        raise ErCardError(
            etapa,
            f"{nome_opcao} permaneceu na tela de opções após as tentativas de clique.",
        )
    raise ErCardError(etapa, "A tela remota não ficou pronta dentro do tempo limite.")


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
    janelas_antes = set(driver.window_handles)
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
        return janelas_antes

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
            (By.ID, "buttonLogOn"),
            (By.CSS_SELECTOR, "input[type='button'][value='Entrar']"),
            (By.XPATH, "//input[@type='button' and @value='Entrar']"),
            (By.XPATH, "//button[normalize-space(.)='Entrar']"),
            (By.CSS_SELECTOR, "button[type='submit']"),
        ],
        config.timeout_normal,
        "Login 1"
    )
    WebDriverWait(driver, config.timeout_normal).until(
        lambda navegador: navegador.execute_script(
            """
            const botao = document.getElementById('buttonLogOn');
            return typeof window.cplogon === 'function'
                && botao
                && botao.style.cursor !== 'wait';
            """
        )
    )
    driver.execute_script(
        "arguments[0].scrollIntoView({block:'center', inline:'center'});", botao
    )
    try:
        botao.click()
        log_ercard(log, "Botão Entrar clicado.")
    except Exception:
        driver.execute_script("arguments[0].click();", botao)
        log_ercard(log, "Botão Entrar acionado pelo navegador.")

    fim_confirmacao = time.monotonic() + min(config.timeout_normal, 8)
    while time.monotonic() < fim_confirmacao:
        novas_janelas = [
            janela for janela in driver.window_handles if janela not in janelas_antes
        ]
        if novas_janelas:
            driver.switch_to.window(novas_janelas[-1])
            log_ercard(log, "Nova aba da comunicação detectada e selecionada.")
            resultado = "sucesso"
            break

        if _erro_login_portal_visivel(driver):
            resultado = "erro"
            break

        senhas_visiveis = [
            elemento
            for elemento in driver.find_elements(By.CSS_SELECTOR, "input[type='password']")
            if elemento.is_displayed()
        ]
        if not senhas_visiveis or "html5.html" in driver.current_url.casefold():
            resultado = "sucesso"
            break
        time.sleep(0.4)
    else:
        resultado = "comunicacao"

    if resultado == "erro":
        raise ErCardError("Login 1", "O portal recusou a autenticação.")

    if resultado == "sucesso":
        log_ercard(log, "Login 1 realizado.")
    else:
        log_ercard(log, "Login enviado; aguardando abertura da comunicação.")

    return janelas_antes


def _erro_login_portal_visivel(driver):
    try:
        for elemento in driver.find_elements(By.ID, "span-credentials-ko"):
            if elemento.is_displayed():
                return True
        texto = normalizar(driver.find_element(By.TAG_NAME, "body").text)
        return any(
            termo in texto
            for termo in (
                "credenciais invalidas",
                "usuario bloqueado",
                "authentication data",
            )
        )
    except Exception:
        return False


def _ids_janelas_remotas():
    ids = set()
    try:
        from pywinauto import Desktop

        for janela in Desktop(backend="uia").windows():
            if janela.class_name() == "RAIL_WINDOW" and janela.is_visible():
                ids.add(janela.handle)
    except Exception:
        pass
    return ids


def _nova_janela_remota(ids_anteriores):
    try:
        from pywinauto import Desktop

        for janela in Desktop(backend="uia").windows():
            if (
                janela.class_name() == "RAIL_WINDOW"
                and janela.is_visible()
                and janela.handle not in ids_anteriores
            ):
                return janela
    except Exception:
        pass
    return None


def aguardar_conexao_ercard(
    driver,
    config,
    log,
    ids_janelas_anteriores=None,
    janelas_navegador_anteriores=None,
):
    log_ercard(log, "Conexão sendo estabelecida...")
    log_ercard(
        log,
        "Se o navegador solicitar confirmação, clique em Abrir. "
        "Aguardando o ambiente ficar pronto..."
    )
    fim = time.monotonic() + config.timeout_remoto
    proximo_aviso = time.monotonic() + 10
    ids_janelas_anteriores = set(ids_janelas_anteriores or ())
    janelas_navegador_anteriores = set(janelas_navegador_anteriores or ())

    while time.monotonic() < fim:
        janelas_atuais = list(driver.window_handles)
        novas_janelas = [
            janela
            for janela in janelas_atuais
            if janela not in janelas_navegador_anteriores
        ]
        try:
            janela_atual = driver.current_window_handle
        except Exception:
            janela_atual = None
        destino = novas_janelas[-1] if novas_janelas else (janelas_atuais[-1] if janelas_atuais else None)
        if destino is not None and janela_atual != destino:
            try:
                driver.switch_to.window(destino)
                log_ercard(log, "Nova aba da comunicação detectada e selecionada.")
            except Exception:
                time.sleep(0.5)
                continue

        if _erro_login_portal_visivel(driver):
            raise ErCardError("Login 1", "O portal recusou a autenticação.")

        janela_remota = _nova_janela_remota(ids_janelas_anteriores)
        if janela_remota is not None:
            titulo = janela_remota.window_text() or "janela sem título"
            log_ercard(log, f"Nova janela remota detectada: {titulo}")
            time.sleep(2)
            return janela_remota.handle

        try:
            texto = normalizar(driver.find_element(By.TAG_NAME, "body").text)
            conectando = "aguarde enquanto a conexao" in texto
            canvases_visiveis = any(
                canvas.is_displayed()
                for canvas in driver.find_elements(By.TAG_NAME, "canvas")
            )
            pronto = canvases_visiveis and (
                "crediagora" in texto
                or "er cartao crediagora" in texto
                or "canvas" in driver.page_source.casefold()
            )
            if pronto and not conectando:
                log_ercard(log, "Ambiente remoto carregado no navegador.")
                return None
        except Exception:
            pass

        if time.monotonic() >= proximo_aviso:
            restante = max(0, int(fim - time.monotonic()))
            log_ercard(log, f"Aguardando comunicação... até {restante}s restantes.")
            proximo_aviso = time.monotonic() + 10
        time.sleep(0.5)

    raise ErCardError(
        "Conexão",
        f"A comunicação não ficou pronta em {config.timeout_remoto} segundos."
    )


def selecionar_ambiente_crediagora(driver, config, log):
    log_ercard(log, "Selecionando ambiente CrediAgora...")
    _ativar_janela_com_canvas(driver)
    try:
        _clicar_texto_web(
            driver, "CrediAgora", min(config.timeout_normal, 3), "Ambiente CrediAgora"
        )
        return True
    except ErCardError:
        canvases = driver.find_elements(By.TAG_NAME, "canvas")
        if canvases:
            return _clicar_opcao_canvas(
                driver,
                config.timeout_remoto,
                "Ambiente CrediAgora",
                log,
                posicao_horizontal=0.14,
                nome_opcao="CrediAgora",
                confirmar_mudanca=True,
            )
        return False


def abrir_er_cartao(driver, config, log):
    log_ercard(log, "Abrindo ER Cartão CrediAgora...")
    _ativar_janela_com_canvas(driver)
    janelas_antes = set(driver.window_handles)
    try:
        _clicar_texto_web(
            driver,
            "ER Cartão CrediAgora",
            min(config.timeout_normal, 3),
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
        canvases = driver.find_elements(By.TAG_NAME, "canvas")
        if canvases:
            return _clicar_opcao_canvas(
                driver,
                config.timeout_remoto,
                "Módulo ER Cartão CrediAgora",
                log,
                posicao_horizontal=0.38,
                nome_opcao="ER Cartão CrediAgora",
                cliques=2,
                confirmar_mudanca=True,
                detector_sucesso=_localizar_dialogo_login,
            )
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


def realizar_login_sistema_canvas(driver, config, log):
    log_ercard(log, "Aguardando Login 2 no canvas...")
    fim = time.monotonic() + config.timeout_remoto
    painel = None
    imagem_png = None
    while time.monotonic() < fim:
        imagem_png = driver.get_screenshot_as_png()
        painel = _localizar_dialogo_login(imagem_png)
        if painel is not None:
            break
        time.sleep(0.5)
    if painel is None:
        raise ErCardError("Login 2", "Diálogo de login do ER Card não encontrado.")

    x, y, largura, altura, largura_imagem, altura_imagem = painel
    canvas = _canvas_principal(driver)
    log_ercard(log, "Preenchendo usuário do Login 2...")
    _acao_canvas(
        driver,
        canvas,
        x + int(largura * 0.75),
        y + int(altura * 0.77),
        largura_imagem,
        altura_imagem,
        config.sistema_usuario,
    )
    log_ercard(log, "Preenchendo senha do Login 2...")
    ActionChains(driver).send_keys(Keys.TAB).key_down(Keys.CONTROL).send_keys(
        "a"
    ).key_up(Keys.CONTROL).send_keys(config.sistema_senha).send_keys(
        Keys.ENTER
    ).perform()
    log_ercard(log, "Login 2 enviado; aguardando a aplicação principal...")

    fim = time.monotonic() + config.timeout_remoto
    while time.monotonic() < fim:
        atual = driver.get_screenshot_as_png()
        if _localizar_dialogo_login(atual) is None:
            log_ercard(log, "Login 2 concluído no canvas.")
            return
        time.sleep(0.5)
    raise ErCardError("Login 2", "O diálogo permaneceu aberto após clicar em Ok.")


def abrir_consulta_gerencial_canvas(driver, config, log):
    log_ercard(log, "Aguardando a aplicação principal estabilizar...")
    time.sleep(1.5)
    log_ercard(log, "Abrindo Consultas > Gerencial no canvas...")
    _clicar_relativo_canvas(driver, 0.110, 0.043)
    time.sleep(0.4)
    _clicar_relativo_canvas(driver, 0.130, 0.100)
    time.sleep(0.8)


def selecionar_tabela_contratos_canvas(driver, config, log):
    log_ercard(log, f"Selecionando {CONSULTA_CONTRATOS} no canvas...")
    _clicar_janela_gerencial(driver, 0.770, 0.270)
    time.sleep(0.8)

    # Pesquisa pelo nome completo para não depender da posição da linha na lista.
    _clicar_relativo_canvas(driver, 0.188, 0.047, texto=CONSULTA_CONTRATOS)
    _clicar_relativo_canvas(driver, 0.255, 0.047)
    time.sleep(0.8)
    _clicar_relativo_canvas(driver, 0.120, 0.120)
    _clicar_relativo_canvas(driver, 0.009, 0.055)

    fim = time.monotonic() + config.timeout_normal
    while time.monotonic() < fim:
        painel = _localizar_janela_gerencial(driver.get_screenshot_as_png())
        if painel is not None and painel[3] >= painel[5] * 0.40:
            break
        time.sleep(0.4)
    else:
        raise ErCardError(
            "Consulta", "A consulta 14 não carregou os campos de parâmetros."
        )
    log_ercard(log, "Tabela de Contratos selecionada.")


def configurar_exportacao_csv_canvas(driver, config, arquivo_webfile, log):
    log_ercard(log, f"Selecionando {FORMATO_CSV} no canvas...")
    _clicar_janela_gerencial(driver, 0.950, 0.307)
    time.sleep(0.8)
    _clicar_janela_gerencial(driver, 0.450, 0.365)
    time.sleep(0.8)

    log_ercard(log, "Abrindo seletor do arquivo de exportação...")
    _clicar_janela_gerencial(driver, 0.955, 0.535)
    time.sleep(0.8)

    for numero in (1, 2):
        log_ercard(log, f"Confirmando Restrictions {numero}/2 no canvas...")
        ActionChains(driver).send_keys(Keys.ENTER).perform()
        time.sleep(0.6)

    # WebFile é publicado pela sessão remota como um compartilhamento tsclient.
    # Usar o UNC evita depender do nome dinâmico exibido ("WebFile on MCT...").
    log_ercard(log, r"Abrindo This PC\WebFile via \\tsclient\WebFile...")
    ActionChains(driver).key_down(Keys.CONTROL).send_keys("l").key_up(
        Keys.CONTROL
    ).send_keys(r"\\tsclient\WebFile").send_keys(Keys.ENTER).perform()
    time.sleep(0.8)

    nome = Path(arquivo_webfile).name
    log_ercard(log, f"Informando nome do arquivo no WebFile: {nome}")
    ActionChains(driver).key_down(Keys.ALT).send_keys("n").key_up(
        Keys.ALT
    ).key_down(Keys.CONTROL).send_keys("a").key_up(Keys.CONTROL).send_keys(
        nome
    ).key_down(Keys.ALT).send_keys("o").key_up(Keys.ALT).perform()
    time.sleep(0.8)


def preencher_periodo_canvas(driver, config, log):
    data_final = datetime.now().strftime("%d%m%y")

    _clicar_janela_gerencial(driver, 0.455, 0.706)
    time.sleep(1)
    _clicar_janela_gerencial(driver, 0.455, 0.706)
    time.sleep(0.5)
    ActionChains(driver).key_down(Keys.CONTROL).send_keys("a").key_up(
        Keys.CONTROL
    ).send_keys(Keys.BACKSPACE).perform()
    time.sleep(0.3)
    for caractere in DATA_INICIAL:
        ActionChains(driver).send_keys(caractere).perform()
        time.sleep(0.12)
    time.sleep(1)
    ActionChains(driver).send_keys(Keys.TAB).perform()
    time.sleep(1.5)

    # O TAB da data inicial posiciona o foco diretamente na data final.
    time.sleep(0.5)
    for caractere in data_final:
        ActionChains(driver).send_keys(caractere).perform()
        time.sleep(0.12)
    time.sleep(0.5)
    ActionChains(driver).send_keys(Keys.TAB).perform()
    time.sleep(1.5)
    log_ercard(log, f"Data inicial: {DATA_INICIAL}")
    log_ercard(log, f"Data final: {data_final}")


def exportar_tabela_contratos_canvas(driver, config, arquivo_final, log):
    log_ercard(log, "Iniciando exportação no canvas...")
    tela_antes = driver.get_screenshot_as_png()
    _clicar_janela_gerencial(driver, 0.433, 0.925)
    fim = time.monotonic() + config.timeout_normal
    while time.monotonic() < fim:
        if _mudanca_visual_significativa(
            tela_antes, driver.get_screenshot_as_png()
        ):
            break
        time.sleep(0.2)
    else:
        raise ErCardError("Exportação", "A confirmação da exportação não apareceu.")

    # O foco inicial fica em No; Yes deve ser acionado diretamente pelo mouse.
    instante_yes = datetime.now()
    _clicar_relativo_canvas(driver, 0.502, 0.582)
    log_ercard(log, f"Yes clicado com o mouse em {instante_yes:%d/%m/%Y %H:%M:%S}.")
    log_ercard(log, "Exportando Tabela de Contratos...")
    log_ercard(
        log,
        "Aguardando processamento e download do CSV por até "
        f"{config.timeout_exportacao} segundos; Not Responding não é falha.",
    )


def mover_download_webfile(arquivo_final, config, log):
    arquivo_final = Path(arquivo_final)
    pasta_download = Path(config.pasta_download or arquivo_final.parent)
    origem = pasta_download / arquivo_final.name
    fim = time.monotonic() + config.timeout_exportacao
    ultimo_tamanho = -1
    estavel = 0

    while time.monotonic() < fim:
        if arquivo_final.exists():
            return arquivo_final
        if origem.exists():
            tamanho = origem.stat().st_size
            if tamanho > 0 and tamanho == ultimo_tamanho:
                estavel += 1
                if estavel >= 2:
                    arquivo_final.parent.mkdir(parents=True, exist_ok=True)
                    shutil.move(str(origem), str(arquivo_final))
                    log_ercard(log, f"CSV movido do WebFile para: {arquivo_final}")
                    return arquivo_final
            else:
                estavel = 0
            ultimo_tamanho = tamanho
        time.sleep(1)

    raise ErCardError(
        "Download WebFile",
        f"O CSV não apareceu em {pasta_download} dentro do tempo limite.",
    )


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
    data_final = datetime.now().strftime("%d%m%y")
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
        ids_janelas_anteriores = _ids_janelas_remotas()
        log_ercard(
            log,
            f"Janelas remotas já abertas e ignoradas: {len(ids_janelas_anteriores)}."
        )
        janelas_navegador_anteriores = realizar_login_portal(driver, config, log)
        aguardar_conexao_ercard(
            driver,
            config,
            log,
            ids_janelas_anteriores=ids_janelas_anteriores,
            janelas_navegador_anteriores=janelas_navegador_anteriores,
        )

        if any(
            canvas.is_displayed()
            for canvas in driver.find_elements(By.TAG_NAME, "canvas")
        ):
            normalizar_janela_ercard_windows(driver.title, log)

        ambiente_web = selecionar_ambiente_crediagora(driver, config, log)
        modulo_web = abrir_er_cartao(driver, config, log) if ambiente_web else False

        modo_canvas = bool(driver.find_elements(By.TAG_NAME, "canvas"))
        ui = InterfaceWindowsErCard(config, log)
        if not ambiente_web and not modo_canvas:
            ui.selecionar_texto_em_janela(r"ER Systems", "CrediAgora")
        if not modulo_web and not modo_canvas:
            ui.selecionar_texto_em_janela(r"ER Systems", "ER Cartão CrediAgora")

        if modo_canvas:
            realizar_login_sistema_canvas(driver, config, log)
            abrir_consulta_gerencial_canvas(driver, config, log)
            selecionar_tabela_contratos_canvas(driver, config, log)
            configurar_exportacao_csv_canvas(driver, config, arquivo_final, log)
            preencher_periodo_canvas(driver, config, log)
            exportar_tabela_contratos_canvas(driver, config, arquivo_final, log)
            mover_download_webfile(arquivo_final, config, log)
        else:
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
