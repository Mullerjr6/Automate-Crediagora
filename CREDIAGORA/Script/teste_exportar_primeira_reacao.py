import time
from datetime import datetime
from pathlib import Path

from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys

import ercard
import start


def log(mensagem):
    print(mensagem, flush=True)


def digitar_lentamente(driver, texto, intervalo=0.12):
    for caractere in texto:
        ActionChains(driver).send_keys(caractere).perform()
        time.sleep(intervalo)


def capturar(driver, config, nome):
    caminho = config.pasta_erros / nome
    driver.save_screenshot(str(caminho))
    log(f"CAPTURA={caminho}")
    return caminho


def executar():
    driver = start.iniciar_chrome()
    config = start.criar_configuracao_ercard()
    config.validar()
    try:
        ercard.abrir_ercard(driver, config, log)
        ids = ercard._ids_janelas_remotas()
        abas = ercard.realizar_login_portal(driver, config, log)
        ercard.aguardar_conexao_ercard(
            driver,
            config,
            log,
            ids_janelas_anteriores=ids,
            janelas_navegador_anteriores=abas,
        )

        janela = ercard.normalizar_janela_ercard_windows(driver.title, log)
        log(
            "[JANELA] "
            f"hwnd={janela['hwnd']}; IsZoomed={janela['depois']['maximizada']}; "
            f"monitor={janela['depois']['monitores'][0]['dispositivo']}; "
            f"retangulo_visivel={janela['depois']['retangulo_visivel']}"
        )

        ercard.selecionar_ambiente_crediagora(driver, config, log)
        ercard.abrir_er_cartao(driver, config, log)
        ercard.realizar_login_sistema_canvas(driver, config, log)
        ercard.abrir_consulta_gerencial_canvas(driver, config, log)
        ercard.selecionar_tabela_contratos_canvas(driver, config, log)

        ercard._clicar_janela_gerencial(driver, 0.950, 0.307)
        time.sleep(0.8)
        ercard._clicar_janela_gerencial(driver, 0.450, 0.365)
        time.sleep(0.8)
        log("[FASE2] ARQUIVO CSV (EXCEL) selecionado por mouse.")

        ercard._clicar_janela_gerencial(driver, 0.955, 0.535)
        time.sleep(0.8)
        for numero in (1, 2):
            ActionChains(driver).send_keys(Keys.ENTER).perform()
            time.sleep(0.6)
            log(f"[FASE2] Restrictions {numero}/2 confirmado.")

        nome_arquivo = ercard.gerar_nome_tabela_contratos()
        ActionChains(driver).key_down(Keys.ALT).send_keys("n").key_up(
            Keys.ALT
        ).key_down(Keys.CONTROL).send_keys("a").key_up(Keys.CONTROL).send_keys(
            nome_arquivo
        ).key_down(Keys.ALT).send_keys("o").key_up(Keys.ALT).perform()
        time.sleep(1.2)
        log(f"[FASE2] Destino: \\\\tsclient\\WebFile\\{nome_arquivo}")

        ercard._clicar_janela_gerencial(driver, 0.455, 0.706)
        time.sleep(1)
        ercard._clicar_janela_gerencial(driver, 0.455, 0.706)
        time.sleep(0.5)
        ActionChains(driver).key_down(Keys.CONTROL).send_keys("a").key_up(
            Keys.CONTROL
        ).send_keys(Keys.BACKSPACE).perform()
        time.sleep(0.3)
        digitar_lentamente(driver, "010124")
        time.sleep(1)
        ActionChains(driver).send_keys(Keys.TAB).perform()
        time.sleep(1.5)

        data_final = datetime.now().strftime("%d%m%y")
        time.sleep(0.5)
        digitar_lentamente(driver, data_final)
        time.sleep(0.5)
        ActionChains(driver).send_keys(Keys.TAB).perform()
        time.sleep(1.5)
        log(f"[FASE2] Datas digitadas: 010124 e {data_final}.")

        capturar(driver, config, "fase2_antes_exportar.png")
        input("PAUSA_ANTES_EXPORTAR: continue somente se todas as validacoes estiverem corretas... ")

        imagem_antes = driver.get_screenshot_as_png()
        instante_clique = datetime.now()
        ercard._clicar_janela_gerencial(driver, 0.433, 0.925)
        log(f"[EXPORTAR] Clique unico enviado em {instante_clique:%d/%m/%Y %H:%M:%S.%f}")

        inicio = time.monotonic()
        imagem_reacao = None
        while time.monotonic() - inicio < 30:
            atual = driver.get_screenshot_as_png()
            if ercard._mudanca_visual_significativa(imagem_antes, atual):
                imagem_reacao = atual
                break
            time.sleep(0.1)

        demora = time.monotonic() - inicio
        if imagem_reacao is None:
            imagem_reacao = driver.get_screenshot_as_png()
            log(f"[EXPORTAR] Nenhuma mudanca visual significativa em {demora:.3f}s.")
        else:
            log(f"[EXPORTAR] Primeira reacao visual detectada em {demora:.3f}s.")

        caminho = config.pasta_erros / "fase2_pos_exportar_primeira_reacao.png"
        caminho.write_bytes(imagem_reacao)
        log(f"CAPTURA={caminho}")
        log("[EXPORTAR] Interacoes suspensas; nenhuma tecla ou clique adicional enviado.")
        input("PAUSA_REACAO: pressione Enter somente apos documentar a primeira reacao... ")
    finally:
        try:
            driver.quit()
        except Exception:
            pass


if __name__ == "__main__":
    executar()
