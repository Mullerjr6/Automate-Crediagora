import time

from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys

import ercard
import start


def log(mensagem):
    print(mensagem, flush=True)


def capturar(driver, config, nome):
    caminho = config.pasta_erros / nome
    driver.save_screenshot(str(caminho))
    log(f"CAPTURA={caminho}")


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
        ercard.selecionar_ambiente_crediagora(driver, config, log)
        ercard.abrir_er_cartao(driver, config, log)
        ercard.realizar_login_sistema_canvas(driver, config, log)
        ercard.abrir_consulta_gerencial_canvas(driver, config, log)
        ercard.selecionar_tabela_contratos_canvas(driver, config, log)

        ercard._clicar_janela_gerencial(driver, 0.950, 0.307)
        time.sleep(0.8)
        ercard._clicar_janela_gerencial(driver, 0.450, 0.365)
        time.sleep(0.8)
        log("[WEBFILE] ARQUIVO CSV (EXCEL) selecionado por mouse.")

        ercard._clicar_janela_gerencial(driver, 0.955, 0.535)
        time.sleep(0.8)
        for numero in (1, 2):
            log(f"[WEBFILE] Confirmando Restrictions {numero}/2.")
            ActionChains(driver).send_keys(Keys.ENTER).perform()
            time.sleep(0.6)

        capturar(driver, config, "webfile_01_janela_open.png")
        log("[WEBFILE] Janela Open aberta. Aguardando valida\u00e7\u00e3o visual.")
        input("PAUSA_OPEN: pressione Enter somente se WebFile e tipo CSV estiverem corretos... ")

        ActionChains(driver).key_down(Keys.ALT).send_keys("n").key_up(
            Keys.ALT
        ).key_down(Keys.CONTROL).send_keys("a").key_up(Keys.CONTROL).send_keys(
            "Tabela contratos teste"
        ).perform()
        log("[WEBFILE] File name preenchido: Tabela contratos teste")
        ActionChains(driver).key_down(Keys.ALT).send_keys("o").key_up(Keys.ALT).perform()
        time.sleep(1.5)
        log("[WEBFILE] Open acionado; aguardando retorno para a tela gerencial.")
        capturar(driver, config, "webfile_02_retorno_gerencial.png")
        log("[WEBFILE] Teste encerrado sem preencher datas e sem exportar.")
    finally:
        try:
            driver.quit()
        except Exception:
            pass


if __name__ == "__main__":
    executar()
