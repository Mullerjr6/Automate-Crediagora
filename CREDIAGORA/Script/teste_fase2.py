import time
from datetime import datetime
from pathlib import Path

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
        ercard.selecionar_ambiente_crediagora(driver, config, log)
        ercard.abrir_er_cartao(driver, config, log)
        ercard.realizar_login_sistema_canvas(driver, config, log)
        ercard.abrir_consulta_gerencial_canvas(driver, config, log)
        ercard.selecionar_tabela_contratos_canvas(driver, config, log)
        log("[FASE2] Consulta 14 selecionada e tela gerencial expandida.")

        ercard._clicar_janela_gerencial(driver, 0.950, 0.307)
        time.sleep(0.8)
        ercard._clicar_janela_gerencial(driver, 0.450, 0.365)
        time.sleep(1)
        capturar(driver, config, "fase2_01_validar_csv.png")
        input("PAUSA_CSV: pressione Enter somente ap\u00f3s validar a captura... ")

        log("[FASE2] Abrindo exclusivamente o seletor de Arquivo para Exporta\u00e7\u00e3o.")
        ercard._clicar_janela_gerencial(driver, 0.955, 0.535)
        time.sleep(0.8)
        for numero in (1, 2):
            log(f"[FASE2] Confirmando Restrictions {numero}/2.")
            ActionChains(driver).send_keys(Keys.ENTER).perform()
            time.sleep(0.6)

        log(r"[FASE2] Abrindo \\tsclient\WebFile.")
        ActionChains(driver).key_down(Keys.CONTROL).send_keys("l").key_up(
            Keys.CONTROL
        ).send_keys(r"\\tsclient\WebFile").send_keys(Keys.ENTER).perform()
        time.sleep(1)
        ActionChains(driver).key_down(Keys.ALT).send_keys("n").key_up(
            Keys.ALT
        ).key_down(Keys.CONTROL).send_keys("a").key_up(Keys.CONTROL).send_keys(
            "Tabela contratos teste"
        ).key_down(Keys.ALT).send_keys("o").key_up(Keys.ALT).perform()
        time.sleep(1)
        log("[FASE2] File name preenchido e Open acionado.")

        data_final = datetime.now().strftime("%d%m%y")
        ercard._clicar_janela_gerencial(driver, 0.455, 0.706, texto="010124")
        ercard._clicar_janela_gerencial(driver, 0.455, 0.780, texto=data_final)
        log(f"[FASE2] Data inicial: 010124; data final: {data_final}.")
        capturar(driver, config, "fase2_02_antes_exportar.png")
        input("PAUSA_FINAL: pressione Enter somente ap\u00f3s validar todos os campos... ")

        ercard._clicar_janela_gerencial(driver, 0.433, 0.925)
        time.sleep(0.5)
        ActionChains(driver).send_keys(Keys.LEFT).send_keys(Keys.ENTER).perform()
        time.sleep(3)
        capturar(driver, config, "fase2_03_apos_exportar.png")
        log("[FASE2] Exportar acionado; captura p\u00f3s-exporta\u00e7\u00e3o salva.")
        input("PAUSA_POS_EXPORTAR: pressione Enter ap\u00f3s verificar a captura... ")

        destino = config.pasta_exportacao / "Tabela contratos teste.csv"
        ercard.mover_download_webfile(destino, config, log)
        ercard.validar_csv_exportado(
            destino, config.pasta_exportacao, config.timeout_exportacao, log
        )
    finally:
        driver.quit()


if __name__ == "__main__":
    executar()
