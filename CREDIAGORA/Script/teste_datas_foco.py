import time
from datetime import datetime

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


def adquirir_foco_e_preencher(driver, y_relativo, valor, nome_campo):
    ercard._clicar_janela_gerencial(driver, 0.455, y_relativo)
    time.sleep(1)
    ercard._clicar_janela_gerencial(driver, 0.455, y_relativo)
    time.sleep(0.5)
    ActionChains(driver).key_down(Keys.CONTROL).send_keys("a").key_up(
        Keys.CONTROL
    ).perform()
    time.sleep(0.3)
    ActionChains(driver).send_keys(Keys.BACKSPACE).perform()
    time.sleep(0.3)
    digitar_lentamente(driver, valor)
    time.sleep(1)
    ercard._clicar_janela_gerencial(driver, 0.700, 0.850)
    time.sleep(1)
    log(f"[DATAS] {nome_campo} recebeu {valor} com aquisi\u00e7\u00e3o expl\u00edcita de foco.")


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
        log("[DATAS] ARQUIVO CSV (EXCEL) selecionado por mouse.")

        ercard._clicar_janela_gerencial(driver, 0.955, 0.535)
        time.sleep(0.8)
        for numero in (1, 2):
            ActionChains(driver).send_keys(Keys.ENTER).perform()
            time.sleep(0.6)
            log(f"[DATAS] Restrictions {numero}/2 confirmado.")
        ActionChains(driver).key_down(Keys.ALT).send_keys("n").key_up(
            Keys.ALT
        ).key_down(Keys.CONTROL).send_keys("a").key_up(Keys.CONTROL).send_keys(
            "Tabela contratos teste"
        ).key_down(Keys.ALT).send_keys("o").key_up(Keys.ALT).perform()
        time.sleep(1.2)
        log(r"[DATAS] Destino: \\tsclient\WebFile\Tabela contratos teste.csv")

        adquirir_foco_e_preencher(driver, 0.706, "010124", "DATA INICIAL")
        capturar(driver, config, "fase2_data_inicial_foco.png")
        input("PAUSA_INICIAL: continue somente se exibir 01/01/2024... ")

        data_final = datetime.now().strftime("%d%m%y")
        adquirir_foco_e_preencher(driver, 0.780, data_final, "DATA FINAL")
        capturar(driver, config, "fase2_data_final_foco.png")
        log("[DATAS] Teste encerrado sem clicar em Exportar.")
    finally:
        try:
            driver.quit()
        except Exception:
            pass


if __name__ == "__main__":
    executar()
