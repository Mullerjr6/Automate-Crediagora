import time

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

        driver.maximize_window()
        time.sleep(1)
        estado = driver.execute_script(
            "return {outerWidth, outerHeight, "
            "availWidth: screen.availWidth, availHeight: screen.availHeight};"
        )
        log(
            "[MAXIMIZADA] "
            f"janela={estado['outerWidth']}x{estado['outerHeight']}; "
            f"tela_disponivel={estado['availWidth']}x{estado['availHeight']}"
        )

        ercard.abrir_consulta_gerencial_canvas(driver, config, log)
        ercard.selecionar_tabela_contratos_canvas(driver, config, log)
        ercard._clicar_janela_gerencial(driver, 0.950, 0.307)
        time.sleep(0.8)
        ercard._clicar_janela_gerencial(driver, 0.450, 0.365)
        time.sleep(0.8)
        log("[DATA INICIAL] ARQUIVO CSV (EXCEL) selecionado por mouse.")

        ercard._clicar_janela_gerencial(driver, 0.955, 0.535)
        time.sleep(0.8)
        for numero in (1, 2):
            ActionChains(driver).send_keys(Keys.ENTER).perform()
            time.sleep(0.6)
            log(f"[DATA INICIAL] Restrictions {numero}/2 confirmado.")
        ActionChains(driver).key_down(Keys.ALT).send_keys("n").key_up(
            Keys.ALT
        ).key_down(Keys.CONTROL).send_keys("a").key_up(Keys.CONTROL).send_keys(
            "Tabela contratos teste"
        ).key_down(Keys.ALT).send_keys("o").key_up(Keys.ALT).perform()
        time.sleep(1.2)
        log(r"[DATA INICIAL] Destino: \\tsclient\WebFile\Tabela contratos teste.csv")

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
        log("[DATA INICIAL] TAB enviado uma única vez; nenhum texto adicional digitado.")

        caminho = config.pasta_erros / "fase2_data_inicial_tab.png"
        driver.save_screenshot(str(caminho))
        log(f"CAPTURA={caminho}")
        log("[DATA INICIAL] Teste encerrado sem preencher DATA FINAL e sem exportar.")
    finally:
        try:
            driver.quit()
        except Exception:
            pass


if __name__ == "__main__":
    executar()
