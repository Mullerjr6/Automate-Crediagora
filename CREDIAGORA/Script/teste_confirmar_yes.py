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


def salvar_png(caminho, conteudo):
    caminho.write_bytes(conteudo)
    log(f"CAPTURA={caminho}")


def aguardar_mudanca(driver, imagem_antes, timeout=30):
    inicio = time.monotonic()
    while time.monotonic() - inicio < timeout:
        atual = driver.get_screenshot_as_png()
        if ercard._mudanca_visual_significativa(imagem_antes, atual):
            return atual, time.monotonic() - inicio
        time.sleep(0.1)
    return driver.get_screenshot_as_png(), time.monotonic() - inicio


def executar():
    driver = start.iniciar_chrome()
    config = start.criar_configuracao_ercard()
    config.validar()
    nome_arquivo = ercard.gerar_nome_tabela_contratos()
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
            f"IsZoomed={janela['depois']['maximizada']}; "
            f"monitor={janela['depois']['monitores'][0]['dispositivo']}; "
            f"retangulo={janela['depois']['retangulo_visivel']}"
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
        ercard._clicar_janela_gerencial(driver, 0.955, 0.535)
        time.sleep(0.8)
        for numero in (1, 2):
            ActionChains(driver).send_keys(Keys.ENTER).perform()
            time.sleep(0.6)
            log(f"[FASE2] Restrictions {numero}/2 confirmado.")
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

        antes_exportar = driver.get_screenshot_as_png()
        salvar_png(config.pasta_erros / "fase2_yes_antes_exportar.png", antes_exportar)
        input("PAUSA_ANTES_EXPORTAR: continue somente se todos os campos estiverem corretos... ")

        ercard._clicar_janela_gerencial(driver, 0.433, 0.925)
        confirmacao, demora_confirmacao = aguardar_mudanca(driver, antes_exportar)
        log(f"[EXPORTAR] Confirmacao detectada em {demora_confirmacao:.3f}s.")
        salvar_png(config.pasta_erros / "fase2_confirmacao_antes_yes.png", confirmacao)
        input("PAUSA_YES: continue somente se o dialogo e o botao Yes estiverem corretos... ")

        instante_yes = datetime.now()
        ercard._clicar_relativo_canvas(driver, 0.502, 0.582)
        log(f"[YES] Clique unico enviado em {instante_yes:%d/%m/%Y %H:%M:%S.%f}")
        primeira_reacao, demora = aguardar_mudanca(driver, confirmacao)
        log(f"[YES] Primeira reacao visual detectada em {demora:.3f}s.")
        salvar_png(
            config.pasta_erros / "fase2_pos_confirmar_yes.png", primeira_reacao
        )
        log("[YES] Interacoes suspensas; nenhuma tecla ou clique adicional enviado.")
        log(f"[YES] Nome esperado: {nome_arquivo}")
        input("PAUSA_POS_YES: pressione Enter somente apos documentar a reacao... ")
    finally:
        try:
            driver.quit()
        except Exception:
            pass


if __name__ == "__main__":
    executar()
