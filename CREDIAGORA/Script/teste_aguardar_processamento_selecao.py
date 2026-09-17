import time
from datetime import datetime, timedelta
from pathlib import Path

from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys

import ercard
import start
from teste_confirmar_yes import aguardar_mudanca, digitar_lentamente, log, salvar_png


TIMEOUT_PROCESSAMENTO = 5 * 60
INTERVALO_MONITORAMENTO = 2


def arquivo_exportado_existe(config, nome_arquivo):
    candidatos = {
        Path(config.pasta_download) / nome_arquivo,
        Path(config.pasta_exportacao) / nome_arquivo,
    }
    encontrados = [caminho for caminho in candidatos if caminho.exists()]
    return encontrados


def monitorar_processamento(driver, config, nome_arquivo, imagem_selecionando):
    inicio = time.monotonic()
    verificacoes = 0

    while True:
        decorrido = time.monotonic() - inicio
        if decorrido >= TIMEOUT_PROCESSAMENTO:
            captura = driver.get_screenshot_as_png()
            return {
                "motivo": "timeout",
                "captura": captura,
                "duracao": decorrido,
                "arquivos": arquivo_exportado_existe(config, nome_arquivo),
                "verificacoes": verificacoes,
            }

        time.sleep(min(INTERVALO_MONITORAMENTO, TIMEOUT_PROCESSAMENTO - decorrido))
        verificacoes += 1
        captura = driver.get_screenshot_as_png()
        arquivos = arquivo_exportado_existe(config, nome_arquivo)

        if arquivos:
            return {
                "motivo": "arquivo_criado",
                "captura": captura,
                "duracao": time.monotonic() - inicio,
                "arquivos": arquivos,
                "verificacoes": verificacoes,
            }

        if ercard._mudanca_visual_significativa(imagem_selecionando, captura):
            return {
                "motivo": "mudanca_visual",
                "captura": captura,
                "duracao": time.monotonic() - inicio,
                "arquivos": arquivos,
                "verificacoes": verificacoes,
            }


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
        salvar_png(
            config.pasta_erros / "fase2_processamento_antes_exportar.png",
            antes_exportar,
        )
        input("PAUSA_ANTES_EXPORTAR: valide todos os campos e pressione Enter... ")

        ercard._clicar_janela_gerencial(driver, 0.433, 0.925)
        confirmacao, demora_confirmacao = aguardar_mudanca(driver, antes_exportar)
        log(f"[EXPORTAR] Confirmacao detectada em {demora_confirmacao:.3f}s.")
        salvar_png(
            config.pasta_erros / "fase2_processamento_confirmacao.png",
            confirmacao,
        )
        input("PAUSA_YES: valide a confirmacao e pressione Enter... ")

        instante_yes = datetime.now()
        ercard._clicar_relativo_canvas(driver, 0.502, 0.582)
        log(f"[YES] Clique unico enviado em {instante_yes:%d/%m/%Y %H:%M:%S.%f}")

        selecionando, demora_selecionando = aguardar_mudanca(
            driver, confirmacao, timeout=90
        )
        instante_selecionando = instante_yes + timedelta(seconds=demora_selecionando)
        log(
            "[PROCESSAMENTO] Primeira reacao detectada em "
            f"{demora_selecionando:.3f}s, as "
            f"{instante_selecionando:%d/%m/%Y %H:%M:%S.%f}."
        )
        salvar_png(
            config.pasta_erros / "fase2_processamento_selecionando.png",
            selecionando,
        )
        input(
            "PAUSA_SELECIONANDO: confirme 'Aguarde. Selecionando dados...' "
            "e pressione Enter para observar sem interagir... "
        )

        log(
            "[PROCESSAMENTO] Monitoramento passivo iniciado: "
            "intervalo=2s; timeout=300s."
        )
        resultado = monitorar_processamento(
            driver, config, nome_arquivo, selecionando
        )
        instante_fim = datetime.now()
        salvar_png(
            config.pasta_erros / "fase2_apos_processamento_selecao.png",
            resultado["captura"],
        )
        log(f"[PROCESSAMENTO] Resultado: {resultado['motivo']}.")
        log(f"[PROCESSAMENTO] Horario final: {instante_fim:%d/%m/%Y %H:%M:%S.%f}")
        log(
            f"[PROCESSAMENTO] Duracao do estado monitorado: "
            f"{resultado['duracao']:.3f}s."
        )
        log(f"[PROCESSAMENTO] Verificacoes: {resultado['verificacoes']}.")
        if resultado["arquivos"]:
            for arquivo in resultado["arquivos"]:
                log(f"[PROCESSAMENTO] CSV encontrado: {arquivo}")
        else:
            log(f"[PROCESSAMENTO] CSV ainda nao encontrado: {nome_arquivo}")
        log("[PROCESSAMENTO] Nenhuma interacao enviada durante a espera.")
        input("PAUSA_FINAL: pressione Enter somente apos documentar o novo estado... ")
    finally:
        try:
            driver.quit()
        except Exception:
            pass


if __name__ == "__main__":
    executar()
