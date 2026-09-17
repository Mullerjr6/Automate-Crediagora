import time

from PIL import ImageGrab

import ercard
import start


def log(mensagem):
    print(mensagem, flush=True)


def descrever_monitores(monitores):
    if not monitores:
        return "nenhum"
    return ", ".join(
        f"{item['dispositivo']}"
        f"{' (principal)' if item['principal'] else ''}"
        f" {item['retangulo']}"
        for item in monitores
    )


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

        titulo_selenium = driver.title
        relatorio = ercard.normalizar_janela_ercard_windows(titulo_selenium, log)
        antes = relatorio["antes"]
        depois = relatorio["depois"]

        log(f"[JANELA] titulo: {relatorio['titulo']}")
        log(f"[JANELA] classe: {relatorio['classe']}")
        log(f"[JANELA] hwnd: {relatorio['hwnd']}")
        log(f"[JANELA] posicao/tamanho antes: {antes['retangulo']}")
        log(f"[JANELA] retangulo visivel antes: {antes['retangulo_visivel']}")
        log(f"[JANELA] monitor(es) antes: {descrever_monitores(antes['monitores'])}")
        log(f"[JANELA] IsZoomed antes: {antes['maximizada']}")
        log(f"[JANELA] posicao/tamanho depois: {depois['retangulo']}")
        log(f"[JANELA] retangulo visivel depois: {depois['retangulo_visivel']}")
        log(f"[JANELA] monitor depois: {descrever_monitores(depois['monitores'])}")
        log(f"[JANELA] IsZoomed depois: {depois['maximizada']}")
        log(f"[JANELA] validacao final: {relatorio['valida']}")

        monitor = depois["monitores"][0]["retangulo"]
        captura = ImageGrab.grab(bbox=monitor, all_screens=True)
        caminho = config.pasta_erros / "teste_janela_monitor_principal.png"
        captura.save(caminho)
        log(f"CAPTURA={caminho}")
        time.sleep(0.5)
        log("[JANELA] Teste encerrado antes da Fase 2.")
    finally:
        try:
            driver.quit()
        except Exception:
            pass


if __name__ == "__main__":
    executar()
