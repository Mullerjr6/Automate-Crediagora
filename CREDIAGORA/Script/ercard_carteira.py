"""Fluxo separado de relatorios gerenciais do ERCard."""

import time
from pathlib import Path

import ercard
from ercard import executar_fase_ercard
from start import criar_configuracao_ercard, iniciar_chrome, log

PASTA_EXPORTACAO = Path.home() / "Desktop" / "ATT manh\u00e3"
CONSULTAS = [
    ("Relat\u00f3rio Gerencial - Carteira", "RELATORIO GERENCIAL - CARTEIRA", "Relatorio gerencial - carteira"),
    ("06 - Tabela Propostas - Cr\u00e9dito", "06 - TABELA PROPOSTAS - CREDITO", "Tabela propostas - credito"),
    ("14 - Tabela Contratos", "14 - TABELA CONTRATOS", "Tabela contratos"),
]

def main():
    inicio_total = time.monotonic()
    PASTA_EXPORTACAO.mkdir(parents=True, exist_ok=True)
    config = criar_configuracao_ercard(PASTA_EXPORTACAO)
    driver = None
    try:
        log("========== INICIO FLUXO ERCard - RELATORIOS ==========")
        log(f"Pasta de destino: {PASTA_EXPORTACAO}")
        for nome_relatorio, consulta, prefixo in CONSULTAS:
            inicio_relatorio = time.monotonic()
            log(f"[ERCard][INICIO] {nome_relatorio}")
            log(f"[ERCard][DOWNLOAD] Solicitação iniciada: {nome_relatorio}")
            ercard.CONSULTA_CONTRATOS = consulta
            ercard.gerar_nome_tabela_contratos = (lambda agora=None, prefixo=prefixo: f"Tabela contratos - {prefixo} {(agora or __import__('datetime').datetime.now()):%d-%m-%Y %H-%M}.csv")
            driver = iniciar_chrome(headless=False, manter_navegador=False)
            arquivo = executar_fase_ercard(driver, config, log)
            duracao = (time.monotonic() - inicio_relatorio) / 60
            log(f"[ERCard][IMPORTACAO] Arquivo salvo na pasta de destino: {arquivo}")
            log(f"[ERCard][FIM] {nome_relatorio} concluido em {duracao:.1f} minutos")
            driver.quit()
            driver = None
        total = (time.monotonic() - inicio_total) / 60
        log(f"========== FLUXO ERCard CONCLUIDO EM {total:.1f} MINUTOS ==========")
        return 0
    except Exception as erro:
        total = (time.monotonic() - inicio_total) / 60
        log(f"[ERCard][ERRO] Processo interrompido apos {total:.1f} minutos: {erro}")
        return 1
    finally:
        if driver is not None:
            try:
                driver.quit()
            except Exception:
                pass

if __name__ == "__main__":
    raise SystemExit(main())
