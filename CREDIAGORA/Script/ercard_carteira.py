"""Fluxo separado de exportacoes gerenciais do ERCard.

Executa, na mesma sessao autenticada, Carteira, Tabela Propostas - Credito
e Tabela Contratos, salvando os CSVs em Desktop\\ATT manha.
"""

from pathlib import Path

import ercard
from ercard import executar_fase_ercard
from start import (
    PASTA_AUTOMACAO,
    criar_configuracao_ercard,
    iniciar_chrome,
    log,
)


PASTA_EXPORTACAO = Path.home() / "Desktop" / "ATT manh\u00e3"
CONSULTAS = [
    ("RELATORIO GERENCIAL - CARTEIRA", "Relatorio gerencial - carteira", "relatorio gerencial - carteira"),
    ("06 - TABELA PROPOSTAS - CREDITO", "Tabela propostas - credito", "tabela propostas - credito"),
    ("14 - TABELA CONTRATOS", "Tabela contratos", "tabela contratos"),
]


def main():
    PASTA_EXPORTACAO.mkdir(parents=True, exist_ok=True)
    config = criar_configuracao_ercard(PASTA_EXPORTACAO)
    driver = None
    try:
        log("========== INICIO FLUXO ERCARD CARTEIRA ==========")
        log(f"Pasta de exportacao: {PASTA_EXPORTACAO}")
        for consulta, prefixo, _ in CONSULTAS:
            # O ercard.py permanece no fluxo original (consulta 14). Este
            # processo separado altera apenas a consulta desta execução.
            ercard.CONSULTA_CONTRATOS = consulta
            ercard.gerar_nome_tabela_contratos = (
                lambda agora=None, prefixo=prefixo: f"Tabela contratos - {prefixo} "
                f"{(agora or __import__('datetime').datetime.now()):%d-%m-%Y %H-%M}.csv"
            )
            driver = iniciar_chrome(headless=False, manter_navegador=False)
            arquivo = executar_fase_ercard(driver, config, log)
            log(f"[ERCARD CARTEIRA] {consulta} concluído: {arquivo}")
            driver.quit()
            driver = None
        log("========== FLUXO ERCARD CARTEIRA CONCLUIDO ==========")
        return 0
    except Exception as erro:
        log(f"[ERCARD CARTEIRA][ERRO] {erro}")
        return 1
    finally:
        if driver is not None:
            try:
                driver.quit()
            except Exception:
                pass


if __name__ == "__main__":
    raise SystemExit(main())
