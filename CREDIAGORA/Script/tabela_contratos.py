"""Fluxo autônomo do ERCard para o relatório 14 - Tabela Contratos."""

import json
import time
from datetime import datetime
from pathlib import Path

import ercard
from completo import (
    criar_configuracao_ercard,
    iniciar_chrome,
    limpar_pagina_para_reinicio,
    log,
)
from ercard import executar_fase_ercard

PASTA_EXPORTACAO = (
    Path.home()
    / "TJI PROMOTORA DE VENDAS EIRELI"
    / "Crediagora-doc - dados"
    / "Automacoes"
    / "Automatização de processos Crediagora"
    / "Logs de exportações automatizadas"
)
ESTADO_FLUXO = Path(__file__).resolve().parent / "logs" / "ercard_tabela_contratos_estado.json"
NOME_RELATORIO = "14 - Tabela Contratos"
CONSULTA = "14 - TABELA CONTRATOS"
PREFIXO_ARQUIVO = "14 - tabela contratos"


def salvar_estado(**campos):
    """Registra o resultado de forma atômica para diagnóstico e retomada."""
    ESTADO_FLUXO.parent.mkdir(parents=True, exist_ok=True)
    dados = {
        "relatorio": NOME_RELATORIO,
        "atualizado_em": datetime.now().astimezone().isoformat(timespec="seconds"),
        **campos,
    }
    temporario = ESTADO_FLUXO.with_suffix(".tmp")
    temporario.write_text(json.dumps(dados, ensure_ascii=False, indent=2), encoding="utf-8")
    temporario.replace(ESTADO_FLUXO)


def main():
    inicio = time.monotonic()
    driver = None
    PASTA_EXPORTACAO.mkdir(parents=True, exist_ok=True)
    salvar_estado(status="executando", inicio=datetime.now().astimezone().isoformat(timespec="seconds"))
    try:
        log(f"========== INÍCIO ERCard - {NOME_RELATORIO} ==========")
        log(f"[ERCard][INÍCIO] {NOME_RELATORIO}")
        log(f"[ERCard][DESTINO] {PASTA_EXPORTACAO}")
        ercard.CONSULTA_CONTRATOS = CONSULTA
        ercard.gerar_nome_tabela_contratos = (
            lambda agora=None: f"{PREFIXO_ARQUIVO} {(agora or datetime.now()):%d-%m-%Y %H-%M}_importar.csv"
        )
        config = criar_configuracao_ercard(PASTA_EXPORTACAO)
        driver = iniciar_chrome(headless=False, manter_navegador=False)
        arquivo = executar_fase_ercard(driver, config, log)
        if not arquivo.is_file() or arquivo.stat().st_size <= 0:
            raise RuntimeError(f"Arquivo final inválido: {arquivo}")
        duracao = (time.monotonic() - inicio) / 60
        salvar_estado(
            status="concluido",
            arquivo=str(arquivo),
            duracao_minutos=round(duracao, 2),
        )
        log(f"[ERCard][FIM] {NOME_RELATORIO} concluído em {duracao:.1f} minutos")
        return 0
    except Exception as erro:  # noqa: BLE001
        duracao = (time.monotonic() - inicio) / 60
        limpar_pagina_para_reinicio(driver, "ercard-tabela-contratos")
        salvar_estado(status="erro", erro=str(erro), duracao_minutos=round(duracao, 2))
        log(f"[ERCard][ERRO] {NOME_RELATORIO} interrompido após {duracao:.1f} minutos: {erro}")
        return 1
    finally:
        if driver is not None:
            try:
                driver.quit()
            except Exception as erro_fechar:  # noqa: BLE001
                log(f"[ERCard][AVISO] N?o foi poss?vel fechar o Chrome: {erro_fechar}")


if __name__ == "__main__":
    raise SystemExit(main())
