"""Exporta Vendas e Receita do Panorama sem alterar planilhas."""

from __future__ import annotations

import argparse
import json
import shutil
import time
from datetime import datetime
from pathlib import Path

from completo import (
    PASTA_EXPORTACOES_AUTOMACAO,
    baixar_exportacao,
    encerrar_sessao_portal,
    entrar_em_vendas_e_emprestimo,
    fazer_login,
    iniciar_chrome,
    log,
    voltar_para_emprestimos,
)

PASTA_EXPORTACAO = PASTA_EXPORTACOES_AUTOMACAO
ESTADO_FLUXO = Path(__file__).resolve().parent / "logs" / "panorama_exportacoes_carteira_estado.json"
EXPORTACOES = (
    ("vendas", "Exportação vendas"),
    ("receita", "Receita gerada"),
)


def salvar_estado(**campos):
    ESTADO_FLUXO.parent.mkdir(parents=True, exist_ok=True)
    estado = {
        "fluxo": "panorama-exportacoes-carteira",
        "atualizado_em": datetime.now().astimezone().isoformat(timespec="seconds"),
        **campos,
    }
    temporario = ESTADO_FLUXO.with_suffix(".tmp")
    temporario.write_text(json.dumps(estado, ensure_ascii=False, indent=2), encoding="utf-8")
    temporario.replace(ESTADO_FLUXO)


def nome_destino(chave, origem):
    extensao = origem.suffix.lower() or ".csv"
    return PASTA_EXPORTACAO / (
        f"{chave} panorama {datetime.now():%d-%m-%Y %H-%M}_importar{extensao}"
    )


def mover_exportacao(chave, origem):
    origem = Path(origem)
    if not origem.is_file() or origem.stat().st_size <= 0:
        raise RuntimeError(f"Download inválido para {chave}: {origem}")
    destino = nome_destino(chave, origem)
    destino.unlink(missing_ok=True)
    shutil.move(str(origem), str(destino))
    if not destino.is_file() or destino.stat().st_size <= 0:
        raise RuntimeError(f"Arquivo final inválido para {chave}: {destino}")
    return destino


def main(argv=None):
    parser = argparse.ArgumentParser(
        description="Exporta apenas Vendas e Receita do Panorama, sem alterar planilhas."
    )
    parser.add_argument(
        "--check",
        action="store_true",
        help="Valida somente a pasta de destino, sem abrir o Chrome.",
    )
    args = parser.parse_args(argv)
    if args.check:
        PASTA_EXPORTACAO.mkdir(parents=True, exist_ok=True)
        log(f"[PANORAMA][CHECK] Destino válido: {PASTA_EXPORTACAO}")
        return 0

    inicio = time.monotonic()
    driver = None
    PASTA_EXPORTACAO.mkdir(parents=True, exist_ok=True)
    salvar_estado(status="executando", etapa="inicializacao", arquivos={})
    arquivos = {}
    try:
        log("========== INÍCIO PANORAMA - VENDAS E RECEITA ==========")
        log(f"[PANORAMA][DESTINO] {PASTA_EXPORTACAO}")
        driver = iniciar_chrome(headless=True, manter_navegador=False)
        fazer_login(driver, preparar_janela=False)
        entrar_em_vendas_e_emprestimo(driver)

        for indice, (chave, layout) in enumerate(EXPORTACOES, start=1):
            salvar_estado(status="executando", etapa=chave, indice=indice, arquivos=arquivos)
            log(f"[PANORAMA][INÍCIO] {layout}")
            origem = baixar_exportacao(driver, layout)
            destino = mover_exportacao(chave, origem)
            arquivos[chave] = str(destino)
            log(f"[PANORAMA][EXPORTAÇÃO] {layout}: {destino}")
            if chave == "vendas":
                voltar_para_emprestimos(driver)

        duracao = (time.monotonic() - inicio) / 60
        salvar_estado(status="concluido", etapa="finalizado", arquivos=arquivos, duracao_minutos=round(duracao, 2))
        log(f"========== FIM PANORAMA em {duracao:.1f} minutos ==========")
        return 0
    except Exception as erro:  # noqa: BLE001
        duracao = (time.monotonic() - inicio) / 60
        salvar_estado(status="erro", etapa="erro", arquivos=arquivos, erro=str(erro), duracao_minutos=round(duracao, 2))
        log(f"[PANORAMA][ERRO] Exportação interrompida: {erro}")
        return 1
    finally:
        if driver is not None:
            try:
                encerrar_sessao_portal(driver)
            except Exception as erro:  # noqa: BLE001
                log(f"[PANORAMA][AVISO] Não foi possível encerrar a sessão: {erro}")
            try:
                driver.quit()
            except Exception as erro:  # noqa: BLE001
                log(f"[PANORAMA][AVISO] Não foi possível fechar o Chrome: {erro}")


if __name__ == "__main__":
    raise SystemExit(main())
