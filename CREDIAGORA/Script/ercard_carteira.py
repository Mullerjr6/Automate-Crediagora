"""Fluxo separado de relatorios gerenciais do ERCard."""
# ruff: noqa: BLE001, S110

import argparse
import json
import time
from datetime import datetime
from pathlib import Path

import ercard
from ercard import executar_fase_ercard
from start import (
    criar_configuracao_ercard,
    iniciar_chrome,
    limpar_pagina_para_reinicio,
    log,
)

PASTA_EXPORTACAO = Path.home() / "Desktop" / "ATT manh\u00e3"
ESTADO_FLUXO = Path(__file__).resolve().parent / "logs" / "ercard_carteira_estado.json"
VERSAO_FLUXO = 2
CONSULTAS = [
    ("14 - Tabela Contratos", "14 - TABELA CONTRATOS", "14 - tabela contratos"),
    ("06 - Propostas", "06 - TABELA PROPOSTAS - CREDITO", "06 - propostas"),
    ("Carteira Mifa", "RELATORIO GERENCIAL - CARTEIRA", "Carteira mifa"),
]


def carregar_estado():
    """Retorna o estado anterior sem deixar um JSON incompleto bloquear a retomada."""
    try:
        dados = json.loads(ESTADO_FLUXO.read_text(encoding="utf-8"))
        return dados if isinstance(dados, dict) else {}
    except (FileNotFoundError, OSError, json.JSONDecodeError):
        return {}


def salvar_estado(dados):
    """Grava o estado de forma atômica para preservar a etapa em caso de interrupção."""
    ESTADO_FLUXO.parent.mkdir(parents=True, exist_ok=True)
    temporario = ESTADO_FLUXO.with_suffix(".tmp")
    temporario.write_text(json.dumps(dados, ensure_ascii=False, indent=2), encoding="utf-8")
    temporario.replace(ESTADO_FLUXO)


def atualizar_estado(estado, **campos):
    estado.update(campos)
    estado["atualizado_em"] = datetime.now().astimezone().isoformat(timespec="seconds")
    salvar_estado(estado)


def recuperar_relatorios_ja_exportados(estado):
    """Migra um estado antigo encontrando somente arquivos gerados na data atual."""
    if estado.get("relatorios"):
        return
    data_hoje = datetime.now().strftime("%d-%m-%Y")
    relatorios = {}
    for indice, (nome_relatorio, _, prefixo) in enumerate(CONSULTAS, start=1):
        inicio_nome = f"{prefixo} {data_hoje}".casefold()
        candidatos = [
            arquivo
            for arquivo in PASTA_EXPORTACAO.glob("*.csv")
            if arquivo.name.casefold().startswith(inicio_nome) and arquivo.stat().st_size > 0
        ]
        if candidatos:
            arquivo = max(candidatos, key=lambda item: item.stat().st_mtime)
            relatorios[str(indice)] = {
                "nome": nome_relatorio,
                "arquivo": str(arquivo),
                "recuperado_de_estado_anterior": True,
            }
    estado["relatorios"] = relatorios


def limpar_pasta_exportacao():
    """Remove somente CSVs e temporarios criados pelo fluxo ERCard."""
    PASTA_EXPORTACAO.mkdir(parents=True, exist_ok=True)
    removidos = 0
    for arquivo in PASTA_EXPORTACAO.iterdir():
        if arquivo.is_file() and arquivo.suffix.casefold() in {".csv", ".tmp", ".part"}:
            arquivo.unlink()
            removidos += 1
            log(f"[ERCard][LIMPEZA] Removido: {arquivo.name}")
    log(f"[ERCard][LIMPEZA] Concluída. Arquivos removidos: {removidos}")
    return removidos

def main(argv=None):
    parser = argparse.ArgumentParser()
    parser.add_argument("--a-partir", type=int, default=1, choices=(1, 2, 3))
    parser.add_argument(
        "--limpar-pasta",
        action="store_true",
        help="Limpa somente os CSVs exportados da pasta ATT manhã.",
    )
    args = parser.parse_args(argv)
    if args.limpar_pasta:
        limpar_pasta_exportacao()
        return 0

    inicio_total = time.monotonic()
    PASTA_EXPORTACAO.mkdir(parents=True, exist_ok=True)
    config = criar_configuracao_ercard(PASTA_EXPORTACAO)
    driver = None
    etapa_atual = args.a_partir
    estado = carregar_estado() if args.a_partir > 1 else {}
    if not estado or args.a_partir == 1 or estado.get("versao_fluxo") != VERSAO_FLUXO:
        estado = {
            "versao_fluxo": VERSAO_FLUXO,
            "inicio": datetime.now().astimezone().isoformat(timespec="seconds"),
            "relatorios": {},
        }
    estado.setdefault("relatorios", {})
    if args.a_partir > 1:
        recuperar_relatorios_ja_exportados(estado)
    try:
        log("========== INICIO FLUXO ERCard - RELATORIOS ==========")
        log(f"Pasta de destino: {PASTA_EXPORTACAO}")
        for indice, (nome_relatorio, consulta, prefixo) in enumerate(CONSULTAS, start=1):
            if indice < args.a_partir:
                continue
            etapa_atual = indice
            atualizar_estado(
                estado,
                etapa=indice,
                nome=nome_relatorio,
                status="executando",
            )
            inicio_relatorio = time.monotonic()
            log(f"[ERCard][INICIO] {nome_relatorio}")
            log(f"[ERCard][DOWNLOAD] Solicitação iniciada: {nome_relatorio}")
            ercard.CONSULTA_CONTRATOS = consulta
            ercard.gerar_nome_tabela_contratos = (
                lambda agora=None, prefixo=prefixo: f"{prefixo} {(agora or datetime.now()):%d-%m-%Y %H-%M}.csv"
            )
            driver = iniciar_chrome(headless=False, manter_navegador=False)
            arquivo = executar_fase_ercard(driver, config, log)
            duracao = (time.monotonic() - inicio_relatorio) / 60
            log(f"[ERCard][IMPORTACAO] Arquivo salvo na pasta de destino: {arquivo}")
            if not arquivo.exists():
                raise RuntimeError(f"Arquivo exportado não encontrado: {arquivo}")
            if arquivo.stat().st_size == 0:
                raise RuntimeError(f"Arquivo exportado vazio: {arquivo}")
            log(f"[ERCard][FIM] {nome_relatorio} concluido em {duracao:.1f} minutos")
            estado["relatorios"][str(indice)] = {
                "nome": nome_relatorio,
                "arquivo": str(arquivo),
                "duracao_minutos": round(duracao, 2),
            }
            atualizar_estado(
                estado,
                etapa=indice,
                nome=nome_relatorio,
                status="executando",
            )
            try:
                driver.quit()
            except Exception as erro_fechar:
                log(f"[ERCard][AVISO] Chrome já encerrou ou não respondeu: {erro_fechar}")
            driver = None
        total = (time.monotonic() - inicio_total) / 60
        log(f"========== FLUXO ERCard CONCLUIDO EM {total:.1f} MINUTOS ==========")
        atualizar_estado(
            estado,
            etapa=3,
            nome=CONSULTAS[2][0],
            status="concluido",
            duracao_minutos=round(total, 2),
        )
        return 0
    except Exception as erro:
        total = (time.monotonic() - inicio_total) / 60
        limpar_pagina_para_reinicio(driver, f"ercard-carteira etapa {etapa_atual}")
        log(f"[ERCard][ERRO] Processo interrompido apos {total:.1f} minutos: {erro}")
        atualizar_estado(
            estado,
            etapa=etapa_atual,
            status="erro",
            erro=str(erro),
            duracao_minutos=round(total, 2),
        )
        return 1
    finally:
        if driver is not None:
            try:
                driver.quit()
            except Exception:
                pass

if __name__ == "__main__":
    raise SystemExit(main())
