# ruff: noqa: BLE001

"""Orquestra as exportações da tarde com estado, trava e validação.

O modo padrão atualiza as planilhas oficiais autorizadas de Retorno de
Ligações e Vendas Moda, além de salvar cópias dos arquivos exportados.
Use ``--somente-exportacoes`` quando quiser apenas baixar cópias.
"""

from __future__ import annotations

import argparse
import json
import os
import shutil
import tempfile
import time
from datetime import datetime
from pathlib import Path

import completo
import ercard
import gestor

ROOT = Path(__file__).resolve().parent
PASTA_DESTINO = completo.PASTA_EXPORTACOES_AUTOMACAO
PASTA_LOGS = ROOT / "logs"
ARQUIVO_ESTADO = PASTA_LOGS / "sequencia_exportacoes_estado.json"
ARQUIVO_TRAVA = PASTA_LOGS / "sequencia_exportacoes.lock"


def salvar_estado(**campos) -> None:
    """Atualiza o estado de forma atômica para diagnóstico e retomada."""
    PASTA_LOGS.mkdir(parents=True, exist_ok=True)
    estado = {
        "fluxo": "sequencia-exportacoes",
        "atualizado_em": datetime.now().astimezone().isoformat(timespec="seconds"),
        **campos,
    }
    temporario = ARQUIVO_ESTADO.with_suffix(".tmp")
    temporario.write_text(
        json.dumps(estado, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    temporario.replace(ARQUIVO_ESTADO)


def adquirir_trava() -> None:
    """Impede duas sequências de navegador de rodarem ao mesmo tempo."""
    PASTA_LOGS.mkdir(parents=True, exist_ok=True)
    if ARQUIVO_TRAVA.exists():
        try:
            texto = ARQUIVO_TRAVA.read_text(encoding="utf-8")
            pid = int(texto.split("pid=", 1)[1].split()[0])
            os.kill(pid, 0)
        except (FileNotFoundError, IndexError, OSError, ValueError):
            ARQUIVO_TRAVA.unlink(missing_ok=True)
    try:
        fd = os.open(ARQUIVO_TRAVA, os.O_CREAT | os.O_EXCL | os.O_WRONLY)
        try:
            os.write(fd, f"pid={os.getpid()} inicio={time.time()}".encode())
        finally:
            os.close(fd)
    except FileExistsError as erro:
        raise RuntimeError(
            f"Já existe uma execução da sequência: {ARQUIVO_TRAVA}"
        ) from erro


def liberar_trava() -> None:
    ARQUIVO_TRAVA.unlink(missing_ok=True)


def salvar_download(arquivo, nome: str, carimbo: str):
    origem = Path(arquivo)
    if not origem.is_file() or origem.stat().st_size <= 0:
        raise RuntimeError(f"Arquivo baixado inválido: {origem}")
    destino = PASTA_DESTINO / f"{nome}_{carimbo}_importar{origem.suffix.lower()}"
    destino.unlink(missing_ok=True)
    shutil.copy2(origem, destino)
    if not destino.is_file() or destino.stat().st_size <= 0:
        raise RuntimeError(f"Cópia final inválida: {destino}")
    print(f"[SEQUENCIA][SALVO] {destino}", flush=True)
    return destino


def executar_etapa(nome: str, funcao, etapas: dict):
    inicio = time.monotonic()
    etapas[nome] = {"status": "executando", "inicio": datetime.now().isoformat()}
    salvar_estado(status="executando", etapa=nome, etapas=etapas)
    try:
        resultado = funcao()
        duracao = (time.monotonic() - inicio) / 60
        etapas[nome] = {
            "status": "concluido",
            "duracao_minutos": round(duracao, 2),
            "fim": datetime.now().isoformat(),
        }
        salvar_estado(status="executando", etapa=nome, etapas=etapas)
        print(f"[SEQUENCIA][OK] {nome} em {duracao:.1f} min", flush=True)
        return resultado
    except Exception as erro:
        duracao = (time.monotonic() - inicio) / 60
        etapas[nome] = {
            "status": "erro",
            "erro": str(erro),
            "duracao_minutos": round(duracao, 2),
            "fim": datetime.now().isoformat(),
        }
        salvar_estado(status="erro", etapa=nome, etapas=etapas)
        raise


def parse_args(argv=None):
    parser = argparse.ArgumentParser(
        description="Exporta os arquivos da Crediagora em uma única sequência."
    )
    parser.add_argument(
        "--check",
        action="store_true",
        help="Valida importações e pastas sem abrir o navegador.",
    )
    parser.add_argument(
        "--atualizar-gestor",
        action="store_true",
        help="Força a atualização da planilha oficial do Gestor.",
    )
    parser.add_argument(
        "--atualizar-retorno-oficial",
        action="store_true",
        help="Força a atualização do Retorno de Ligações oficial.",
    )
    parser.add_argument(
        "--somente-exportacoes",
        action="store_true",
        help="Não altera planilhas; salva somente as cópias exportadas.",
    )
    parser.add_argument(
        "--sem-ercard",
        action="store_true",
        help="Não executa a etapa 14 - Tabela Contratos.",
    )
    return parser.parse_args(argv)


def validar_check(args) -> int:
    PASTA_DESTINO.mkdir(parents=True, exist_ok=True)
    if not PASTA_DESTINO.is_dir():
        raise RuntimeError(f"Pasta de exportação inválida: {PASTA_DESTINO}")
    validar_destinos_oficiais(args)
    print(f"[SEQUENCIA][CHECK] destino={PASTA_DESTINO}")
    if args.somente_exportacoes:
        print("[SEQUENCIA][CHECK] modo somente exportações")
    else:
        print("[SEQUENCIA][CHECK] modo padrão atualiza Retorno e Vendas Moda oficiais")
    return 0


def validar_destinos_oficiais(args) -> None:
    """Impede execução se os destinos oficiais tiverem sido redirecionados."""
    if args.somente_exportacoes:
        return
    retorno_oficial = completo.ARQUIVO_RETORNO_LIGACOES_OFICIAL.resolve()
    retorno_configurado = completo.ARQUIVO_RETORNO_LIGACOES.resolve()
    if not retorno_oficial.is_file():
        raise FileNotFoundError(f"Retorno oficial não encontrado: {retorno_oficial}")
    if retorno_configurado != retorno_oficial:
        raise RuntimeError(
            "Proteção acionada: Retorno de Ligações não aponta para o arquivo oficial."
        )
    config_gestor = completo.criar_configuracao_gestor()
    gestor_configurado = Path(config_gestor.arquivo_excel).resolve()
    metas_oficial = completo.ARQUIVO_METAS_GESTOR.resolve()
    if not metas_oficial.is_file():
        raise FileNotFoundError(f"Vendas Moda oficial não encontrada: {metas_oficial}")
    if gestor_configurado != metas_oficial:
        raise RuntimeError(
            "Proteção acionada: Gestor não aponta para a planilha Vendas Moda oficial."
        )


def _abrir_panorama():
    """Abre o Panorama invisível; ERCard continua visível por exigência técnica."""
    driver = completo.iniciar_chrome(headless=True, manter_navegador=False)
    completo.fazer_login(driver, preparar_janela=False)
    return driver


def _executar_com_trava(args) -> int:
    PASTA_DESTINO.mkdir(parents=True, exist_ok=True)
    carimbo = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    etapas: dict[str, dict] = {}
    falhas: list[str] = []
    driver = None
    atualizar_retorno_original = completo.atualizar_retorno_ligacoes

    salvar_estado(
        status="executando",
        etapa="inicializacao",
        inicio=datetime.now().astimezone().isoformat(timespec="seconds"),
        modo={
            "atualizar_gestor": not args.somente_exportacoes,
            "atualizar_retorno_oficial": not args.somente_exportacoes,
            "sem_ercard": args.sem_ercard,
        },
        etapas=etapas,
        arquivos=[],
    )

    with tempfile.TemporaryDirectory(prefix="downloads_sequencia_") as temporario:
        completo.PASTA_DOWNLOAD = Path(temporario)
        print(
            "[SEQUENCIA] Ordem: ERCard Tabela Contratos -> Panorama -> Gestor",
            flush=True,
        )

        def registrar_falha(etapa, erro):
            mensagem = f"[SEQUENCIA][ERRO][{etapa}] {erro}"
            falhas.append(mensagem)
            completo.log(mensagem)
            print(mensagem, flush=True)

        def salvar_retorno(arquivo_zip, _arquivo_destino, _log, aguardar_arquivo=None):
            return salvar_download(arquivo_zip, "retorno_ligacoes", carimbo)

        def atualizar_e_salvar_retorno(
            arquivo_zip, arquivo_destino, log, aguardar_arquivo=None
        ):
            resultado = atualizar_retorno_original(
                arquivo_zip,
                arquivo_destino,
                log,
                aguardar_arquivo=aguardar_arquivo,
            )
            salvar_download(arquivo_zip, "retorno_ligacoes", carimbo)
            return resultado

        atualizar_retorno_oficial = not args.somente_exportacoes
        atualizar_gestor = not args.somente_exportacoes
        completo.atualizar_retorno_ligacoes = (
            atualizar_e_salvar_retorno if atualizar_retorno_oficial else salvar_retorno
        )

        # O ERCard vem primeiro porque é a etapa mais demorada. Assim, o
        # Panorama e o Gestor terminam mais perto do horário-limite da tarde.
        if not args.sem_ercard:

            def etapa_ercard():
                config = completo.criar_configuracao_ercard(PASTA_DESTINO)
                ercard.CONSULTA_CONTRATOS = "14 - TABELA CONTRATOS"

                def nome_arquivo_ercard(agora=None):
                    momento = agora or datetime.now()
                    return (
                        f"14_tabela_contratos_{momento:%Y-%m-%d_%H-%M-%S}_importar.csv"
                    )

                ercard.gerar_nome_tabela_contratos = nome_arquivo_ercard
                driver_ercard = completo.iniciar_chrome(
                    headless=False, manter_navegador=False
                )
                try:
                    arquivo = ercard.executar_fase_ercard(
                        driver_ercard, config, completo.log
                    )
                    if not arquivo.is_file() or arquivo.stat().st_size <= 0:
                        raise RuntimeError(f"Download ERCard inválido: {arquivo}")
                    return arquivo
                finally:
                    try:
                        driver_ercard.quit()
                    except Exception as erro:
                        completo.log(f"[ERCard][AVISO] Falha ao fechar Chrome: {erro}")

            try:
                executar_etapa("ercard_tabela_contratos", etapa_ercard, etapas)
            except Exception as erro:
                registrar_falha("ERCARD TABELA CONTRATOS", erro)

        try:
            try:
                driver = executar_etapa("panorama_login", _abrir_panorama, etapas)
            except Exception as erro:
                registrar_falha("PANORAMA LOGIN", erro)

            if driver is not None:

                def etapa_retorno():
                    arquivo = completo.processar_retorno_ligacoes(driver)
                    if atualizar_retorno_oficial:
                        print(
                            "[SEQUENCIA][RETORNO] planilha oficial atualizada",
                            flush=True,
                        )
                    return arquivo

                try:
                    executar_etapa("retorno_ligacoes", etapa_retorno, etapas)
                except Exception as erro:
                    registrar_falha("RETORNO LIGACOES", erro)

                def etapa_volume():
                    completo.entrar_em_vendas_e_emprestimo(driver)
                    origem = completo.baixar_exportacao(driver, "Exportação vendas")
                    return salvar_download(origem, "volume", carimbo)

                try:
                    executar_etapa("volume", etapa_volume, etapas)
                except Exception as erro:
                    registrar_falha("VOLUME", erro)

                def etapa_receita():
                    completo.voltar_para_emprestimos(driver)
                    origem = completo.baixar_exportacao(driver, "Receita gerada")
                    return salvar_download(origem, "receita", carimbo)

                try:
                    executar_etapa("receita", etapa_receita, etapas)
                except Exception as erro:
                    registrar_falha("RECEITA", erro)

                try:
                    driver.quit()
                except Exception as erro:
                    registrar_falha("FECHAMENTO PANORAMA", erro)
                finally:
                    driver = None

            def etapa_gestor():
                resultado = gestor.executar_fase_gestor(
                    completo.criar_configuracao_gestor(),
                    completo.log,
                    atualizar_planilha=atualizar_gestor,
                )
                arquivo_xlsx = completo.exportar_valores_gestor(
                    resultado, completo.log, PASTA_DESTINO
                )
                arquivo_json = (
                    PASTA_DESTINO / f"gestor_valores_{carimbo}_importar.json"
                )
                arquivo_json.write_text(
                    json.dumps(
                        {
                            "data_relatorio": str(resultado.data_relatorio),
                            "valores": {
                                chave: float(valor)
                                for chave, valor in resultado.valores.items()
                            },
                            "total_geral": float(resultado.total_relatorio),
                        },
                        ensure_ascii=False,
                        indent=2,
                    ),
                    encoding="utf-8",
                )
                print(f"[SEQUENCIA][SALVO] {arquivo_xlsx}", flush=True)
                print(f"[SEQUENCIA][SALVO] {arquivo_json}", flush=True)
                return arquivo_xlsx

            try:
                executar_etapa("gestor", etapa_gestor, etapas)
            except Exception as erro:
                registrar_falha("GESTOR", erro)
        finally:
            completo.atualizar_retorno_ligacoes = atualizar_retorno_original
            if driver is not None:
                try:
                    driver.quit()
                except Exception as erro:
                    registrar_falha("FECHAMENTO PANORAMA", erro)

    status = "concluido" if not falhas else "erro"
    salvar_estado(
        status=status,
        etapa="finalizado",
        fim=datetime.now().astimezone().isoformat(timespec="seconds"),
        etapas=etapas,
        falhas=falhas,
    )
    print(
        f"[SEQUENCIA] {status}. Arquivos salvos em: {PASTA_DESTINO}",
        flush=True,
    )
    return 0 if not falhas else 1


def main(argv=None) -> int:
    args = parse_args(argv)
    if args.check:
        try:
            return validar_check(args)
        except Exception as erro:
            print(f"[SEQUENCIA][FATAL] {erro}", flush=True)
            return 2
    if args.somente_exportacoes and (
        args.atualizar_gestor or args.atualizar_retorno_oficial
    ):
        print(
            "[SEQUENCIA][FATAL] --somente-exportacoes não pode ser combinado "
            "com opções de atualização oficial.",
            flush=True,
        )
        return 2
    try:
        validar_destinos_oficiais(args)
    except Exception as erro:
        print(f"[SEQUENCIA][FATAL] {erro}", flush=True)
        return 2
    try:
        adquirir_trava()
    except RuntimeError as erro:
        print(f"[SEQUENCIA][BLOQUEADA] {erro}", flush=True)
        return 2
    try:
        return _executar_com_trava(args)
    except Exception as erro:
        salvar_estado(
            status="erro",
            etapa="fatal",
            erro=str(erro),
            fim=datetime.now().astimezone().isoformat(timespec="seconds"),
        )
        print(f"[SEQUENCIA][FATAL] {erro}", flush=True)
        return 1
    finally:
        liberar_trava()


if __name__ == "__main__":
    raise SystemExit(main())
