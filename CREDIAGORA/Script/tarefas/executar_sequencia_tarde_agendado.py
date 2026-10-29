"""Executa a sequência consolidada da tarde em dias úteis."""

from __future__ import annotations

import argparse
import json
import os
import subprocess
import sys
from datetime import date, datetime
from pathlib import Path

from executar_retorno_ligacoes_oficial_agendado import feriados_nacionais

SCRIPT_DIR = Path(__file__).resolve().parent.parent
LOG = SCRIPT_DIR / "logs" / "sequencia_tarde_agendado.log"
ESTADO = SCRIPT_DIR / "logs" / "sequencia_exportacoes_estado.json"
ESTADO_CONTRATOS = SCRIPT_DIR / "logs" / "ercard_tabela_contratos_estado.json"


def registrar(mensagem: str) -> None:
    linha = f"[{datetime.now():%d/%m/%Y %H:%M:%S}] {mensagem}"
    print(linha, flush=True)
    LOG.parent.mkdir(parents=True, exist_ok=True)
    with LOG.open("a", encoding="utf-8") as arquivo:
        arquivo.write(linha + "\n")


def executar(comando: list[str], nome: str) -> int:
    registrar(f"Iniciando {nome}.")
    with LOG.open("a", encoding="utf-8") as arquivo:
        resultado = subprocess.run(
            comando,
            cwd=SCRIPT_DIR,
            stdout=arquivo,
            stderr=subprocess.STDOUT,
            check=False,
            text=True,
        )
    registrar(f"{nome} finalizado com código {resultado.returncode}.")
    return resultado.returncode


def tabela_contratos_concluida(inicio: float) -> bool:
    """Confirma o 14 concluído na sequência ou no fluxo próprio do ERCard."""
    try:
        if ESTADO.exists() and ESTADO.stat().st_mtime >= inicio:
            dados = json.loads(ESTADO.read_text(encoding="utf-8"))
            etapa = dados.get("etapas", {}).get("ercard_tabela_contratos", {})
            if etapa.get("status") == "concluido":
                return True

        if not ESTADO_CONTRATOS.exists() or ESTADO_CONTRATOS.stat().st_mtime < inicio:
            return False
        dados = json.loads(ESTADO_CONTRATOS.read_text(encoding="utf-8"))
        arquivo = Path(dados.get("arquivo", ""))
        return (
            dados.get("status") == "concluido"
            and arquivo.is_file()
            and "_importar" in arquivo.stem
        )
    except (OSError, json.JSONDecodeError, TypeError):
        return False


def etapas_com_erro(inicio: float) -> set[str]:
    """Lê somente o estado gravado pela execução atual."""
    try:
        if not ESTADO.exists() or ESTADO.stat().st_mtime < inicio:
            return set()
        dados = json.loads(ESTADO.read_text(encoding="utf-8"))
        etapas = dados.get("etapas", {})
        return {
            nome
            for nome, etapa in etapas.items()
            if isinstance(etapa, dict) and etapa.get("status") == "erro"
        }
    except (OSError, json.JSONDecodeError, TypeError):
        return set()


def executar_fallbacks(erros: set[str], sem_ercard: bool) -> int:
    """Recupera apenas as etapas que falharam, uma vez cada."""
    resultados: list[int] = []

    if erros & {"panorama_login", "retorno_ligacoes"}:
        resultados.append(
            executar(
                [
                    sys.executable,
                    "-u",
                    str(SCRIPT_DIR / "agente_crediagora.py"),
                    "--fluxo",
                    "retorno-ligacoes-oficial",
                ],
                "Fallback Retorno de Ligações oficial",
            )
        )

    if erros & {"panorama_login", "volume", "receita"}:
        resultados.append(
            executar(
                [
                    sys.executable,
                    "-u",
                    str(SCRIPT_DIR / "panorama_exportacoes_carteira.py"),
                ],
                "Fallback Panorama Vendas e Receita",
            )
        )

    if "gestor" in erros:
        resultados.append(
            executar(
                [
                    sys.executable,
                    "-u",
                    str(SCRIPT_DIR / "completo.py"),
                    "--somente-gestor",
                ],
                "Fallback Gestor oficial",
            )
        )

    if "ercard_tabela_contratos" in erros and not sem_ercard:
        resultados.append(
            executar(
                [
                    sys.executable,
                    "-u",
                    str(SCRIPT_DIR / "tabela_contratos.py"),
                ],
                "Fallback 14 - Tabela Contratos",
            )
        )

    if not resultados:
        registrar("Nenhuma etapa recuperável identificada; fallback não iniciado.")
        return 1
    return 0 if all(resultado == 0 for resultado in resultados) else 1


def executar_fallback_final_14(
    erros: set[str],
    sem_ercard: bool,
    inicio: float | None = None,
) -> int:
    """Reprocessa etapas com erro e preserva o 14 como última saída segura.

    A sequência pode atualizar várias fontes independentes. Se a recuperação
    específica não resolver o problema, insistir nas mesmas telas costuma
    aumentar o risco de travar o portal. Nesse caso, fazemos somente a
    exportação do 14 - Tabela Contratos, que é o artefato mínimo confiável.
    """
    resultado_reprocessamento = executar_fallbacks(erros, sem_ercard)
    if resultado_reprocessamento == 0:
        registrar("Reprocessamento das etapas com erro concluído.")
        return 0

    if sem_ercard:
        registrar(
            "Reprocessamento falhou, mas a porta final do 14 foi desativada "
            "por --sem-ercard."
        )
        return resultado_reprocessamento

    if inicio is not None and tabela_contratos_concluida(inicio):
        registrar(
            "[PORTA FINAL] O 14 - Tabela Contratos já foi concluído nesta "
            "execução; não haverá nova abertura do ERCard."
        )
        return resultado_reprocessamento

    registrar(
        "[PORTA FINAL] Reprocessamento sem sucesso; executando somente "
        "14 - Tabela Contratos."
    )
    resultado_final = executar(
        [
            sys.executable,
            "-u",
            str(SCRIPT_DIR / "tabela_contratos.py"),
        ],
        "Porta final 14 - Tabela Contratos",
    )
    if resultado_final == 0:
        registrar(
            "[PORTA FINAL] 14 - Tabela Contratos concluído; as demais etapas "
            "permanecem sinalizadas para diagnóstico."
        )
    else:
        registrar(
            "[PORTA FINAL][ERRO] Não foi possível concluir o 14 - "
            "Tabela Contratos."
        )
    return resultado_final


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(
        description="Executa exportações e atualizações oficiais em uma única sequência."
    )
    parser.add_argument(
        "--check",
        action="store_true",
        help="Valida somente a configuração, sem abrir o navegador.",
    )
    parser.add_argument(
        "--sem-ercard",
        action="store_true",
        help="Executa a sequência sem a etapa 14 - Tabela Contratos.",
    )
    args = parser.parse_args(argv)

    registrar(
        "Tarefa recebida pelo Agendador; "
        f"pid={os.getpid()} session={os.environ.get('SESSIONNAME', 'desconhecida')}"
    )

    if args.check:
        registrar(
            "[CHECK] Sequência configurada: exportações + Gestor oficial + Retorno oficial."
        )
        return 0

    hoje = date.today()
    if hoje.weekday() >= 5:
        registrar("Fim de semana: execução não iniciada.")
        return 0
    feriado = feriados_nacionais(hoje.year).get(hoje)
    if feriado:
        registrar(f"Feriado nacional ({feriado}): execução não iniciada.")
        return 0

    comando = [
        sys.executable,
        "-u",
        str(SCRIPT_DIR / "sequencia_exportacoes.py"),
        "--atualizar-gestor",
        "--atualizar-retorno-oficial",
    ]
    if args.sem_ercard:
        comando.append("--sem-ercard")

    inicio = datetime.now().timestamp()
    resultado = executar(comando, "Sequência consolidada da tarde")
    if resultado == 0:
        return resultado

    erros = etapas_com_erro(inicio)
    if "ercard_tabela_contratos" not in erros and tabela_contratos_concluida(inicio):
        registrar("Tabela Contratos concluído na sequência; fallback do ERCard ignorado.")
    registrar(f"Sequência com falha; etapas para recuperação: {sorted(erros) or 'nenhuma'}.")
    return executar_fallback_final_14(erros, args.sem_ercard, inicio)


if __name__ == "__main__":
    raise SystemExit(main())
