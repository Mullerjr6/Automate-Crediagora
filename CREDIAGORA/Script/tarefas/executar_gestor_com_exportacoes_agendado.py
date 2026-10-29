"""Baixa Vendas/Receita e, em seguida, atualiza o Gestor oficial."""

from __future__ import annotations

import argparse
import subprocess
import sys
from datetime import date

from executar_retorno_ligacoes_oficial_agendado import (
    SCRIPT_DIR,
    feriados_nacionais,
)

LOG = SCRIPT_DIR / "logs" / "agendado_gestor_com_exportacoes.log"


def registrar(mensagem: str) -> None:
    linha = f"[{date.today():%d/%m/%Y}] {mensagem}"
    print(linha, flush=True)
    LOG.parent.mkdir(parents=True, exist_ok=True)
    with LOG.open("a", encoding="utf-8") as arquivo:
        arquivo.write(linha + "\n")


def executar(comando: list[str], nome: str) -> int:
    registrar(f"Iniciando {nome}.")
    with LOG.open("a", encoding="utf-8") as arquivo_log:
        resultado = subprocess.run(
            comando,
            cwd=SCRIPT_DIR,
            check=False,
            stdout=arquivo_log,
            stderr=subprocess.STDOUT,
            text=True,
        )
    registrar(f"{nome} finalizado com código {resultado.returncode}.")
    return resultado.returncode


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(
        description="Valida e executa as exportações Panorama seguidas do Gestor."
    )
    parser.add_argument(
        "--check",
        action="store_true",
        help="Valida os scripts sem abrir o Chrome nem atualizar planilhas.",
    )
    args = parser.parse_args(argv)
    if args.check:
        registrar("[CHECK] Sequência configurada: Panorama Vendas/Receita -> Gestor oficial.")
        return 0

    hoje = date.today()
    if hoje.weekday() >= 5:
        registrar("Fim de semana: execução não iniciada.")
        return 0
    feriado = feriados_nacionais(hoje.year).get(hoje)
    if feriado:
        registrar(f"Feriado nacional ({feriado}): execução não iniciada.")
        return 0

    codigo_exportacoes = executar(
        [sys.executable, "-u", "panorama_exportacoes_carteira.py"],
        "exportações Panorama Vendas e Receita",
    )
    if codigo_exportacoes != 0:
        registrar("Gestor não iniciado: as exportações Panorama falharam.")
        return codigo_exportacoes

    return executar(
        [sys.executable, "-u", "completo.py", "--somente-gestor"],
        "Gestor oficial",
    )


if __name__ == "__main__":
    raise SystemExit(main())
