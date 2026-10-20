"""Dispara o relatório 14 do ERCard em dias úteis não feriados."""

from __future__ import annotations

import subprocess
import sys
from datetime import date

from executar_retorno_ligacoes_oficial_agendado import SCRIPT_DIR, feriados_nacionais

LOG = SCRIPT_DIR / "logs" / "agendado_ercard_tabela_contratos.log"


def registrar(mensagem: str) -> None:
    linha = f"[{date.today():%d/%m/%Y}] {mensagem}"
    print(linha, flush=True)
    LOG.parent.mkdir(parents=True, exist_ok=True)
    with LOG.open("a", encoding="utf-8") as arquivo:
        arquivo.write(linha + "\n")


def main() -> int:
    hoje = date.today()
    if hoje.weekday() >= 5:
        registrar("Fim de semana: execução não iniciada.")
        return 0

    feriado = feriados_nacionais(hoje.year).get(hoje)
    if feriado:
        registrar(f"Feriado nacional ({feriado}): execução não iniciada.")
        return 0

    registrar("Iniciando agente do ERCard – 14 Tabela Contratos.")
    with LOG.open("a", encoding="utf-8") as arquivo_log:
        resultado = subprocess.run(
            [
                sys.executable,
                "-u",
                "agente_crediagora.py",
                "--fluxo",
                "ercard-tabela-contratos",
            ],
            cwd=SCRIPT_DIR,
            check=False,
            stdout=arquivo_log,
            stderr=subprocess.STDOUT,
            text=True,
        )
    registrar(f"Agente finalizado com código {resultado.returncode}.")
    return resultado.returncode


if __name__ == "__main__":
    raise SystemExit(main())
