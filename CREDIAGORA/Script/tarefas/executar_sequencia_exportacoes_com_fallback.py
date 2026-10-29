"""Executa a sequência oficial e recupera apenas as etapas que falharem."""

from __future__ import annotations

import subprocess
import sys
from datetime import datetime
from pathlib import Path

from executar_sequencia_tarde_agendado import (
    etapas_com_erro,
    executar_fallback_final_14,
)

ROOT = Path(__file__).resolve().parents[1]
LOG_DIR = ROOT / "logs"
LOG_DIR.mkdir(parents=True, exist_ok=True)
LOG = LOG_DIR / f"sequencia_exportacoes_{datetime.now():%Y-%m-%d_%H-%M-%S}.log"


def executar(script: str, arquivo_log) -> int:
    return subprocess.run(
        [sys.executable, str(ROOT / script)],
        cwd=ROOT,
        stdout=arquivo_log,
        stderr=subprocess.STDOUT,
        check=False,
    ).returncode


def main() -> int:
    inicio = datetime.now().timestamp()
    with LOG.open("a", encoding="utf-8") as arquivo_log:
        arquivo_log.write(
            f"[FALLBACK] Início: {datetime.now():%Y-%m-%d %H:%M:%S}\n"
        )
        resultado = executar("sequencia_exportacoes.py", arquivo_log)
        arquivo_log.write(
            f"[FALLBACK] Sequência terminou com código {resultado}\n"
        )
        if resultado == 0:
            return 0
        erros = etapas_com_erro(inicio)
        arquivo_log.write(
            f"[FALLBACK] Etapas com erro identificadas: {sorted(erros) or 'nenhuma'}\n"
        )
        return executar_fallback_final_14(erros, sem_ercard=False, inicio=inicio)


if __name__ == "__main__":
    raise SystemExit(main())
