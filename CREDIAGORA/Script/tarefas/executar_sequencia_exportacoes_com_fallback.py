"""Executa a sequência de downloads e tenta o Tabela Contratos se ela falhar."""

from __future__ import annotations

import subprocess
import sys
from datetime import datetime
from pathlib import Path


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
    with LOG.open("a", encoding="utf-8") as arquivo_log:
        arquivo_log.write(f"[FALLBACK] Início: {datetime.now():%Y-%m-%d %H:%M:%S}\n")
        resultado = executar("sequencia_exportacoes.py", arquivo_log)
        arquivo_log.write(f"[FALLBACK] Sequência terminou com código {resultado}\n")
        if resultado == 0:
            return 0
        arquivo_log.write("[FALLBACK] Executando somente 14 - Tabela Contratos.\n")
        fallback = executar("tabela_contratos.py", arquivo_log)
        arquivo_log.write(f"[FALLBACK] Tabela Contratos terminou com código {fallback}\n")
        return fallback


if __name__ == "__main__":
    raise SystemExit(main())
