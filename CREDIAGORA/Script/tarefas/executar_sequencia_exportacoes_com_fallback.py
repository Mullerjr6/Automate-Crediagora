"""Executa a sequência de downloads e tenta o Tabela Contratos se ela falhar."""

from __future__ import annotations

import json
import subprocess
import sys
from datetime import datetime
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
LOG_DIR = ROOT / "logs"
LOG_DIR.mkdir(parents=True, exist_ok=True)
LOG = LOG_DIR / f"sequencia_exportacoes_{datetime.now():%Y-%m-%d_%H-%M-%S}.log"
ESTADO_TABELA = LOG_DIR / "ercard_tabela_contratos_estado.json"


def executar(script: str, arquivo_log) -> int:
    return subprocess.run(
        [sys.executable, str(ROOT / script)],
        cwd=ROOT,
        stdout=arquivo_log,
        stderr=subprocess.STDOUT,
        check=False,
    ).returncode


def tabela_contratos_concluida_nesta_execucao(inicio: float) -> bool:
    """Evita repetir o ERCard quando só uma etapa anterior falhou."""
    try:
        if not ESTADO_TABELA.exists() or ESTADO_TABELA.stat().st_mtime < inicio:
            return False
        estado = json.loads(ESTADO_TABELA.read_text(encoding="utf-8"))
        arquivo = Path(estado.get("arquivo", ""))
        return (
            estado.get("status") == "concluido"
            and arquivo.is_file()
            and "_importar" in arquivo.stem
        )
    except (OSError, json.JSONDecodeError, TypeError):
        return False


def main() -> int:
    inicio = datetime.now().timestamp()
    with LOG.open("a", encoding="utf-8") as arquivo_log:
        arquivo_log.write(f"[FALLBACK] Início: {datetime.now():%Y-%m-%d %H:%M:%S}\n")
        resultado = executar("sequencia_exportacoes.py", arquivo_log)
        arquivo_log.write(f"[FALLBACK] Sequência terminou com código {resultado}\n")
        if resultado == 0:
            return 0
        if tabela_contratos_concluida_nesta_execucao(inicio):
            arquivo_log.write(
                "[FALLBACK] Tabela Contratos já concluída nesta execução; fallback ignorado.\n"
            )
            return resultado
        arquivo_log.write("[FALLBACK] Executando somente 14 - Tabela Contratos.\n")
        fallback = executar("tabela_contratos.py", arquivo_log)
        arquivo_log.write(f"[FALLBACK] Tabela Contratos terminou com código {fallback}\n")
        return fallback


if __name__ == "__main__":
    raise SystemExit(main())
