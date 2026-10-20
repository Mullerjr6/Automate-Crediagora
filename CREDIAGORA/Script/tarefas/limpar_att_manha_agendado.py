"""Esvazia com segurança a pasta de exportações ATT manhã."""

from __future__ import annotations

import shutil
from datetime import datetime
from pathlib import Path

PASTA_ATT_MANHA = Path.home() / "Desktop" / "ATT manhã"
LOG = Path(__file__).resolve().parent.parent / "logs" / "limpeza_att_manha.log"


def registrar(mensagem: str) -> None:
    linha = f"[{datetime.now():%d/%m/%Y %H:%M:%S}] {mensagem}"
    print(linha, flush=True)
    LOG.parent.mkdir(parents=True, exist_ok=True)
    with LOG.open("a", encoding="utf-8") as arquivo:
        arquivo.write(linha + "\n")


def main() -> int:
    destino_esperado = (Path.home() / "Desktop" / "ATT manhã").resolve()
    if PASTA_ATT_MANHA.resolve() != destino_esperado:
        raise RuntimeError("Proteção acionada: destino de limpeza inesperado.")
    PASTA_ATT_MANHA.mkdir(parents=True, exist_ok=True)

    removidos = 0
    for item in PASTA_ATT_MANHA.iterdir():
        if item.is_symlink() or item.is_file():
            item.unlink()
        elif item.is_dir():
            shutil.rmtree(item)
        removidos += 1
    registrar(f"Pasta ATT manhã restaurada: {removidos} item(ns) removido(s).")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
