"""Permite executar os testes manuais sem mudar os modulos de producao."""

import sys
from pathlib import Path


PASTA_SCRIPT = Path(__file__).resolve().parents[1]
if str(PASTA_SCRIPT) not in sys.path:
    sys.path.insert(0, str(PASTA_SCRIPT))
