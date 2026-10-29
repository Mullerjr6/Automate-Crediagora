"""Executa exclusivamente a atualização automática do Retorno de Ligações oficial."""

import os
from pathlib import Path

PASTA_OFICIAL = (
    Path.home()
    / "TJI PROMOTORA DE VENDAS EIRELI"
    / "Crediagora-doc - dados"
    / "tabela fat"
)
ARQUIVO_OFICIAL = PASTA_OFICIAL / "fat_retorno_ligacoes.xlsx"


def main():
    if not ARQUIVO_OFICIAL.is_file():
        raise FileNotFoundError(f"Arquivo oficial não encontrado: {ARQUIVO_OFICIAL}")

    # Deve ocorrer antes da importação de completo.py, que resolve os caminhos na carga.
    os.environ["CREDIAGORA_TABELA_FAT_DIR"] = str(PASTA_OFICIAL)
    os.environ["CREDIAGORA_RUNTIME_DIR"] = str(Path(__file__).resolve().parent.parent)

    import completo

    # O completo.py usa nomes de arquivos de teste por padrão. Neste fluxo
    # explicitamente autorizado, substitui-se somente o destino do Retorno.
    completo.ARQUIVO_RETORNO_LIGACOES = ARQUIVO_OFICIAL

    if completo.ARQUIVO_RETORNO_LIGACOES.resolve() != ARQUIVO_OFICIAL.resolve():
        raise RuntimeError("Proteção acionada: o fluxo não está apontando para o Retorno de Ligações oficial.")
    return completo.main([
        "--exportacao", "retorno-ligacoes",
        "--headless",
        "--sem-ercard",
        "--sem-gestor",
    ])


if __name__ == "__main__":
    raise SystemExit(main())
