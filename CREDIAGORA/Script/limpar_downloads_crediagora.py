import argparse
import os
from pathlib import Path


PASTA_DOWNLOAD = Path(
    os.getenv(
        "CREDIAGORA_DOWNLOAD_DIR",
        str(Path.home() / "Downloads" / "crediagora")
    )
).expanduser()

PADROES_DOWNLOAD = [
    "*.xlsx",
    "*.xls",
    "*.xlsm",
    "*.xlsb",
    "*.csv",
    "*.zip",
    "*.crdownload",
    "*.tmp",
    "*.download",
    "*.part",
]


def validar_pasta_download_segura():
    caminho = PASTA_DOWNLOAD.resolve()
    home = Path.home().resolve()
    caminhos_proibidos = [
        home,
        (home / "Downloads").resolve(),
        (home / "Desktop").resolve(),
    ]

    if PASTA_DOWNLOAD.anchor:
        caminhos_proibidos.append(Path(PASTA_DOWNLOAD.anchor).resolve())

    if caminho in caminhos_proibidos:
        raise ValueError(
            "Pasta de downloads insegura para limpeza automática: "
            f"{PASTA_DOWNLOAD}. Use uma subpasta dedicada, como Downloads\\crediagora."
        )

    if "crediagora" not in str(caminho).lower():
        raise ValueError(
            "Pasta de downloads precisa ser dedicada à automação Crediagora. "
            f"Valor atual: {PASTA_DOWNLOAD}"
        )


def listar_arquivos_para_limpeza():
    arquivos = []

    for padrao in PADROES_DOWNLOAD:
        arquivos.extend(
            arquivo
            for arquivo in PASTA_DOWNLOAD.glob(padrao)
            if arquivo.is_file()
        )

    return sorted(set(arquivos), key=lambda arquivo: arquivo.name.lower())


def limpar_downloads(apenas_conferir=False):
    validar_pasta_download_segura()

    if not PASTA_DOWNLOAD.exists():
        print(f"Pasta não encontrada: {PASTA_DOWNLOAD}")
        return 0

    arquivos = listar_arquivos_para_limpeza()

    if apenas_conferir:
        print(f"Pasta conferida: {PASTA_DOWNLOAD}")
        print(f"Arquivos que seriam apagados: {len(arquivos)}")
        for arquivo in arquivos:
            print(f"- {arquivo.name}")
        return len(arquivos)

    arquivos_apagados = 0

    for arquivo in arquivos:
        try:
            arquivo.unlink()
            arquivos_apagados += 1
            print(f"Apagado: {arquivo.name}")
        except Exception as erro:
            print(f"Não foi possível apagar {arquivo.name}: {erro}")

    print("----------------------------------------")
    print(f"Limpeza finalizada. Arquivos apagados: {arquivos_apagados}")
    return arquivos_apagados


def parse_args():
    parser = argparse.ArgumentParser(
        description="Limpa arquivos exportados da pasta de downloads do Crediagora."
    )
    parser.add_argument(
        "--check",
        action="store_true",
        help="Mostra quais arquivos seriam apagados, sem remover nada."
    )
    return parser.parse_args()


if __name__ == "__main__":
    argumentos = parse_args()
    limpar_downloads(apenas_conferir=argumentos.check)
