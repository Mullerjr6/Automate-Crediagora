from pathlib import Path

# ============================================================
# LIMPAR PASTA DE DOWNLOADS DO CREDIAGORA
# ============================================================
# Este script apaga os arquivos da pasta:
# C:\\Users\\CREDIAGORA\\Downloads\\crediagora
#
# Por segurança, ele apaga somente arquivos Excel, CSV e arquivos
# temporários de download do Chrome.
# ============================================================

PASTA_DOWNLOAD = Path(r"C:\Users\CREDIAGORA\Downloads\crediagora")

EXTENSOES_PARA_APAGAR = [
    "*.xlsx",
    "*.xls",
    "*.xlsm",
    "*.xlsb",
    "*.csv",
    "*.crdownload"
]


def limpar_downloads():
    if not PASTA_DOWNLOAD.exists():
        print(f"Pasta não encontrada: {PASTA_DOWNLOAD}")
        return

    arquivos_apagados = 0

    for extensao in EXTENSOES_PARA_APAGAR:
        for arquivo in PASTA_DOWNLOAD.glob(extensao):
            try:
                if arquivo.is_file():
                    arquivo.unlink()
                    arquivos_apagados += 1
                    print(f"Apagado: {arquivo.name}")
            except Exception as erro:
                print(f"Não foi possível apagar {arquivo.name}: {erro}")

    print("----------------------------------------")
    print(f"Limpeza finalizada. Arquivos apagados: {arquivos_apagados}")


if __name__ == "__main__":
    limpar_downloads()
