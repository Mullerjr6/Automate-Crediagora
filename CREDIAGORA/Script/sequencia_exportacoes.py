"""Baixa os arquivos dos fluxos e salva somente cópias para importação."""

from __future__ import annotations

import json
import shutil
import tempfile
from datetime import datetime
from pathlib import Path

import completo
import ercard
import gestor

PASTA_DESTINO = Path(
    r"C:\Users\CREDIAGORA\TJI PROMOTORA DE VENDAS EIRELI\Crediagora-doc - dados\Automacoes\Automatização de processos Crediagora\Logs de exportações automatizadas"
)


def main() -> int:
    PASTA_DESTINO.mkdir(parents=True, exist_ok=True)
    carimbo = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    falhas: list[str] = []

    with tempfile.TemporaryDirectory(prefix="downloads_panorama_") as temporario:
        completo.PASTA_DOWNLOAD = Path(temporario)
        driver = None

        def salvar_download(arquivo, nome):
            origem = Path(arquivo)
            if not origem.is_file() or origem.stat().st_size == 0:
                raise RuntimeError(f"Arquivo baixado inválido: {origem}")
            destino = PASTA_DESTINO / f"{nome}_{carimbo}_importar{origem.suffix}"
            shutil.copy2(origem, destino)
            print(f"[DOWNLOAD][SALVO] {destino}")
            return destino

        def registrar_falha(etapa, erro):
            mensagem = f"[DOWNLOAD][ERRO][{etapa}] {erro}"
            falhas.append(mensagem)
            completo.log(mensagem)
            print(mensagem)

        def salvar_retorno(
            arquivo_zip, _arquivo_destino, _log, aguardar_arquivo=None
        ):
            return salvar_download(arquivo_zip, "retorno_ligacoes")

        atualizar_retorno_original = completo.atualizar_retorno_ligacoes
        completo.atualizar_retorno_ligacoes = salvar_retorno
        try:
            driver = completo.iniciar_chrome(headless=False, manter_navegador=False)
            completo.fazer_login(driver, preparar_janela=True)

            try:
                completo.processar_retorno_ligacoes(driver)
            except Exception as erro:  # noqa: BLE001
                registrar_falha("RETORNO LIGACOES", erro)

            try:
                completo.entrar_em_vendas_e_emprestimo(driver)
                arquivo_volume = completo.baixar_exportacao(
                    driver, "Exportação vendas"
                )
                salvar_download(arquivo_volume, "volume")
            except Exception as erro:  # noqa: BLE001
                registrar_falha("VOLUME", erro)

            try:
                completo.voltar_para_emprestimos(driver)
                arquivo_receita = completo.baixar_exportacao(driver, "Receita gerada")
                salvar_download(arquivo_receita, "receita")
            except Exception as erro:  # noqa: BLE001
                registrar_falha("RECEITA", erro)

            try:
                config_gestor = completo.criar_configuracao_gestor()
                resultado_gestor = gestor.executar_fase_gestor(
                    config_gestor,
                    completo.log,
                    atualizar_planilha=False,
                )
                arquivo_valores_xlsx = completo.exportar_valores_gestor(
                    resultado_gestor, completo.log, PASTA_DESTINO
                )
                arquivo_valores = PASTA_DESTINO / f"gestor_valores_{carimbo}_importar.json"
                arquivo_valores.write_text(
                    json.dumps(
                        {
                            "data_relatorio": str(resultado_gestor.data_relatorio),
                            "valores": {
                                chave: float(valor)
                                for chave, valor in resultado_gestor.valores.items()
                            },
                            "total_geral": float(resultado_gestor.total_relatorio),
                        },
                        ensure_ascii=False,
                        indent=2,
                    ),
                    encoding="utf-8",
                )
                print(f"[DOWNLOAD][SALVO] {arquivo_valores_xlsx}")
                print(f"[DOWNLOAD][SALVO] {arquivo_valores}")
            except Exception as erro:  # noqa: BLE001
                registrar_falha("GESTOR", erro)

            try:
                config_ercard = completo.criar_configuracao_ercard(PASTA_DESTINO)
                ercard.CONSULTA_CONTRATOS = "14 - TABELA CONTRATOS"

                def nome_arquivo_ercard(agora=None):
                    momento = agora or datetime.now()
                    return f"14_tabela_contratos_{momento:%Y-%m-%d_%H-%M-%S}_importar.csv"

                ercard.gerar_nome_tabela_contratos = nome_arquivo_ercard
                driver_ercard = completo.iniciar_chrome(
                    headless=False, manter_navegador=False
                )
                try:
                    arquivo_ercard = ercard.executar_fase_ercard(
                        driver_ercard, config_ercard, completo.log
                    )
                    if (
                        not arquivo_ercard.is_file()
                        or arquivo_ercard.stat().st_size == 0
                    ):
                        raise RuntimeError(f"Download ERCard inválido: {arquivo_ercard}")
                    print(f"[DOWNLOAD][SALVO] {arquivo_ercard}")
                finally:
                    try:
                        driver_ercard.quit()
                    except Exception:  # noqa: BLE001
                        completo.log("[ERCard][AVISO] Falha ao fechar Chrome.")
            except Exception as erro:  # noqa: BLE001
                registrar_falha("ERCARD TABELA CONTRATOS", erro)

            print(f"[DOWNLOAD] Concluído. Arquivos salvos em: {PASTA_DESTINO}")
            return 1 if falhas else 0
        except Exception as erro:  # noqa: BLE001
            registrar_falha("INICIALIZACAO", erro)
            return 1
        finally:
            completo.atualizar_retorno_ligacoes = atualizar_retorno_original
            if driver is not None:
                try:
                    driver.quit()
                except Exception:  # noqa: BLE001
                    completo.log("[PANORAMA][AVISO] Falha ao fechar Chrome.")


if __name__ == "__main__":
    raise SystemExit(main())
