"""Baixa arquivos do Panorama e somente os salva na pasta de exportações."""

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

    atualizar_metas_original = gestor.atualizar_excel_metas

    def nao_atualizar_metas(*args, **kwargs):
        completo.log("[GESTOR] Valores capturados; planilha não será salva.")
        return {}, None

    with tempfile.TemporaryDirectory(prefix="downloads_panorama_") as temporario:
        pasta_download = Path(temporario)
        completo.PASTA_DOWNLOAD = pasta_download
        driver = None

        def salvar_download(arquivo, nome):
            origem = Path(arquivo)
            destino = PASTA_DESTINO / f"{nome}_{carimbo}_importar{origem.suffix}"
            shutil.copy2(origem, destino)
            print(f"[DOWNLOAD][SALVO] {destino}")
            return destino

        def salvar_retorno(arquivo_zip, arquivo_destino, log, aguardar_arquivo=None):
            return salvar_download(arquivo_zip, "retorno_ligacoes")

        atualizar_retorno_original = completo.atualizar_retorno_ligacoes
        completo.atualizar_retorno_ligacoes = salvar_retorno
        gestor.atualizar_excel_metas = nao_atualizar_metas
        try:
            # Gestor sempre primeiro: somente captura os valores do relatório.
            driver = completo.iniciar_chrome(headless=False, manter_navegador=False)
            completo.fazer_login(driver, preparar_janela=True)
            completo.processar_retorno_ligacoes(driver)

            completo.entrar_em_vendas_e_emprestimo(driver)
            salvar_download(completo.baixar_exportacao(driver, "Exportação vendas"), "volume")

            completo.voltar_para_emprestimos(driver)
            salvar_download(completo.baixar_exportacao(driver, "Receita gerada"), "receita")

            # ERCard: somente o relatório 14 - Tabela Contratos, por último.
            resultado_gestor = gestor.executar_fase_gestor(
                completo.criar_configuracao_gestor(), completo.log
            )
            arquivo_gestor = PASTA_DESTINO / f"gestor_valores_{carimbo}_importar.json"
            arquivo_gestor.write_text(
                json.dumps(
                    {
                        "data_relatorio": str(resultado_gestor.data_relatorio),
                        "valores": {k: float(v) for k, v in resultado_gestor.valores.items()},
                        "total_geral": float(resultado_gestor.total_relatorio),
                    },
                    ensure_ascii=False,
                    indent=2,
                ),
                encoding="utf-8",
            )
            print(f"[DOWNLOAD][SALVO] {arquivo_gestor}")

            config_ercard = completo.criar_configuracao_ercard(PASTA_DESTINO)
            for consulta, nome in (
                ("14 - TABELA CONTRATOS", "14_tabela_contratos"),
            ):
                ercard.CONSULTA_CONTRATOS = consulta
                ercard.gerar_nome_tabela_contratos = (
                    lambda agora=None, prefixo=nome: f"{prefixo}_{(agora or datetime.now()):%Y-%m-%d_%H-%M-%S}_importar.csv"
                )
                driver_ercard = completo.iniciar_chrome(headless=False, manter_navegador=False)
                try:
                    arquivo_ercard = ercard.executar_fase_ercard(driver_ercard, config_ercard, completo.log)
                    if not arquivo_ercard.is_file() or arquivo_ercard.stat().st_size == 0:
                        raise RuntimeError(f"Download ERCard inválido: {arquivo_ercard}")
                    print(f"[DOWNLOAD][SALVO] {arquivo_ercard}")
                finally:
                    try:
                        driver_ercard.quit()
                    except Exception:
                        pass
            print(f"[DOWNLOAD] Concluído. Arquivos salvos em: {PASTA_DESTINO}")
            return 0
        except Exception as erro:
            print(f"[DOWNLOAD][ERRO] {erro}")
            return 1
        finally:
            completo.atualizar_retorno_ligacoes = atualizar_retorno_original
            gestor.atualizar_excel_metas = atualizar_metas_original
            if driver is not None:
                try:
                    driver.quit()
                except Exception:
                    pass


if __name__ == "__main__":
    raise SystemExit(main())
