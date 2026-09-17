import csv
import os
import re
import shutil
import time
import uuid
from copy import copy
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.formula.translate import Translator
from openpyxl.utils import get_column_letter


ABA_DADOS = "Tabela Conta"
COLUNAS_DADOS = 50
COLUNA_MODELO_INICIAL = 51
COLUNA_MODELO_FINAL = 91
COLUNAS_DATA = {
    "DataContabil", "DataEfetivacao", "Nascimento", "PrimeiroVencimento",
    "DataPagamentoPrimeiraParcela",
}
COLUNAS_TEXTO = {"Proposta", "Cpf", "Fone01", "Fone02"}
CABECALHOS_VAZIOS_PERMITIDOS = {23: "Fone01", 24: "Fone02"}


class AtualizacaoIndicadoresError(RuntimeError):
    def __init__(self, etapa, mensagem):
        self.etapa = etapa
        super().__init__(f"[INDICADORES][ERRO][{etapa}] {mensagem}")


@dataclass(frozen=True)
class ResultadoAtualizacaoIndicadores:
    csv_utilizado: Path
    xlsx_atualizado: Path
    total_linhas_csv: int
    registros_importados: int
    primeira_linha: int
    ultima_linha: int
    colunas_importadas: int
    colunas_modelo: int
    formulas_modelo: int


def _log(log, mensagem):
    log(f"[INDICADORES] {mensagem}")


def _validar_csv_estavel(caminho, tentativas=3, intervalo=1):
    caminho = Path(caminho)
    if not caminho.exists():
        raise AtualizacaoIndicadoresError("CSV", f"Arquivo nao encontrado: {caminho}")
    if caminho.suffix.casefold() != ".csv":
        raise AtualizacaoIndicadoresError("CSV", f"Extensao invalida: {caminho.suffix}")
    tamanhos = []
    for indice in range(tentativas):
        tamanhos.append(caminho.stat().st_size)
        if indice + 1 < tentativas:
            time.sleep(intervalo)
    if not tamanhos[-1]:
        raise AtualizacaoIndicadoresError("CSV", "O arquivo esta vazio.")
    if len(set(tamanhos)) != 1:
        raise AtualizacaoIndicadoresError("CSV", f"Download nao estabilizado: {tamanhos}")
    return caminho


def _ler_csv(caminho):
    amostra_bytes = caminho.read_bytes()[:65536]
    for codificacao in ("utf-8-sig", "cp1252", "latin-1"):
        try:
            amostra = amostra_bytes.decode(codificacao)
            break
        except UnicodeDecodeError:
            continue
    else:
        raise AtualizacaoIndicadoresError("CSV", "Codificacao nao identificada.")
    try:
        dialeto = csv.Sniffer().sniff(amostra, delimiters=";,\t|")
    except csv.Error as erro:
        raise AtualizacaoIndicadoresError("CSV", "Delimitador nao identificado.") from erro

    linhas = []
    with caminho.open("r", encoding=codificacao, newline="") as arquivo:
        for numero, linha in enumerate(csv.reader(arquivo, dialeto), start=1):
            if len(linha) < COLUNAS_DADOS:
                raise AtualizacaoIndicadoresError(
                    "CSV", f"Linha {numero}: {len(linha)} colunas; esperado: 50."
                )
            if any(valor.strip() for valor in linha[COLUNAS_DADOS:]):
                raise AtualizacaoIndicadoresError(
                    "CSV", f"Linha {numero} possui dados depois da coluna AX."
                )
            linhas.append(linha[:COLUNAS_DADOS])
    if len(linhas) < 2:
        raise AtualizacaoIndicadoresError("CSV", "CSV sem registros de dados.")
    return linhas[0], linhas[1:], codificacao, dialeto.delimiter


def _converter(valor, cabecalho):
    if valor == "":
        return None
    if cabecalho in COLUNAS_TEXTO:
        return valor
    if cabecalho in COLUNAS_DATA:
        try:
            return datetime.strptime(valor, "%d/%m/%Y")
        except ValueError as erro:
            raise AtualizacaoIndicadoresError(
                "Conversao", f"Data invalida em {cabecalho}: {valor!r}"
            ) from erro
    numero = valor.strip().replace(".", "").replace(",", ".")
    if re.fullmatch(r"[-+]?\d+", numero):
        return int(numero)
    if re.fullmatch(r"[-+]?(?:\d+\.\d*|\d*\.\d+)", numero):
        return float(numero)
    return valor


def _validar_cabecalhos(cabecalho_csv, cabecalhos_xlsx):
    for posicao, (csv_nome, xlsx_nome) in enumerate(
        zip(cabecalho_csv, cabecalhos_xlsx), start=1
    ):
        if csv_nome == xlsx_nome:
            continue
        if (
            not csv_nome
            and CABECALHOS_VAZIOS_PERMITIDOS.get(posicao) == xlsx_nome
        ):
            continue
        raise AtualizacaoIndicadoresError(
            "CSV",
            f"Cabecalho divergente na coluna {posicao}: "
            f"CSV={csv_nome!r}; XLSX={xlsx_nome!r}.",
        )
    return cabecalhos_xlsx


def _traduzir(modelo, origem, destino):
    if isinstance(modelo, str) and modelo.startswith("="):
        try:
            return Translator(modelo, origin=origem).translate_formula(destino)
        except Exception as erro:
            raise AtualizacaoIndicadoresError(
                "Formulas", f"Falha ao traduzir {origem} para {destino}: {modelo}"
            ) from erro
    return modelo


def _validar_planilha(ws, cabecalhos, modelo, registros, ultima, antiga_ultima):
    atuais = tuple(ws.cell(1, c).value for c in range(1, COLUNAS_DADOS + 1))
    if atuais != cabecalhos:
        raise AtualizacaoIndicadoresError("Validacao", "Cabecalhos foram alterados.")
    if ws.cell(2, 1).value != registros[0][0] or ws.cell(2, 50).value != registros[0][-1]:
        raise AtualizacaoIndicadoresError("Validacao", "Primeiro registro nao confere.")
    if ws.cell(ultima, 1).value != registros[-1][0]:
        raise AtualizacaoIndicadoresError("Validacao", "Ultimo registro nao confere.")
    for coluna, valor in enumerate(modelo, start=COLUNA_MODELO_INICIAL):
        origem = f"{get_column_letter(coluna)}2"
        if ws.cell(2, coluna).value != valor:
            raise AtualizacaoIndicadoresError("Validacao", f"Modelo alterado em {origem}.")
        destino = ws.cell(ultima, coluna)
        coordenada_destino = f"{get_column_letter(coluna)}{ultima}"
        if destino.value != _traduzir(valor, origem, coordenada_destino):
            raise AtualizacaoIndicadoresError("Validacao", f"Modelo incorreto em {coordenada_destino}.")
    for linha in range(ultima + 1, antiga_ultima + 1):
        for coluna in range(1, COLUNA_MODELO_FINAL + 1):
            if ws.cell(linha, coluna).value is not None:
                raise AtualizacaoIndicadoresError(
                    "Validacao", f"Conteudo antigo em {ws.cell(linha, coluna).coordinate}."
                )


def _validar_salvo(caminho, cabecalhos, modelo, registros, ultima, antiga_ultima):
    if not caminho.exists() or caminho.stat().st_size <= 0:
        raise AtualizacaoIndicadoresError("Pos-salvamento", "XLSX ausente ou vazio.")
    wb = None
    try:
        wb = load_workbook(caminho, read_only=True, data_only=False, keep_links=True)
        if ABA_DADOS not in wb.sheetnames:
            raise AtualizacaoIndicadoresError("Pos-salvamento", f"Aba {ABA_DADOS!r} ausente.")
        ws = wb[ABA_DADOS]
        linhas_chave = {}
        limite = max(ultima, antiga_ultima)
        for numero, valores in enumerate(
            ws.iter_rows(
                min_row=1,
                max_row=limite,
                min_col=1,
                max_col=COLUNA_MODELO_FINAL,
                values_only=True,
            ),
            start=1,
        ):
            if numero in (1, 2, ultima):
                linhas_chave[numero] = valores
            if numero > ultima and any(valor is not None for valor in valores):
                raise AtualizacaoIndicadoresError(
                    "Pos-salvamento", f"Conteudo antigo permaneceu na linha {numero}."
                )
        if tuple(linhas_chave[1][:50]) != cabecalhos:
            raise AtualizacaoIndicadoresError("Pos-salvamento", "Cabecalhos alterados.")
        if (
            linhas_chave[2][0] != registros[0][0]
            or linhas_chave[2][49] != registros[0][-1]
            or linhas_chave[ultima][0] != registros[-1][0]
        ):
            raise AtualizacaoIndicadoresError("Pos-salvamento", "Dados gravados nao conferem.")
        if tuple(linhas_chave[2][50:91]) != modelo:
            raise AtualizacaoIndicadoresError("Pos-salvamento", "Modelo AY2:CM2 foi alterado.")
        for coluna, valor in enumerate(modelo, start=51):
            origem = f"{get_column_letter(coluna)}2"
            destino = f"{get_column_letter(coluna)}{ultima}"
            if linhas_chave[ultima][coluna - 1] != _traduzir(valor, origem, destino):
                raise AtualizacaoIndicadoresError(
                    "Pos-salvamento", f"Formula final incorreta em {destino}."
                )
    finally:
        if wb is not None:
            wb.close()


def atualizar_indicadores_fpd1(csv_atual, xlsx_destino, log=print):
    inicio = time.monotonic()
    csv_atual = _validar_csv_estavel(csv_atual)
    xlsx_destino = Path(xlsx_destino)
    if not xlsx_destino.exists():
        raise AtualizacaoIndicadoresError("XLSX", f"Destino nao encontrado: {xlsx_destino}")
    cabecalho_csv, linhas_csv, codificacao, delimitador = _ler_csv(csv_atual)
    _log(log, f"CSV utilizado: {csv_atual}")
    _log(log, f"Codificacao={codificacao}; delimitador={delimitador!r}.")
    _log(log, f"Total de linhas do CSV: {len(linhas_csv) + 1}.")

    wb = None
    temporario = xlsx_destino.with_name(f".{xlsx_destino.stem}.{uuid.uuid4().hex}.tmp.xlsx")
    backup = xlsx_destino.with_name(f".{xlsx_destino.stem}.{uuid.uuid4().hex}.bak.xlsx")
    precisa_restaurar = False
    try:
        wb = load_workbook(xlsx_destino, data_only=False, keep_links=True)
        if ABA_DADOS not in wb.sheetnames:
            raise AtualizacaoIndicadoresError("XLSX", f"Aba {ABA_DADOS!r} ausente.")
        ws = wb[ABA_DADOS]
        antiga_ultima = ws.max_row
        cabecalhos = tuple(ws.cell(1, c).value for c in range(1, 51))
        cabecalhos_efetivos = _validar_cabecalhos(cabecalho_csv, cabecalhos)
        modelo = tuple(ws.cell(2, c).value for c in range(51, 92))
        formulas = sum(isinstance(v, str) and v.startswith("=") for v in modelo)
        if not formulas:
            raise AtualizacaoIndicadoresError("Formulas", "AY2:CM2 sem formulas-modelo.")
        registros = [
            tuple(_converter(v, cabecalhos_efetivos[i]) for i, v in enumerate(linha))
            for linha in linhas_csv
        ]
        ultima = len(registros) + 1
        limite = max(antiga_ultima, ultima)
        estilos_modelo = [copy(ws.cell(2, c)._style) for c in range(1, 92)]

        for linha in ws.iter_rows(min_row=2, max_row=limite, min_col=1, max_col=91):
            for celula in linha:
                celula.value = None
        for numero_linha, registro in enumerate(registros, start=2):
            for coluna, valor in enumerate(registro, start=1):
                celula = ws.cell(numero_linha, coluna, valor)
                if numero_linha > antiga_ultima:
                    celula._style = copy(estilos_modelo[coluna - 1])
            for coluna, valor in enumerate(modelo, start=51):
                celula = ws.cell(numero_linha, coluna)
                origem = ws.cell(2, coluna).coordinate
                celula.value = _traduzir(valor, origem, celula.coordinate)
                if numero_linha > antiga_ultima:
                    celula._style = copy(estilos_modelo[coluna - 1])

        _validar_planilha(ws, cabecalhos, modelo, registros, ultima, antiga_ultima)
        wb.calculation.fullCalcOnLoad = True
        wb.calculation.forceFullCalc = True
        wb.calculation.calcMode = "auto"
        wb.save(temporario)
        wb.close()
        wb = None
        _validar_salvo(temporario, cabecalhos, modelo, registros, ultima, antiga_ultima)
        shutil.copy2(xlsx_destino, backup)
        os.replace(temporario, xlsx_destino)
        precisa_restaurar = True
        _validar_salvo(xlsx_destino, cabecalhos, modelo, registros, ultima, antiga_ultima)
        precisa_restaurar = False
        backup.unlink(missing_ok=True)

        resultado = ResultadoAtualizacaoIndicadores(
            csv_atual, xlsx_destino, len(linhas_csv) + 1, len(registros), 2,
            ultima, 50, 41, formulas,
        )
        _log(log, f"Registros importados: {len(registros)}.")
        _log(log, f"Ultima linha escrita: {ultima}.")
        _log(log, "Colunas importadas: 50 (A:AX).")
        _log(log, f"Colunas-modelo: 41 (AY:CM); formulas-modelo: {formulas}.")
        primeira_formula = next(v for v in modelo if isinstance(v, str) and v.startswith("="))
        _log(log, f"Primeira formula: {primeira_formula}")
        _log(log, f"Ultima linha com formulas: {ultima}.")
        _log(log, f"Duracao da atualizacao: {time.monotonic() - inicio:.1f}s.")
        _log(log, "ATUALIZACAO INDICADORES CONCLUIDA COM SUCESSO")
        return resultado
    except AtualizacaoIndicadoresError:
        raise
    except Exception as erro:
        raise AtualizacaoIndicadoresError("Atualizacao", str(erro)) from erro
    finally:
        if wb is not None:
            wb.close()
        temporario.unlink(missing_ok=True)
        if precisa_restaurar and backup.exists():
            os.replace(backup, xlsx_destino)
        else:
            backup.unlink(missing_ok=True)
