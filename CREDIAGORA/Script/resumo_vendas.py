"""Totais mensais do Panorama registrados apos salvar cada planilha."""

import configparser
import os
import re
import tempfile
import unicodedata
from dataclasses import dataclass
from datetime import datetime
from decimal import ROUND_HALF_UP, Decimal, InvalidOperation
from pathlib import Path

from receita_cpc import converter_data_cpc

CAMPOS = {
    "VENDAS": ("VALOR CONTRATO", "Soma do valor dos contratos"),
    "RECEITA": ("COMISSAO PREVISTA", "Soma da comissao prevista"),
}


@dataclass(frozen=True)
class ResumoMensal:
    tipo: str
    competencia: str
    registros: int
    total: Decimal


def _normalizar(texto):
    texto = unicodedata.normalize("NFKD", str(texto or ""))
    return " ".join("".join(c for c in texto if not unicodedata.combining(c)).upper().split())


def _indice(cabecalhos, nome):
    indices = [i for i, valor in enumerate(cabecalhos) if _normalizar(valor) == nome]
    if len(indices) != 1:
        raise ValueError(f"Resumo: esperado um unico cabecalho {nome!r}.")
    return indices[0]


def _valor_monetario(valor):
    if valor is None or isinstance(valor, str) and not valor.strip():
        return Decimal(0)
    if isinstance(valor, bool):
        raise TypeError("Valor monetario booleano invalido.")
    texto = str(valor).strip()
    if isinstance(valor, str):
        texto = texto.removeprefix("R$").strip()
        if "," in texto:
            if not re.fullmatch(r"[+-]?(?:\d+|\d{1,3}(?:\.\d{3})+),\d+", texto):
                raise ValueError("Valor monetario brasileiro invalido.")
            texto = texto.replace(".", "").replace(",", ".")
        elif re.fullmatch(r"[+-]?\d{1,3}(?:\.\d{3})+", texto):
            texto = texto.replace(".", "")
    try:
        numero = Decimal(texto)
    except InvalidOperation as erro:
        raise ValueError("Valor monetario invalido.") from erro
    if not numero.is_finite():
        raise ValueError("Valor monetario nao finito.")
    return numero


def calcular_resumo(linhas, cabecalhos, tipo, referencia):
    if tipo not in CAMPOS:
        raise ValueError(f"Tipo de resumo desconhecido: {tipo}")
    coluna_valor = _indice(cabecalhos, CAMPOS[tipo][0])
    coluna_data = _indice(cabecalhos, "DATA CPC")
    total = Decimal(0)
    registros = 0
    for numero, linha in enumerate(linhas, start=2):
        if len(linha) <= max(coluna_data, coluna_valor):
            raise ValueError(f"Resumo {tipo}: colunas ausentes na linha {numero}.")
        data = converter_data_cpc(linha[coluna_data])
        if data is None:
            raise ValueError(f"Resumo {tipo}: DATA CPC ausente na linha {numero}.")
        if (data.year, data.month) != (referencia.year, referencia.month):
            continue
        try:
            total += _valor_monetario(linha[coluna_valor])
        except (TypeError, ValueError) as erro:
            raise ValueError(f"Resumo {tipo}: valor monetario invalido na linha {numero}.") from erro
        registros += 1
    return ResumoMensal(tipo, referencia.strftime("%m/%Y"), registros, total)


def salvar_resumo(pasta, resumo, planilha, atualizado_em=None):
    atualizado_em = atualizado_em or datetime.now().astimezone()
    pasta = Path(pasta)
    pasta.mkdir(parents=True, exist_ok=True)
    destino = pasta / f"{atualizado_em:%d-%m-%Y}.txt"
    documento = configparser.ConfigParser(interpolation=None)
    documento.optionxform = lambda optionstr: optionstr
    if destino.exists():
        with destino.open(encoding="utf-8") as arquivo:
            documento.read_file(arquivo)
    else:
        documento["RESUMO DIARIO"] = {"Data": atualizado_em.strftime("%d/%m/%Y")}
        for tipo in CAMPOS:
            documento[tipo] = {"Status": "Aguardando atualizacao"}
    valor = f"{resumo.total.quantize(Decimal('0.01'), rounding=ROUND_HALF_UP):,.2f}"
    valor = valor.translate(str.maketrans({",": ".", ".": ","}))
    documento[resumo.tipo] = {
        "Status": "Atualizacao concluida",
        "Atualizado em": atualizado_em.strftime("%d/%m/%Y %H:%M:%S %z"),
        "Periodo dos registros (DATA CPC)": resumo.competencia,
        "Registros exportados do mes": str(resumo.registros),
        CAMPOS[resumo.tipo][1]: f"R$ {valor}",
        "Planilha atualizada": str(planilha),
    }
    temporario = None
    try:
        with tempfile.NamedTemporaryFile(mode="w", encoding="utf-8", dir=pasta, suffix=".tmp", delete=False) as arquivo:
            temporario = Path(arquivo.name)
            documento.write(arquivo)
        os.replace(temporario, destino)
    finally:
        if temporario is not None:
            temporario.unlink(missing_ok=True)
    return destino

def ler_resumo(pasta, data):
    pasta = Path(pasta)
    destino = pasta / f"{data:%d-%m-%Y}.txt"
    if not destino.exists():
        raise FileNotFoundError(f"Resumo diario nao encontrado: {destino}")
    documento = configparser.ConfigParser(interpolation=None)
    documento.optionxform = lambda optionstr: optionstr
    with destino.open(encoding="utf-8") as arquivo:
        documento.read_file(arquivo)
    resumo_diario = {}
    for tipo in CAMPOS:
        if tipo not in documento:
            raise ValueError(f"Resumo diario corrompido: secao {tipo} ausente.")
        secao = documento[tipo]
        try:
            competencia = secao["Periodo dos registros (DATA CPC)"]
            registros = int(secao["Registros exportados do mes"])
            total = _valor_monetario(secao[CAMPOS[tipo][1]])
        except KeyError as erro:
            raise ValueError(f"Resumo diario corrompido: campo ausente na secao {tipo}.") from erro
        resumo_diario[tipo] = ResumoMensal(tipo, competencia, registros, total)
    return resumo_diario

