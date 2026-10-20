"""ImportaÃ§Ã£o integral de contato campanha Normal, sem converter campos do CSV."""

import csv
import hashlib
import io
import os
import shutil
import time
import unicodedata
import uuid
import warnings
import zipfile
from copy import copy
from datetime import datetime
from difflib import SequenceMatcher
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.formula.translate import Translator
from openpyxl.styles.numbers import is_date_format
from openpyxl.utils import get_column_letter, range_boundaries


class RetornoLigacoesError(RuntimeError):
    def __init__(self, etapa, mensagem):
        super().__init__(f"[PANORAMA][RETORNO LIGACOES][ETAPA {etapa}] {mensagem}")


def registrar(log, etapa, mensagem):
    log(f"[PANORAMA][RETORNO LIGACOES][ETAPA {etapa}] {mensagem}")


def inventariar_downloads(pasta):
    pasta = Path(pasta)
    pasta.mkdir(parents=True, exist_ok=True)
    return {p.name for p in pasta.iterdir() if p.is_file()}


def esperar_zip_novo(pasta, anteriores, timeout, verificar_sessao=lambda: None):
    """Exige nome novo, tamanho estÃ¡vel e ZIP completo; nunca aceita arquivo antigo."""
    pasta = Path(pasta)
    fim = time.monotonic() + timeout
    estados = {}
    while time.monotonic() < fim:
        verificar_sessao()
        arquivos = [p for p in pasta.iterdir() if p.is_file() and p.name not in anteriores]
        # Browser.setDownloadBehavior(allowAndName) usa um UUID atÃ© concluirmos a validaÃ§Ã£o.
        def candidato(p):
            if p.suffix.lower() == ".zip":
                return True
            try:
                uuid.UUID(p.name)
                return True
            except ValueError:
                return False

        zips = [p for p in arquivos if candidato(p)]
        if len(zips) > 1:
            raise RetornoLigacoesError(2, "Mais de um ZIP novo; nÃ£o Ã© seguro escolher a exportaÃ§Ã£o.")
        if zips:
            caminho = zips[0]
            estado = (caminho.stat().st_size, caminho.stat().st_mtime_ns)
            pendentes = any(p.name.lower().startswith(caminho.name.lower() + ".")
                            and p.suffix.lower() in (".crdownload", ".part", ".tmp", ".download")
                            for p in arquivos)
            if estado[0] and estados.get(caminho.name) == estado and not pendentes:
                validar_zip(caminho)
                return caminho
            estados[caminho.name] = estado
        time.sleep(0.5)
    raise RetornoLigacoesError(2, f"ZIP novo nÃ£o concluÃ­do em {timeout}s; arquivos preservados.")


def validar_zip(caminho, etapa=2):
    try:
        with zipfile.ZipFile(caminho) as arquivo:
            if arquivo.testzip() is not None:
                raise RetornoLigacoesError(etapa, "ZIP corrompido (CRC invÃ¡lido).")
    except (zipfile.BadZipFile, OSError, RuntimeError) as erro:
        raise RetornoLigacoesError(etapa, f"NÃ£o foi possÃ­vel validar o ZIP: {erro}") from erro


def ler_csv_zip(caminho, log=print):
    validar_zip(caminho)
    with zipfile.ZipFile(caminho) as arquivo:
        candidatos = [i for i in arquivo.infolist()
                      if not i.is_dir() and i.filename.lower().endswith(".csv")]
        if len(candidatos) != 1:
            raise RetornoLigacoesError(2, f"Esperado um CSV no ZIP; encontrados {len(candidatos)}.")
        dados = arquivo.read(candidatos[0])  # Leitura em memÃ³ria evita caminhos de extraÃ§Ã£o inseguros.
        registrar(log, 2, f"CSV identificado: {candidatos[0].filename}.")
    if not dados:
        raise RetornoLigacoesError(3, "CSV vazio.")
    codificacoes = ("utf-16",) if dados.startswith((b"\xff\xfe", b"\xfe\xff")) else ("utf-8-sig", "cp1252")
    texto = None
    codificacao = None
    dialeto = None
    for codificacao in codificacoes:
        try:
            texto = dados.decode(codificacao)
            break
        except UnicodeError:
            continue
    if texto is None or "\x00" in texto:
        raise RetornoLigacoesError(3, "CodificaÃ§Ã£o invÃ¡lida ou conteÃºdo danificado.")
    if "\ufffd" in texto:
        registrar(log, 3, f"Aviso: o CSV original contÃ©m {texto.count(chr(65533))} caracteres de substituiÃ§Ã£o; conteÃºdo preservado sem tentar corrigir nomes.")
    try:
        # O cabeÃ§alho determina o separador, sem depender de campos livres nas linhas seguintes.
        dialeto = csv.Sniffer().sniff(texto.splitlines()[0], delimiters=";,\t|")
        leitor = csv.reader(io.StringIO(texto, newline=""), dialeto, strict=True)
        cabecalho = next(leitor)
        if len(cabecalho) < 9 or any(not c.strip() for c in cabecalho[:9]):
            raise RetornoLigacoesError(3, "CabeÃ§alho deve conter pelo menos nove colunas identificadas.")
        linhas = []
        linhas_com_colunas_extras = []
        for linha in leitor:
            if not linha or not any(c.strip() for c in linha):
                continue
            if len(linha) < len(cabecalho):
                raise RetornoLigacoesError(
                    3,
                    f"Registro na linha {leitor.line_num} possui somente {len(linha)} "
                    f"colunas; eram esperadas {len(cabecalho)}.",
                )
            if len(linha) > len(cabecalho):
                # O Panorama pode acrescentar telefones após a coluna I. A planilha
                # possui contrato fixo A:I; os excedentes não podem deslocar dados.
                linhas_com_colunas_extras.append((leitor.line_num, len(linha)))
            valores = tuple(linha[:9])
            if not any(c.strip() for c in valores):
                raise RetornoLigacoesError(3, f"Registro sem dados em A:I na linha {leitor.line_num}.")
            if any(len(v) > 32767 for v in valores):
                raise RetornoLigacoesError(3, "Campo excede o limite de caracteres do Excel.")
            linhas.append(valores)
    except (csv.Error, StopIteration, IndexError) as erro:
        raise RetornoLigacoesError(3, f"CSV fora do formato esperado: {erro}") from erro
    if not linhas or len(linhas) > 1048575:
        raise RetornoLigacoesError(3, "CSV sem registros ou acima do limite de linhas do Excel.")
    if codificacao is None or dialeto is None:
        raise RetornoLigacoesError(3, "NÃ£o foi possÃ­vel identificar a codificaÃ§Ã£o ou o delimitador do CSV.")
    if linhas_com_colunas_extras:
        exemplo_linha, exemplo_colunas = linhas_com_colunas_extras[0]
        registrar(
            log,
            3,
            f"Registros com telefones excedentes apÃ³s I: {len(linhas_com_colunas_extras)}; "
            f"primeiro na linha {exemplo_linha} ({exemplo_colunas} colunas).",
        )
    registrar(log, 3, f"CodificaÃ§Ã£o={codificacao}; delimitador={dialeto.delimiter!r}; colunas={len(cabecalho)}; registros={len(linhas)}.")
    return tuple(cabecalho[:9]), linhas


def normalizar_cabecalho(valor):
    texto = unicodedata.normalize("NFKD", str(valor or "").strip().casefold())
    texto = "".join(c for c in texto if c.isalnum())
    # O CSV real usa 'datacontato'; o modelo fornecido chama a mesma coluna E de 'Data_Contrato'.
    return "datacontato" if texto == "datacontrato" else texto


def chave_status(valor):
    """Compara status ignorando acentos e caracteres de substituiÃ§Ã£o do CSV."""
    texto = str(valor or "")
    # Corrige acentos em Latin-1 e mojibake UTF-8 interpretado como Latin-1.
    try:
        reparado = texto.encode("latin-1").decode("utf-8")
        if reparado:
            texto = reparado
    except (UnicodeEncodeError, UnicodeDecodeError):
        pass
    texto = texto.replace("ï¿½", "").replace("�", "")
    texto = unicodedata.normalize("NFKD", texto)
    texto = "".join(c for c in texto if unicodedata.category(c) != "Mn")
    return "".join(c.casefold() for c in texto if c.isalnum())


def canonicalizar_status(linhas, ws, log=print):
    validos = [ws.cell(linha, 1).value for linha in range(2, ws.max_row + 1)]
    validos = [str(v) for v in validos if v not in (None, "")]
    por_chave = {chave_status(v): v for v in validos}
    corrigidos = 0
    resultado = []
    for registro in linhas:
        atual = registro[5]
        canonico = por_chave.get(chave_status(atual))
        if canonico is None and "ï¿½" in str(atual):
            candidatos = sorted(
                ((SequenceMatcher(None, chave_status(atual), chave_status(v)).ratio(), v) for v in validos),
                reverse=True,
            )
            if candidatos and candidatos[0][0] >= 0.50:
                canonico = candidatos[0][1]
        if canonico is not None and canonico != atual:
            registro = registro[:5] + (canonico,) + registro[6:]
            corrigidos += 1
        resultado.append(registro)
    if corrigidos:
        registrar(log, 3, f"Status normalizados pela aba validador: {corrigidos}.")
    return resultado


def identificar_aba(wb, cabecalho):
    esperado = tuple(map(normalizar_cabecalho, cabecalho))
    candidatas = [ws for ws in wb if tuple(normalizar_cabecalho(ws.cell(1, c).value)
                                          for c in range(1, 10)) == esperado]
    if len(candidatas) != 1:
        raise RetornoLigacoesError(4, f"NÃ£o foi possÃ­vel identificar uma aba Ãºnica pelos cabeÃ§alhos A:I (encontradas {len(candidatas)}).")
    return candidatas[0]


def modelos_formulas(ws):
    modelos = {}
    for coluna in range(10, 13):
        for linha in range(2, ws.max_row + 1):
            celula = ws.cell(linha, coluna)
            if celula.data_type == "f" and isinstance(celula.value, str):
                modelos[coluna] = (celula.coordinate, celula.value)
                break
        if coluna not in modelos:
            raise RetornoLigacoesError(6, f"FÃ³rmula-modelo ausente na coluna {get_column_letter(coluna)}.")
    # Usa funÃ§Ãµes compatÃ­veis com versÃµes do Excel que nÃ£o calculam XLOOKUP/LET.
    modelos[10] = ("J2", '=IFERROR(INDEX(validador!$B$2:$B$13,MATCH(F2,validador!$A$2:$A$13,0)),"")')
    modelos[11] = ("K2", '=IFERROR(INDEX(validador!$C$2:$C$13,MATCH(J2,validador!$B$2:$B$13,0)),"")')
    return modelos


def _hash_arquivo(caminho):
    with open(caminho, "rb") as arquivo:
        return hashlib.file_digest(arquivo, "sha256").hexdigest()


def _preservados(wb, aba):
    return {(ws.title, c.coordinate): (c.value, c.data_type, copy(c._style))
            for ws in wb for c in ws._cells.values()
            if (ws.title != aba or c.row == 1 or c.column > 12)
            and (c.value is not None or c.has_style)}


def _validar_resultado(wb, aba, linhas, modelos, antiga_ultima, preservados):
    ws = wb[aba]
    for numero, registro in enumerate(linhas, 2):
        for coluna, valor in enumerate(registro, 1):
            celula = ws.cell(numero, coluna)
            if (celula.value or "") != (valor or "") or celula.data_type == "f":
                raise RetornoLigacoesError(7, f"Valor divergente em {celula.coordinate}.")
        for coluna, (origem, formula) in modelos.items():
            celula = ws.cell(numero, coluna)
            if celula.value != Translator(formula, origin=origem).translate_formula(celula.coordinate):
                raise RetornoLigacoesError(7, f"FÃ³rmula divergente em {celula.coordinate}.")
    for linha in ws.iter_rows(min_row=len(linhas) + 2, max_row=antiga_ultima, max_col=12):
        if any(c.value is not None for c in linha):
            raise RetornoLigacoesError(7, "Restaram dados ou fÃ³rmulas excedentes.")
    if _preservados(wb, aba) != preservados:
        raise RetornoLigacoesError(7, "CabeÃ§alhos, outras colunas ou abas sofreram alteraÃ§Ã£o inesperada.")


def atualizar_retorno_ligacoes(zip_atual, destino, log=print, aguardar_arquivo=None):
    etapa = 3
    wb = None
    destino = Path(destino)
    temporario = destino.with_name(f".{destino.stem}.{uuid.uuid4().hex}.tmp.xlsx")
    try:
        registrar(log, 3, "Lendo dados do CSV...")
        cabecalho, linhas = ler_csv_zip(zip_atual, log)
        etapa = 4
        if not destino.is_file():
            raise RetornoLigacoesError(4, f"Excel nÃ£o encontrado: {destino}")
        if aguardar_arquivo:
            aguardar_arquivo(destino)
        original = _hash_arquivo(destino)
        registrar(log, 4, "Abrindo Excel e identificando a aba pelos cabeÃ§alhos...")
        with warnings.catch_warnings():
            warnings.simplefilter("error", UserWarning)
            wb = load_workbook(destino, data_only=False, keep_links=False)
        ws = identificar_aba(wb, cabecalho)
        linhas = canonicalizar_status(linhas, wb["validador"], log)
        modelos = modelos_formulas(ws)
        if is_date_format(ws.cell(2, 5).number_format):
            # MantÃ©m o tipo de data da planilha: YEAR/MONTH em L exigem datas vÃ¡lidas.
            registros = []
            for numero, registro in enumerate(linhas, 2):
                data = registro[4]
                try:
                    valor_data = datetime.fromisoformat(data) if data else None
                    if valor_data is not None and valor_data.tzinfo is not None:
                        raise ValueError("fuso horÃ¡rio nÃ£o suportado no Excel")
                except ValueError as erro:
                    raise RetornoLigacoesError(3, f"Data invÃ¡lida na coluna E, linha {numero}.") from erro
                registros.append(registro[:4] + (valor_data,) + registro[5:])
            linhas = registros
            registrar(log, 3, "Coluna E mantida como data do Excel para as fÃ³rmulas YEAR/MONTH; demais campos preservados como texto.")
        antiga_ultima = ws.max_row
        ultima = len(linhas) + 1
        if any(r.min_col <= 12 and r.max_row >= 2 for r in ws.merged_cells.ranges):
            raise RetornoLigacoesError(4, "HÃ¡ cÃ©lulas mescladas no intervalo da base.")
        preservados = _preservados(wb, ws.title)
        estilos = {c: copy(ws.cell(2, c)._style) for c in range(1, 13)}
        registrar(log, 4, f"Aba de destino: {ws.title}; Ãºltima linha anterior: {antiga_ultima}.")
        etapa = 5
        registrar(log, 5, "Substituindo valores antigos de A:I e removendo excedentes de J:L...")
        for linha in ws.iter_rows(min_row=2, max_row=antiga_ultima, max_col=12):
            for celula in linha:
                celula.value = None
        for numero, registro in enumerate(linhas, 2):
            for coluna, valor in enumerate(registro, 1):
                celula = ws.cell(numero, coluna)
                celula.value = valor
                if isinstance(valor, str):
                    celula.data_type = "s"  # Inclusive textos comeÃ§ando com '=': nunca fÃ³rmula do CSV.
            if numero > antiga_ultima:
                for coluna in range(1, 13):
                    ws.cell(numero, coluna)._style = copy(estilos[coluna])
        registrar(log, 5, f"Novos registros inseridos: {len(linhas)}.")
        etapa = 6
        for coluna, (origem, formula) in modelos.items():
            tradutor = Translator(formula, origin=origem)
            for numero in range(2, ultima + 1):
                celula = ws.cell(numero, coluna)
                celula.value = tradutor.translate_formula(celula.coordinate)
        registrar(log, 6, f"FÃ³rmulas originais preenchidas em J2:L{ultima}.")
        for tabela in ws.tables.values():
            referencia = tabela.ref
            if not referencia:
                continue
            c1, l1, c2, _ = range_boundaries(referencia)
            if c1 is None or l1 is None or c2 is None:
                continue
            coluna_inicial, linha_inicial, coluna_final = int(c1), int(l1), int(c2)
            if linha_inicial == 1 and coluna_inicial == 1 and coluna_final <= 12:
                if tabela.totalsRowCount:
                    raise RetornoLigacoesError(6, "Tabela com totalizadores exige configuraÃ§Ã£o especÃ­fica.")
                tabela.ref = f"A1:{get_column_letter(coluna_final)}{ultima}"
                if tabela.autoFilter:
                    tabela.autoFilter.ref = tabela.ref
        filtro = ws.auto_filter.ref
        if filtro:
            c1, l1, c2, _ = range_boundaries(filtro)
            if c1 is not None and l1 is not None and c2 is not None:
                coluna_inicial, linha_inicial, coluna_final = int(c1), int(l1), int(c2)
                if linha_inicial == 1 and coluna_inicial == 1 and coluna_final <= 12:
                    ws.auto_filter.ref = f"A1:{get_column_letter(coluna_final)}{ultima}"
        etapa = 7
        wb.calculation.fullCalcOnLoad = True
        wb.calculation.forceFullCalc = True
        wb.calculation.calcMode = "auto"
        registrar(log, 7, "Salvando arquivo temporÃ¡rio e conferindo todos os registros e fÃ³rmulas...")
        aba = ws.title
        if getattr(wb, "_external_links", None):
            wb._external_links.clear()
        wb.save(temporario)
        wb.close()
        wb = None
        validar_zip(temporario, etapa=7)
        wb = load_workbook(temporario, data_only=False, keep_links=False)
        _validar_resultado(wb, aba, linhas, modelos, antiga_ultima, preservados)
        wb.close()
        wb = None
        if aguardar_arquivo:
            aguardar_arquivo(destino)
        if _hash_arquivo(destino) != original:
            raise RetornoLigacoesError(7, "Excel alterado por outro processo durante a importaÃ§Ã£o; substituiÃ§Ã£o cancelada.")
        try:
            os.replace(temporario, destino)
        except PermissionError:
            # OneDrive pode manter o arquivo de destino aberto durante a
            # sincronização. Como o arquivo já foi validado, grava por cópia
            # mantendo o mesmo caminho e tenta remover o temporário depois.
            shutil.copy2(temporario, destino)
            try:
                os.remove(temporario)
            except OSError:
                pass
        registrar(log, 7, f"Excel validado e salvo no caminho original. Registros={len(linhas)}; Ãºltima linha={ultima}.")
        return destino
    except RetornoLigacoesError:
        raise
    except Exception as erro:
        raise RetornoLigacoesError(etapa, str(erro)) from erro
    finally:
        if wb is not None:
            wb.close()
        temporario.unlink(missing_ok=True)
