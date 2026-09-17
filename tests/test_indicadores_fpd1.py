import csv
import importlib.util
import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from openpyxl import Workbook, load_workbook


RAIZ = Path(__file__).resolve().parents[1]
MODULO = RAIZ / "CREDIAGORA" / "Script" / "indicadores_fpd1.py"


def carregar_modulo():
    spec = importlib.util.spec_from_file_location("indicadores_fpd1_testes", MODULO)
    modulo = importlib.util.module_from_spec(spec)
    sys.modules[spec.name] = modulo
    spec.loader.exec_module(modulo)
    return modulo


def criar_xlsx(caminho, total_linhas=5):
    wb = Workbook()
    ws = wb.active
    ws.title = "Tabela Conta"
    cabecalhos = [f"Campo{i}" for i in range(1, 51)]
    cabecalhos[0], cabecalhos[1], cabecalhos[12] = "DataContabil", "Proposta", "Cpf"
    ws.append(cabecalhos)
    for linha in range(2, total_linhas + 1):
        for coluna in range(1, 51):
            ws.cell(linha, coluna).value = f"antigo-{linha}-{coluna}"
        ws.cell(linha, 51).value = f"=A{linha}"
        ws.cell(linha, 52).value = f"=$A$1&A{linha}"
        ws.cell(linha, 53).value = 7
    wb.save(caminho)
    wb.close()
    return cabecalhos


def criar_csv(caminho, cabecalhos, registros):
    with caminho.open("w", encoding="utf-8-sig", newline="") as arquivo:
        escritor = csv.writer(arquivo, delimiter=";")
        escritor.writerow(cabecalhos + [""])
        for registro in registros:
            escritor.writerow(registro + [""])


class IndicadoresFpd1Tests(unittest.TestCase):
    def test_atualiza_formulas_textos_e_remove_sobras(self):
        modulo = carregar_modulo()
        with tempfile.TemporaryDirectory() as pasta:
            pasta = Path(pasta)
            xlsx, csv_atual = pasta / "indicadores.xlsx", pasta / "contratos.csv"
            cabecalhos = criar_xlsx(xlsx)
            registros = []
            for numero in (1, 2):
                linha = [str(numero)] * 50
                linha[0] = f"0{numero}/01/2024"
                linha[1] = f"000{numero}"
                linha[12] = f"0000000000{numero}"
                registros.append(linha)
            criar_csv(csv_atual, cabecalhos, registros)
            with patch.object(modulo.time, "sleep", return_value=None):
                resultado = modulo.atualizar_indicadores_fpd1(
                    csv_atual, xlsx, lambda _mensagem: None
                )
            self.assertEqual((resultado.registros_importados, resultado.ultima_linha), (2, 3))
            wb = load_workbook(xlsx, data_only=False)
            ws = wb["Tabela Conta"]
            self.assertEqual(ws["B2"].value, "0001")
            self.assertEqual(ws["M2"].value, "00000000001")
            self.assertEqual(ws["AY3"].value, "=A3")
            self.assertEqual(ws["AZ3"].value, "=$A$1&A3")
            self.assertEqual(ws["BA3"].value, 7)
            self.assertIsNone(ws["A4"].value)
            self.assertIsNone(ws["AY4"].value)
            wb.close()

    def test_csv_invalido_nao_altera_original(self):
        modulo = carregar_modulo()
        with tempfile.TemporaryDirectory() as pasta:
            pasta = Path(pasta)
            xlsx, csv_atual = pasta / "indicadores.xlsx", pasta / "contratos.csv"
            cabecalhos = criar_xlsx(xlsx)
            criar_csv(csv_atual, cabecalhos, [["1"] * 50])
            with csv_atual.open("a", encoding="utf-8") as arquivo:
                arquivo.write(";EXTRA")
            original = xlsx.read_bytes()
            with patch.object(modulo.time, "sleep", return_value=None):
                with self.assertRaises(modulo.AtualizacaoIndicadoresError):
                    modulo.atualizar_indicadores_fpd1(
                        csv_atual, xlsx, lambda _mensagem: None
                    )
            self.assertEqual(xlsx.read_bytes(), original)


if __name__ == "__main__":
    unittest.main()
