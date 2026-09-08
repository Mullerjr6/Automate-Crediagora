import importlib.util
import tempfile
import unittest
from datetime import datetime
from pathlib import Path
from unittest.mock import patch


RAIZ_PROJETO = Path(__file__).resolve().parents[1]
ERCARD_PY = RAIZ_PROJETO / "CREDIAGORA" / "Script" / "ercard.py"


def carregar_ercard():
    spec = importlib.util.spec_from_file_location("ercard_testes", ERCARD_PY)
    modulo = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(modulo)
    return modulo


class ErCardTests(unittest.TestCase):
    def test_nome_dinamico_sem_caractere_invalido(self):
        modulo = carregar_ercard()
        nome = modulo.gerar_nome_tabela_contratos(datetime(2026, 9, 8, 16, 35))
        self.assertEqual(nome, "Tabela contratos 08-09-2026 16-35.csv")
        self.assertNotIn(":", nome)

    def test_caminho_unico_nao_sobrescreve(self):
        modulo = carregar_ercard()
        with tempfile.TemporaryDirectory() as pasta:
            pasta = Path(pasta)
            (pasta / "Tabela contratos teste.csv").touch()
            caminho = modulo.caminho_exportacao_unico(
                pasta, "Tabela contratos teste.csv"
            )
            self.assertEqual(caminho.name, "Tabela contratos teste (2).csv")

    def test_configuracao_rejeita_credenciais_ausentes(self):
        modulo = carregar_ercard()
        config = modulo.ErCardConfig("", "", "", "", Path("."), Path("."))
        with self.assertRaises(EnvironmentError):
            config.validar()

    def test_valida_csv_estavel_e_legivel(self):
        modulo = carregar_ercard()
        with tempfile.TemporaryDirectory() as pasta:
            pasta = Path(pasta)
            arquivo = pasta / "Tabela contratos teste.csv"
            arquivo.write_text("codigo;contrato\n1;ABC\n", encoding="utf-8")

            with patch.object(modulo.time, "sleep", return_value=None):
                resultado = modulo.validar_csv_exportado(
                    arquivo, pasta, timeout=2, log=lambda _: None
                )

            self.assertEqual(resultado, arquivo)


if __name__ == "__main__":
    unittest.main()
