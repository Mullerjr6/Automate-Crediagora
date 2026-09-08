import contextlib
import importlib.util
import io
import tempfile
import unittest
from pathlib import Path


RAIZ_PROJETO = Path(__file__).resolve().parents[1]
START_PY = RAIZ_PROJETO / "CREDIAGORA" / "Script" / "start.py"


def carregar_start():
    spec = importlib.util.spec_from_file_location("start_crediagora", START_PY)
    modulo = importlib.util.module_from_spec(spec)
    saida = io.StringIO()

    with contextlib.redirect_stdout(saida):
        spec.loader.exec_module(modulo)

    return modulo, saida.getvalue()


class StartTests(unittest.TestCase):
    def test_importar_nao_executa_automacao(self):
        _, saida = carregar_start()
        self.assertEqual(saida, "")

    def test_normaliza_texto_com_acentos(self):
        modulo, _ = carregar_start()
        self.assertEqual(modulo.normalizar_texto("Empréstimo"), "emprestimo")
        self.assertEqual(modulo.normalizar_texto("Exportação VENDAS"), "exportacao vendas")

    def test_argumentos_padrao(self):
        modulo, _ = carregar_start()
        args = modulo.parse_args([])
        self.assertFalse(args.check)
        self.assertFalse(args.check_login)
        self.assertEqual(args.exportacao, "todas")

    def test_argumento_check_login(self):
        modulo, _ = carregar_start()
        args = modulo.parse_args(["--check-login"])
        self.assertTrue(args.check_login)

    def test_bloqueio_local_protege_mesma_credencial(self):
        modulo, _ = carregar_start()

        with tempfile.TemporaryDirectory() as pasta:
            modulo.ARQUIVO_ESTADO_LOGIN = Path(pasta) / "estado_login.json"
            modulo.salvar_estado_login(True, "usuário bloqueado pelo portal", 0)

            with self.assertRaises(RuntimeError):
                modulo.validar_bloqueio_local_login()

    def test_bloqueio_local_nao_prende_credencial_nova(self):
        modulo, _ = carregar_start()

        with tempfile.TemporaryDirectory() as pasta:
            modulo.ARQUIVO_ESTADO_LOGIN = Path(pasta) / "estado_login.json"
            senha_original = modulo.SENHA
            modulo.salvar_estado_login(True, "usuário bloqueado pelo portal", 0)
            modulo.SENHA = senha_original + "-nova"

            try:
                modulo.validar_bloqueio_local_login()
            finally:
                modulo.SENHA = senha_original

    def test_validacao_rejeita_senha_vazia(self):
        modulo, _ = carregar_start()
        senha_original = modulo.SENHA
        modulo.SENHA = ""

        try:
            with self.assertRaises(EnvironmentError):
                modulo.validar_configuracao()
        finally:
            modulo.SENHA = senha_original

    def test_limpeza_rejeita_pasta_downloads_ampla(self):
        modulo, _ = carregar_start()
        pasta_original = modulo.PASTA_DOWNLOAD
        modulo.PASTA_DOWNLOAD = Path.home() / "Downloads"

        try:
            with self.assertRaises(ValueError):
                modulo.validar_pasta_download_segura()
        finally:
            modulo.PASTA_DOWNLOAD = pasta_original


if __name__ == "__main__":
    unittest.main()
