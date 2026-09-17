import contextlib
import importlib.util
import io
import sys
import tempfile
import time
import unittest
from pathlib import Path


RAIZ_PROJETO = Path(__file__).resolve().parents[1]
START_PY = RAIZ_PROJETO / "CREDIAGORA" / "Script" / "start.py"
SCRIPT_DIR = START_PY.parent

if str(SCRIPT_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPT_DIR))


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

    def test_acompanhamento_de_operacao_exibe_progresso(self):
        modulo, _ = carregar_start()
        modulo.LOGS_ATIVACAO.clear()

        with modulo.acompanhar_operacao("Operação de teste", intervalo=0.01):
            time.sleep(0.03)

        texto = "\n".join(modulo.LOGS_ATIVACAO)
        self.assertIn("Operação de teste em andamento", texto)
        self.assertIn("Operação de teste concluído", texto)

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

    def test_execucao_completa_ignora_ercard_sem_credenciais(self):
        modulo, _ = carregar_start()
        config = modulo.ErCardConfig("", "", "", "", Path("."), Path("."))

        self.assertFalse(modulo.ativar_ercard_se_configurado(config))
        texto = "\n".join(modulo.LOGS_ATIVACAO)
        self.assertIn("Fase ERCard ignorada", texto)

    def test_somente_ercard_exige_credenciais_com_instrucao(self):
        modulo, _ = carregar_start()
        config = modulo.ErCardConfig("", "", "", "", Path("."), Path("."))

        with self.assertRaisesRegex(EnvironmentError, "configurar_ercard.bat"):
            modulo.ativar_ercard_se_configurado(config, somente_ercard=True)

    def test_inicio_chrome_repete_apos_falha_do_perfil(self):
        modulo, _ = carregar_start()
        driver = unittest.mock.MagicMock()

        with unittest.mock.patch.object(
            modulo.webdriver,
            "Chrome",
            side_effect=[RuntimeError("perfil ocupado"), driver],
        ) as iniciar, unittest.mock.patch.object(
            modulo, "encerrar_chrome_da_automacao", return_value=1
        ), unittest.mock.patch.object(
            modulo, "limpar_travas_perfil_chrome"
        ), unittest.mock.patch.object(modulo.time, "sleep"):
            resultado = modulo.iniciar_chrome()

        self.assertIs(resultado, driver)
        self.assertEqual(iniciar.call_count, 2)


if __name__ == "__main__":
    unittest.main()
