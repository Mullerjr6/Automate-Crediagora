import importlib.util
import tempfile
import unittest
from datetime import datetime
from pathlib import Path
from unittest.mock import patch

from selenium.webdriver.common.by import By


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

    def test_login_reconhece_botao_input_do_portal(self):
        modulo = carregar_ercard()

        class Elemento:
            def __init__(self, valor=""):
                self.valor = valor
                self.size = {"width": 100, "height": 30}

            def get_attribute(self, nome):
                return self.valor if nome == "value" else ""

            def clear(self):
                self.valor = ""

            def send_keys(self, valor):
                self.valor += valor

            def click(self):
                return None

            def is_displayed(self):
                return True

        usuario = Elemento()
        senha = Elemento()
        botao = Elemento("Entrar")
        driver = unittest.mock.MagicMock()
        driver.window_handles = ["principal"]
        driver.current_window_handle = "principal"
        driver.current_url = "https://app.ercard.com.br/software/html5.html"

        def elementos(_driver, by, seletor):
            if by == By.XPATH and "@type='text'" in seletor:
                return [("conteúdo principal", usuario)]
            if by == By.CSS_SELECTOR and seletor == "input[type='password']":
                return [("conteúdo principal", senha)]
            return []

        config = modulo.ErCardConfig(
            "usuario", "senha", "sistema", "senha2", Path("."), Path("."),
            timeout_normal=1,
        )

        with patch.object(modulo, "_elementos_em_contextos", side_effect=elementos), \
             patch.object(modulo, "_primeiro_elemento", return_value=botao) as primeiro:
            modulo.realizar_login_portal(driver, config, lambda _mensagem: None)

        seletores = primeiro.call_args.args[1]
        self.assertIn((By.ID, "buttonLogOn"), seletores)

    def test_conexao_ignora_janela_remota_que_ja_estava_aberta(self):
        modulo = carregar_ercard()
        antiga = unittest.mock.MagicMock()
        antiga.class_name.return_value = "RAIL_WINDOW"
        antiga.is_visible.return_value = True
        antiga.handle = 100

        desktop = unittest.mock.MagicMock()
        desktop.windows.return_value = [antiga]
        pywinauto = unittest.mock.MagicMock()
        pywinauto.Desktop.return_value = desktop

        with patch.dict("sys.modules", {"pywinauto": pywinauto}):
            self.assertIsNone(modulo._nova_janela_remota({100}))

    def test_detecta_painel_canvas_quando_opcoes_estao_carregadas(self):
        import cv2
        import numpy as np

        modulo = carregar_ercard()
        imagem = np.zeros((500, 800, 3), dtype=np.uint8)
        imagem[:] = (28, 48, 63)
        cv2.rectangle(imagem, (150, 60), (670, 400), (250, 250, 250), -1)
        cv2.rectangle(imagem, (200, 220), (260, 275), (20, 20, 20), -1)
        sucesso, codificada = cv2.imencode(".png", imagem)

        self.assertTrue(sucesso)
        painel = modulo._localizar_painel_de_opcoes(codificada.tobytes())
        self.assertIsNotNone(painel)
        self.assertEqual(painel[:2], (150, 60))

    def test_detecta_dialogo_login_canvas(self):
        import cv2
        import numpy as np

        modulo = carregar_ercard()
        imagem = np.zeros((962, 1920, 3), dtype=np.uint8)
        imagem[:] = (28, 48, 63)
        cv2.rectangle(imagem, (710, 22), (1211, 491), (250, 250, 250), -1)
        sucesso, codificada = cv2.imencode(".png", imagem)

        self.assertTrue(sucesso)
        painel = modulo._localizar_dialogo_login(codificada.tobytes())
        self.assertIsNotNone(painel)
        self.assertEqual(painel[:2], (710, 22))

    def test_detecta_janela_gerencial_independente_da_posicao(self):
        import cv2
        import numpy as np

        modulo = carregar_ercard()
        imagem = np.full((900, 1600, 3), 140, dtype=np.uint8)
        cv2.rectangle(imagem, (730, 310), (1329, 709), (240, 240, 240), -1)
        sucesso, codificada = cv2.imencode(".png", imagem)

        self.assertTrue(sucesso)
        painel = modulo._localizar_janela_gerencial(codificada.tobytes())
        self.assertIsNotNone(painel)
        self.assertEqual(painel[:4], (730, 310, 600, 400))

    def test_data_inicial_usa_formato_digitado_pelo_ercard(self):
        modulo = carregar_ercard()
        self.assertEqual(modulo.DATA_INICIAL, "010124")


if __name__ == "__main__":
    unittest.main()
