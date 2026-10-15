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
    def test_login_canvas_preenche_campos_separados_e_clica_ok(self):
        modulo = carregar_ercard()
        driver = unittest.mock.Mock()
        config = unittest.mock.Mock(
            timeout_remoto=20, sistema_usuario="teste", sistema_senha="segredo"
        )
        canvas = object()
        painel = (710, 214, 501, 533, 1920, 962)
        with patch.object(modulo, "_localizar_dialogo_login", side_effect=[painel, None]), \
             patch.object(modulo, "_canvas_principal", return_value=canvas), \
             patch.object(modulo, "_acao_canvas") as clicar, \
             patch.object(modulo, "ActionChains") as acoes, \
             patch.object(modulo.time, "sleep"):
            modulo.realizar_login_sistema_canvas(driver, config, lambda _: None)

        self.assertEqual(clicar.call_args_list, [
            unittest.mock.call(driver, canvas, 1085, 624, 1920, 962),
            unittest.mock.call(driver, canvas, 1085, 653, 1920, 962),
            unittest.mock.call(driver, canvas, 910, 715, 1920, 962),
        ])
        digitacao = acoes.return_value.send_keys.call_args_list
        self.assertEqual(digitacao, [unittest.mock.call(c) for c in "testesegredo"])
        self.assertEqual(acoes.return_value.send_keys.return_value.pause.call_count, len("testesegredo"))

    def test_csv_bloqueado_e_retentado_sem_reiniciar_exportacao(self):
        modulo = carregar_ercard()
        with tempfile.TemporaryDirectory() as pasta:
            downloads = Path(pasta) / "downloads"
            downloads.mkdir()
            destino = Path(pasta) / "final" / "contratos.csv"
            (downloads / destino.name).write_text("codigo;contrato\n1;ABC\n", encoding="utf-8")
            config = unittest.mock.Mock(timeout_exportacao=60, pasta_download=downloads)
            tempo = [0]
            mover = modulo.shutil.move
            tentativas = [0]

            def mover_quando_liberado(*args):
                tentativas[0] += 1
                if tentativas[0] == 1:
                    raise PermissionError("arquivo em gravacao")
                return mover(*args)

            with patch.object(modulo.time, "monotonic", side_effect=lambda: tempo[0]), \
                 patch.object(modulo.time, "sleep", side_effect=lambda _: tempo.__setitem__(0, tempo[0] + 1)), \
                 patch.object(modulo.shutil, "move", side_effect=mover_quando_liberado):
                resultado = modulo.mover_download_webfile(destino, config, lambda _: None)
            self.assertEqual(resultado, destino)
            self.assertEqual(tentativas[0], 2)
            self.assertTrue(destino.is_file())

    def test_timeout_do_csv_sinaliza_preservacao_da_sessao(self):
        modulo = carregar_ercard()
        with tempfile.TemporaryDirectory() as pasta:
            config = unittest.mock.Mock(timeout_exportacao=60, pasta_download=Path(pasta))
            with patch.object(modulo.time, "monotonic", side_effect=[0, 61]):
                with self.assertRaises(modulo.ErCardError) as contexto:
                    modulo.mover_download_webfile(Path(pasta) / "contratos.csv", config, lambda _: None)
            self.assertTrue(contexto.exception.preservar_sessao)
            self.assertIn("nao repetir Exportar", str(contexto.exception))

    def test_webfile_aguarda_csv_apos_mais_de_quinze_minutos(self):
        modulo = carregar_ercard()
        with tempfile.TemporaryDirectory() as pasta:
            downloads = Path(pasta) / "downloads"
            downloads.mkdir()
            destino = Path(pasta) / "final" / "contratos.csv"
            origem = downloads / destino.name
            config = unittest.mock.Mock(timeout_exportacao=1800, pasta_download=downloads)
            tempo = [0]
            logs = []

            def aguardar(_segundos):
                tempo[0] += 60
                if tempo[0] >= 960 and not origem.exists():
                    origem.write_text("codigo;contrato\n1;ABC\n", encoding="utf-8")

            with patch.object(modulo.time, "monotonic", side_effect=lambda: tempo[0]), \
                 patch.object(modulo.time, "sleep", side_effect=aguardar):
                resultado = modulo.mover_download_webfile(destino, config, logs.append)

            self.assertEqual(resultado, destino)
            self.assertTrue(destino.is_file())
            self.assertFalse(origem.exists())
            self.assertGreaterEqual(tempo[0], 960)
            self.assertTrue(any("Aguardando consulta 14" in mensagem for mensagem in logs))

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

    def test_modulo_canvas_usa_duplo_clique_e_confirma_mudanca(self):
        modulo = carregar_ercard()
        driver = unittest.mock.MagicMock()
        canvas = unittest.mock.MagicMock()
        canvas.is_displayed.return_value = True
        canvas.size = {"width": 1600, "height": 900}
        driver.find_elements.return_value = [canvas]
        driver.get_screenshot_as_png.return_value = b"imagem"
        driver.execute_script.return_value = {"x": 10, "y": 20}
        acao = unittest.mock.MagicMock()
        acao.move_to_element_with_offset.return_value = acao
        acao.double_click.return_value = acao

        with patch.object(
            modulo,
            "_localizar_painel_de_opcoes",
            return_value=(700, 290, 520, 340, 1920, 960),
        ), patch.object(
            modulo, "_mudanca_visual_significativa", return_value=True
        ), patch.object(
            modulo, "ActionChains", return_value=acao
        ), patch.object(modulo.time, "sleep"):
            resultado = modulo._clicar_opcao_canvas(
                driver,
                timeout=2,
                etapa="Módulo ER Cartão CrediAgora",
                log=lambda _mensagem: None,
                posicao_horizontal=0.38,
                nome_opcao="ER Cartão CrediAgora",
                cliques=2,
                confirmar_mudanca=True,
            )

        self.assertTrue(resultado)
        acao.double_click.assert_called_once_with()
        acao.perform.assert_called_once_with()

    def test_mudanca_no_painel_e_detectada_sem_alterar_tela_inteira(self):
        import cv2
        import numpy as np

        modulo = carregar_ercard()
        antes = np.zeros((960, 1920, 3), dtype=np.uint8)
        depois = antes.copy()
        cv2.rectangle(antes, (700, 290), (1220, 630), (245, 245, 245), -1)
        cv2.rectangle(depois, (700, 290), (1220, 630), (245, 245, 245), -1)
        cv2.rectangle(antes, (790, 470), (820, 500), (20, 20, 20), -1)
        cv2.rectangle(depois, (790, 470), (820, 500), (20, 20, 20), -1)
        cv2.rectangle(depois, (930, 470), (980, 520), (20, 20, 20), -1)
        _, antes_png = cv2.imencode(".png", antes)
        _, depois_png = cv2.imencode(".png", depois)

        self.assertFalse(
            modulo._mudanca_visual_significativa(
                antes_png.tobytes(), depois_png.tobytes()
            )
        )
        self.assertTrue(
            modulo._mudanca_visual_significativa(
                antes_png.tobytes(),
                depois_png.tobytes(),
                regiao=(700, 290, 520, 340),
                proporcao_minima=0.01,
            )
        )

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

    def test_periodo_canvas_preserva_foco_do_tab_e_data_inicial(self):
        modulo = carregar_ercard()
        for hoje in (datetime(2026, 9, 25), datetime(2027, 1, 2)):
            with self.subTest(hoje=hoje):
                eventos = []
                campos = ["data antiga", "outra data antiga", ""]
                foco = [0]
                controle = [False]

                class Acoes:
                    def key_down(self, tecla):
                        if tecla == modulo.Keys.CONTROL:
                            controle[0] = True
                        return self

                    def key_up(self, tecla):
                        if tecla == modulo.Keys.CONTROL:
                            controle[0] = False
                        return self

                    def send_keys(self, tecla):
                        eventos.append(("tecla", tecla))
                        if tecla == modulo.Keys.TAB:
                            foco[0] += 1
                        elif tecla in (modulo.Keys.BACKSPACE, modulo.Keys.DELETE):
                            campos[foco[0]] = ""
                        elif tecla.isdigit() and not controle[0]:
                            campos[foco[0]] += tecla
                        return self

                    def perform(self):
                        return None

                def clicar(*args):
                    eventos.append(("clique", args[1:]))
                    foco[0] = 0

                with patch.object(modulo, "datetime") as relogio, \
                     patch.object(modulo, "ActionChains", side_effect=lambda _: Acoes()), \
                     patch.object(modulo, "_clicar_janela_gerencial", side_effect=clicar) as clique, \
                     patch.object(modulo, "salvar_diagnostico_ercard") as captura, \
                     patch.object(modulo.time, "sleep"):
                    relogio.now.return_value = hoje
                    driver, config = unittest.mock.Mock(), unittest.mock.Mock()
                    modulo.preencher_periodo_canvas(driver, config, lambda _: None)

                self.assertEqual(campos, ["010124", hoje.strftime("%d%m%y"), ""])
                self.assertEqual(foco[0], 2)
                self.assertEqual(clique.call_count, 2)
                primeiro_tab = eventos.index(("tecla", modulo.Keys.TAB))
                self.assertFalse(any(tipo == "clique" for tipo, _ in eventos[primeiro_tab:]))
                captura.assert_called_once()


if __name__ == "__main__":
    unittest.main()
