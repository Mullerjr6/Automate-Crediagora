import contextlib
import importlib.util
import io
import os
import sys
import tempfile
import time
import unittest
from pathlib import Path
from unittest.mock import patch


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
    def test_encerramento_sessao_inativa_nao_registra_falha(self):
        for erro in ("sem_janelas", "sessao_invalida", "janela_fechada"):
            with self.subTest(erro=erro):
                modulo, _ = carregar_start()
                driver = unittest.mock.MagicMock()
                if erro == "sem_janelas":
                    driver.window_handles = []
                else:
                    excecao = (
                        modulo.InvalidSessionIdException if erro == "sessao_invalida"
                        else modulo.NoSuchWindowException
                    )
                    type(driver).window_handles = unittest.mock.PropertyMock(
                        side_effect=excecao("session deleted")
                    )
                with patch.object(modulo, "log") as registrar:
                    self.assertFalse(modulo.encerrar_sessao_portal(driver))
                driver.find_elements.assert_not_called()
                mensagens = " ".join(chamada.args[0] for chamada in registrar.call_args_list)
                self.assertIn("dispensada", mensagens)
                self.assertNotIn("session deleted", mensagens)
                self.assertNotIn("não encontrada", mensagens)

    def test_somente_ercard_nao_tenta_logout_panorama(self):
        modulo, _ = carregar_start()
        driver = unittest.mock.MagicMock()
        with patch.object(modulo, "MODO_HEADLESS", False), \
             patch.object(modulo, "MANTER_NAVEGADOR", False), \
             patch.object(modulo, "criar_pastas"), \
             patch.object(modulo, "criar_configuracao_ercard"), \
             patch.object(modulo, "ativar_ercard_se_configurado", return_value=True), \
             patch.object(modulo, "mostrar_resumo_configuracao"), \
             patch.object(modulo, "iniciar_chrome", return_value=driver), \
             patch.object(modulo, "executar_fase_ercard", return_value="contratos.csv"), \
             patch.object(modulo, "atualizar_indicadores_fpd1", return_value=unittest.mock.MagicMock()) as atualizar, \
             patch.object(modulo, "encerrar_sessao_portal") as sair, \
             patch.object(modulo, "limpar_downloads"), \
             patch.object(modulo, "salvar_log_ativacao"), \
             patch.object(modulo, "log"), \
             patch.object(modulo.time, "sleep"):
            modulo.main(["--somente-ercard"])
        atualizar.assert_called_once()
        sair.assert_not_called()
        driver.quit.assert_called_once()

    def test_timeout_ercard_preserva_navegador_e_nao_atualiza_excel(self):
        for preservar in (True, False):
            with self.subTest(preservar=preservar):
                modulo, _ = carregar_start()
                driver = unittest.mock.MagicMock()
                erro = RuntimeError("CSV nao confirmado")
                erro.preservar_sessao = preservar
                with patch.object(modulo, "MODO_HEADLESS", False), \
                     patch.object(modulo, "MANTER_NAVEGADOR", False), \
                     patch.object(modulo, "criar_pastas"), \
                     patch.object(modulo, "criar_configuracao_ercard"), \
                     patch.object(modulo, "ativar_ercard_se_configurado", return_value=True), \
                     patch.object(modulo, "mostrar_resumo_configuracao"), \
                     patch.object(modulo, "iniciar_chrome", return_value=driver) as abrir, \
                     patch.object(modulo, "executar_fase_ercard", side_effect=erro) as fase, \
                     patch.object(modulo, "atualizar_indicadores_fpd1") as atualizar, \
                     patch.object(modulo, "salvar_diagnostico_erro"), \
                     patch.object(modulo, "salvar_log_ativacao"), \
                     patch.object(modulo, "log"), \
                     patch.object(modulo.time, "sleep"):
                    self.assertEqual(modulo.main(["--somente-ercard"]), 1)
                abrir.assert_called_once_with(headless=False, manter_navegador=True)
                fase.assert_called_once()
                atualizar.assert_not_called()
                if preservar:
                    driver.quit.assert_not_called()
                else:
                    driver.quit.assert_called_once()

    def test_indicadores_usa_nome_sem_acento_na_pasta_configurada(self):
        with tempfile.TemporaryDirectory() as pasta, patch.dict(os.environ, {
            "CREDIAGORA_TABELA_FAT_DIR": pasta,
            "CREDIAGORA_INDICADORES_FPD1_XLSX": "",
        }):
            esperado = Path(pasta) / "Indicadores de FPD1_teste_de_automacao.xlsx"
            esperado.touch()
            modulo, _ = carregar_start()
            self.assertEqual(modulo.ARQUIVO_INDICADORES_FPD1, esperado)
            self.assertTrue(modulo.ARQUIVO_INDICADORES_FPD1.is_file())

    def test_indicadores_respeita_caminho_explicito(self):
        with tempfile.TemporaryDirectory() as pasta:
            esperado = Path(pasta) / "indicadores_personalizados.xlsx"
            with patch.dict(os.environ, {"CREDIAGORA_INDICADORES_FPD1_XLSX": str(esperado)}):
                modulo, _ = carregar_start()
            self.assertEqual(modulo.ARQUIVO_INDICADORES_FPD1, esperado)

    def test_contexto_panorama_ignora_frame_removido(self):
        modulo, _ = carregar_start()
        driver = unittest.mock.MagicMock()
        antigo, valido = unittest.mock.MagicMock(), unittest.mock.MagicMock()
        antigo.is_displayed.side_effect = modulo.StaleElementReferenceException()
        driver.find_elements.return_value = [antigo, valido]
        verificar = unittest.mock.Mock(side_effect=[None, "concluida"])
        self.assertEqual(modulo._buscar_contexto_panorama(driver, verificar), "concluida")
        driver.switch_to.frame.assert_called_once_with(valido)
        driver.switch_to.parent_frame.assert_not_called()

    def test_contexto_panorama_ignora_frame_indisponivel(self):
        modulo, _ = carregar_start()
        driver = unittest.mock.MagicMock()
        antigo, valido = unittest.mock.MagicMock(), unittest.mock.MagicMock()
        driver.find_elements.return_value = [antigo, valido]
        driver.switch_to.frame.side_effect = [modulo.NoSuchFrameException(), None]
        verificar = unittest.mock.Mock(side_effect=[None, "concluida"])
        self.assertEqual(modulo._buscar_contexto_panorama(driver, verificar), "concluida")
        self.assertEqual(driver.switch_to.frame.call_count, 2)

    def test_contexto_panorama_restaura_pai_apos_dom_mudar(self):
        modulo, _ = carregar_start()
        driver = unittest.mock.MagicMock()
        antigo, valido = unittest.mock.MagicMock(), unittest.mock.MagicMock()
        driver.find_elements.return_value = [antigo, valido]
        verificar = unittest.mock.Mock(side_effect=[
            None, modulo.StaleElementReferenceException(), "concluida",
        ])
        self.assertEqual(modulo._buscar_contexto_panorama(driver, verificar), "concluida")
        driver.switch_to.parent_frame.assert_called_once()
        self.assertEqual(driver.switch_to.frame.call_args.args, (valido,))

    def test_contexto_panorama_aninhado_preserva_contexto_do_resultado(self):
        modulo, _ = carregar_start()
        driver = unittest.mock.MagicMock()
        externo, interno = unittest.mock.MagicMock(), unittest.mock.MagicMock()
        driver.find_elements.side_effect = [[externo], [interno]]
        verificar = unittest.mock.Mock(side_effect=[None, None, "concluida"])
        self.assertEqual(modulo._buscar_contexto_panorama(driver, verificar), "concluida")
        driver.switch_to.parent_frame.assert_not_called()

    def test_executar_interceptado_em_frame_fecha_aviso_e_relocaliza(self):
        modulo, _ = carregar_start()
        driver = unittest.mock.MagicMock()
        botao, frame = unittest.mock.MagicMock(), unittest.mock.MagicMock()
        contexto = ["principal"]
        driver.switch_to.default_content.side_effect = lambda: contexto.__setitem__(0, "principal")
        driver.switch_to.parent_frame.side_effect = lambda: contexto.__setitem__(0, "principal")
        driver.switch_to.frame.side_effect = lambda _: contexto.__setitem__(0, "frame")
        driver.find_elements.side_effect = lambda by, seletor: (
            [frame] if seletor == "iframe, frame" and contexto[0] == "principal"
            else [botao] if seletor != "iframe, frame" and contexto[0] == "frame"
            else []
        )
        botao.click.side_effect = [modulo.ElementClickInterceptedException(), None]

        def fechar(*args, **kwargs):
            self.assertEqual(contexto[0], "principal")

        with patch.object(modulo, "_resultado_exportacao_visivel", side_effect=lambda _: (
            contexto[0] == "frame" and botao.click.call_count == 2
        )), patch.object(modulo, "fechar_avisos_exportacao_panorama", side_effect=fechar) as avisos, \
             patch.object(modulo.time, "sleep"):
            modulo.clicar_executar_exportacao(driver)
        avisos.assert_called_once_with(driver, aguardar_segundos=0)
        self.assertEqual(botao.click.call_count, 2)

    def test_executar_resultado_demorado_nao_duplica_envio(self):
        modulo, _ = carregar_start()
        driver = unittest.mock.MagicMock()
        botao = unittest.mock.MagicMock()
        driver.find_elements.side_effect = lambda by, seletor: (
            [] if seletor == "iframe, frame" else [botao]
        )
        with patch.object(modulo, "_resultado_exportacao_visivel", side_effect=[False, False, True]), \
             patch.object(modulo.time, "sleep"):
            modulo.clicar_executar_exportacao(driver)
        botao.click.assert_called_once()

    def test_executar_ambiguo_salva_diagnostico_sem_clicar(self):
        modulo, _ = carregar_start()
        driver = unittest.mock.MagicMock()
        primeiro, segundo = unittest.mock.MagicMock(), unittest.mock.MagicMock()
        driver.find_elements.return_value = [primeiro, segundo]
        with patch.object(modulo, "_resultado_exportacao_visivel", return_value=False), \
             patch.object(modulo, "salvar_diagnostico_erro") as diagnostico:
            with self.assertRaisesRegex(RuntimeError, "Mais de um botao"):
                modulo.clicar_executar_exportacao(driver)
        primeiro.click.assert_not_called()
        segundo.click.assert_not_called()
        diagnostico.assert_called_once_with(driver, "erro_executar")

    def test_executar_timeout_registra_progresso_e_salva_diagnostico(self):
        modulo, _ = carregar_start()
        driver = unittest.mock.MagicMock()
        driver.find_elements.return_value = []
        with patch.object(modulo, "_resultado_exportacao_visivel", return_value=False), \
             patch.object(modulo.time, "monotonic", side_effect=[0, 0, 11, 121]), \
             patch.object(modulo.time, "sleep"), \
             patch.object(modulo, "log_panorama") as log, \
             patch.object(modulo, "salvar_diagnostico_erro") as diagnostico:
            with self.assertRaisesRegex(RuntimeError, "Clique enviado=False"):
                modulo.clicar_executar_exportacao(driver)
        diagnostico.assert_called_once_with(driver, "erro_executar")
        self.assertTrue(any("11s decorridos" in call.args[1] for call in log.call_args_list))

    def test_executar_nao_repete_clique_apos_troca_do_dom(self):
        modulo, _ = carregar_start()
        driver = unittest.mock.MagicMock()
        botao = unittest.mock.MagicMock()
        botao.click.side_effect = modulo.StaleElementReferenceException()
        driver.find_elements.side_effect = lambda by, seletor: (
            [] if seletor == "iframe, frame" else [botao]
        )
        with patch.object(modulo, "_resultado_exportacao_visivel", side_effect=[False, True]), \
             patch.object(modulo.time, "sleep"):
            modulo.clicar_executar_exportacao(driver)
        botao.click.assert_called_once()

    def test_executar_reconhece_resultado_sem_clicar_novamente(self):
        modulo, _ = carregar_start()
        driver = unittest.mock.MagicMock()
        with patch.object(modulo, "_resultado_exportacao_visivel", return_value=True):
            modulo.clicar_executar_exportacao(driver)
        driver.find_elements.assert_not_called()

    def test_resultado_exportacao_exige_mensagem_e_download(self):
        modulo, _ = carregar_start()
        driver = unittest.mock.MagicMock()
        mensagem = unittest.mock.MagicMock()
        driver.find_elements.side_effect = [[mensagem], []]
        self.assertFalse(modulo._resultado_exportacao_visivel(driver))

    def test_etapa_panorama_identifica_os_dois_layouts(self):
        modulo, _ = carregar_start()
        self.assertEqual(modulo.etapa_panorama("Exportação vendas"), "VENDAS")
        self.assertEqual(modulo.etapa_panorama("Receita gerada"), "RECEITA")

    def test_reutiliza_tela_emprestimos_sem_reabrir_menu(self):
        modulo, _ = carregar_start()
        driver = unittest.mock.MagicMock()
        with patch.object(modulo, "fechar_avisos_exportacao_panorama", return_value=0), \
             patch.object(modulo, "tratar_popups_panorama", return_value=0), \
             patch.object(modulo, "ativar_tela_emprestimos_aberta", return_value=True), \
             patch.object(modulo, "clicar_por_texto") as clicar:
            modulo.entrar_em_vendas_e_emprestimo(driver)
        clicar.assert_not_called()

    def test_submenu_emprestimos_aberto_nao_clica_em_vendas(self):
        modulo, _ = carregar_start()
        driver = unittest.mock.MagicMock()

        with patch.object(modulo, "fechar_avisos_exportacao_panorama", return_value=0), \
             patch.object(modulo, "tratar_popups_panorama", return_value=0), \
             patch.object(modulo, "ativar_tela_emprestimos_aberta", return_value=False), \
             patch.object(
                 modulo,
                 "clicar_item_emprestimo_visivel",
                 return_value=("conteudo principal", None),
             ), \
             patch.object(modulo, "clicar_por_texto") as clicar_vendas, \
             patch.object(modulo, "WebDriverWait") as espera:
            espera.return_value.until.return_value = unittest.mock.MagicMock()
            modulo.entrar_em_vendas_e_emprestimo(driver)

        clicar_vendas.assert_not_called()

    def test_localiza_item_emprestimo_visivel_antes_de_alternar_menu(self):
        modulo, _ = carregar_start()
        driver = unittest.mock.MagicMock()
        item = unittest.mock.MagicMock()
        item.is_displayed.return_value = True
        item.is_enabled.return_value = True
        item.text = "Empréstimo"
        item.tag_name = "a"

        def encontrar(by, seletor):
            if by == modulo.By.XPATH and seletor == modulo.SELETORES_ITEM_EMPRESTIMO[0][1]:
                return [item]
            return []

        driver.find_elements.side_effect = encontrar
        with patch.object(modulo, "clicar_elemento_seguro") as clicar, \
             patch.object(modulo.time, "sleep"):
            contexto, erro = modulo.clicar_item_emprestimo_visivel(
                driver, "já aberto"
            )

        self.assertEqual(contexto, "conteúdo principal")
        self.assertIsNone(erro)
        clicar.assert_called_once_with(
            driver, item, "Empréstimo em conteúdo principal"
        )

    def test_aviso_verde_usa_o_x_real_confirmar_leitura(self):
        modulo, _ = carregar_start()
        driver = unittest.mock.MagicMock()
        botao = unittest.mock.MagicMock()
        botao.is_displayed.side_effect = [True, False]
        botao.is_enabled.return_value = True
        driver.current_window_handle = "principal"
        driver.window_handles = ["principal"]

        def encontrar(by, seletor):
            if by == modulo.By.TAG_NAME and seletor == "iframe":
                return []
            if by == modulo.By.CSS_SELECTOR and "confirmar leitura do aviso" in seletor:
                return [botao]
            return []

        driver.find_elements.side_effect = encontrar
        with patch.object(modulo, "clicar_elemento_seguro") as clicar, \
             patch.object(modulo.time, "sleep"):
            quantidade = modulo.fechar_avisos_exportacao_panorama(driver)

        self.assertEqual(quantidade, 1)
        clicar.assert_called_once_with(
            driver, botao, "X real do aviso automatico de exportacao"
        )

    def test_download_fecha_todos_os_avisos_automaticos(self):
        modulo, _ = carregar_start()
        driver = unittest.mock.MagicMock()
        driver.current_window_handle = "principal"
        driver.window_handles = ["principal"]
        driver.find_elements.return_value = []
        with patch.object(modulo, "limpar_downloads"), \
             patch.object(modulo, "abrir_tela_exportacao"), \
             patch.object(modulo, "selecionar_layout"), \
             patch.object(modulo, "clicar_executar_exportacao"), \
             patch.object(modulo, "esperar"), \
             patch.object(modulo, "clicar_primeiro_download"), \
             patch.object(modulo, "esperar_download_novo", return_value=Path("vendas.csv")), \
             patch.object(modulo, "fechar_avisos_exportacao_panorama", return_value=2) as fechar, \
             patch.object(modulo.time, "sleep"):
            resultado = modulo.baixar_exportacao(driver, "Exportação vendas")
        self.assertEqual(resultado, Path("vendas.csv"))
        self.assertEqual(fechar.call_count, 3)
        fechar.assert_any_call(driver, "VENDAS", aguardar_segundos=1.5)
        fechar.assert_any_call(driver, "VENDAS", aguardar_segundos=1)
        fechar.assert_any_call(driver, "VENDAS", aguardar_segundos=2)

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

    def test_argumento_somente_gestor(self):
        modulo, _ = carregar_start()
        args = modulo.parse_args(["--somente-gestor"])
        self.assertTrue(args.somente_gestor)
        self.assertFalse(args.somente_ercard)
        self.assertFalse(args.sem_gestor)

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

    def test_maximiza_chrome_no_monitor_principal_antes_do_login(self):
        modulo, _ = carregar_start()
        driver = unittest.mock.MagicMock()

        def executar_cdp(comando, parametros):
            if comando == "Browser.getWindowForTarget":
                return {"windowId": 17}
            if comando == "Browser.getWindowBounds":
                return {
                    "bounds": {
                        "left": 0,
                        "top": 0,
                        "width": 1920,
                        "height": 1080,
                        "windowState": "maximized",
                    }
                }
            return {}

        driver.execute_cdp_cmd.side_effect = executar_cdp
        with patch.object(
            modulo, "_dimensoes_monitor_principal", return_value=(1920, 1080)
        ), patch.object(modulo.time, "sleep"):
            estado = modulo.maximizar_chrome_antes_login(driver)

        comandos = [chamada.args[0] for chamada in driver.execute_cdp_cmd.call_args_list]
        self.assertEqual(
            comandos,
            [
                "Browser.getWindowForTarget",
                "Browser.setWindowBounds",
                "Browser.setWindowBounds",
                "Browser.setWindowBounds",
                "Browser.getWindowBounds",
            ],
        )
        self.assertEqual(estado["windowState"], "maximized")

    def test_login_normaliza_janela_antes_de_abrir_o_site(self):
        modulo, _ = carregar_start()
        driver = unittest.mock.MagicMock()

        with patch.object(
            modulo,
            "maximizar_chrome_antes_login",
            side_effect=RuntimeError("janela invalida"),
        ):
            with self.assertRaisesRegex(RuntimeError, "janela invalida"):
                modulo.fazer_login(driver)

        driver.get.assert_not_called()


if __name__ == "__main__":
    unittest.main()
