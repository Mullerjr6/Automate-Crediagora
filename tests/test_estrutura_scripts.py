import importlib.util
import subprocess
import sys
import unittest
from pathlib import Path


RAIZ = Path(__file__).resolve().parents[1]
SCRIPT = RAIZ / "CREDIAGORA" / "Script"


class EstruturaScriptsTests(unittest.TestCase):
    def test_testes_manuais_importam_sem_executar_automacao(self):
        pasta = SCRIPT / "testes_manuais"
        nomes = [arquivo.stem for arquivo in sorted(pasta.glob("teste_*.py"))]
        self.assertEqual(len(nomes), 10)
        codigo = (
            "import importlib, sys; "
            f"sys.path.insert(0, {str(pasta)!r}); "
            f"[importlib.import_module(nome) for nome in {nomes!r}]"
        )
        resultado = subprocess.run(
            [sys.executable, "-c", codigo], cwd=RAIZ,
            capture_output=True, text=True, timeout=30,
        )
        self.assertEqual(resultado.returncode, 0, resultado.stderr)

    def test_ferramenta_limpeza_nao_executa_ao_importar(self):
        caminho = SCRIPT / "ferramentas" / "limpar_downloads_crediagora.py"
        spec = importlib.util.spec_from_file_location("ferramenta_limpeza", caminho)
        modulo = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(modulo)
        self.assertTrue(callable(modulo.limpar_downloads))

    def test_pontos_de_entrada_preservados(self):
        for nome in (
            "start.py", "iniciar.bat", "iniciar_ercard.bat",
            "configurar_ercard.bat", "configurar_ercard.ps1",
            "diagnostico.bat", "diagnostico_ercard.bat", "liberar_login.bat",
            "ocr_windows.ps1",
        ):
            with self.subTest(nome=nome):
                self.assertTrue((SCRIPT / nome).is_file())
        self.assertFalse((SCRIPT / "python").exists())


if __name__ == "__main__":
    unittest.main()
