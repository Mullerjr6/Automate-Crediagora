"""Restaura a tarefa recorrente do Tabela Contratos após o teste de hoje."""

import subprocess

if __name__ == "__main__":
    subprocess.run(
        ["schtasks.exe", "/Change", "/TN", "Crediagora - ERCard Tabela Contratos", "/ENABLE"],
        check=False,
    )
