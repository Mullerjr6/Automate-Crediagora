"""Dispara o Retorno de Ligações oficial em dias úteis que não sejam feriados nacionais."""

from __future__ import annotations

import json
import subprocess
import sys
import urllib.request
from datetime import date
from pathlib import Path

SCRIPT_DIR = Path(__file__).resolve().parent.parent
LOG = SCRIPT_DIR / "logs" / "agendado_retorno_ligacoes_oficial.log"
CACHE_DIR = SCRIPT_DIR / "logs" / "feriados"


def registrar(mensagem: str) -> None:
    linha = f"[{date.today():%d/%m/%Y}] {mensagem}"
    print(linha, flush=True)
    LOG.parent.mkdir(parents=True, exist_ok=True)
    with LOG.open("a", encoding="utf-8") as arquivo:
        arquivo.write(linha + "\n")


def domingo_de_pascoa(ano: int) -> date:
    """Calcula a Páscoa pelo algoritmo gregoriano de Meeus/Jones/Butcher."""
    a = ano % 19
    b, c = divmod(ano, 100)
    d, e = divmod(b, 4)
    f = (b + 8) // 25
    g = (b - f + 1) // 3
    h = (19 * a + b - d - g + 15) % 30
    i, k = divmod(c, 4)
    l = (32 + 2 * e + 2 * i - h - k) % 7
    m = (a + 11 * h + 22 * l) // 451
    mes = (h + l - 7 * m + 114) // 31
    dia = (h + l - 7 * m + 114) % 31 + 1
    return date(ano, mes, dia)


def feriados_nacionais_de_contingencia(ano: int) -> dict[date, str]:
    """Feriados nacionais e datas móveis usadas quando a consulta não estiver disponível."""
    pascoa = domingo_de_pascoa(ano)
    return {
        date(ano, 1, 1): "Confraternização Universal",
        pascoa.fromordinal(pascoa.toordinal() - 2): "Paixão de Cristo",
        date(ano, 4, 21): "Tiradentes",
        date(ano, 5, 1): "Dia Mundial do Trabalho",
        date(ano, 9, 7): "Independência do Brasil",
        date(ano, 10, 12): "Nossa Senhora Aparecida",
        date(ano, 11, 2): "Finados",
        date(ano, 11, 15): "Proclamação da República",
        date(ano, 11, 20): "Dia Nacional de Zumbi e da Consciência Negra",
        date(ano, 12, 25): "Natal",
    }


def feriados_nacionais(ano: int) -> dict[date, str]:
    """Obtém e armazena em cache os feriados nacionais publicados pela BrasilAPI."""
    CACHE_DIR.mkdir(parents=True, exist_ok=True)
    cache = CACHE_DIR / f"brasil-{ano}.json"
    feriados = feriados_nacionais_de_contingencia(ano)
    dados: list[dict[str, str]] = []
    try:
        with urllib.request.urlopen(
            f"https://brasilapi.com.br/api/feriados/v1/{ano}", timeout=15
        ) as resposta:
            dados = json.load(resposta)
        cache.write_text(json.dumps(dados, ensure_ascii=False, indent=2), encoding="utf-8")
    except (OSError, TimeoutError, json.JSONDecodeError):
        if cache.exists():
            try:
                dados = json.loads(cache.read_text(encoding="utf-8"))
            except json.JSONDecodeError:
                dados = []

    for item in dados:
        if isinstance(item, dict) and item.get("date") and item.get("name"):
            try:
                feriados[date.fromisoformat(item["date"])] = item["name"]
            except ValueError:
                continue
    return feriados


def main() -> int:
    hoje = date.today()
    if hoje.weekday() >= 5:
        registrar("Fim de semana: execução não iniciada.")
        return 0

    feriado = feriados_nacionais(hoje.year).get(hoje)
    if feriado:
        registrar(f"Feriado nacional ({feriado}): execução não iniciada.")
        return 0

    registrar("Iniciando agente do Retorno de Ligações oficial.")
    with LOG.open("a", encoding="utf-8") as arquivo_log:
        resultado = subprocess.run(
            [sys.executable, "-u", "agente_crediagora.py", "--fluxo", "retorno-ligacoes-oficial"],
            cwd=SCRIPT_DIR,
            check=False,
            stdout=arquivo_log,
            stderr=subprocess.STDOUT,
            text=True,
        )
    registrar(f"Agente finalizado com código {resultado.returncode}.")
    return resultado.returncode


if __name__ == "__main__":
    raise SystemExit(main())
