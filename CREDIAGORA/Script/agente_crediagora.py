"""Orquestrador local dos fluxos Crediagora.

Executa os scripts existentes sem duplicar a lógica de Panorama/ERCard,
registrando estado, duração, tentativas e histórico em JSON.
"""

import argparse
import json
import os
import shutil
import subprocess
import sys
import time
import urllib.request
from html import escape
from datetime import datetime
from pathlib import Path


SCRIPT_DIR = Path(__file__).resolve().parent
ESTADO = SCRIPT_DIR / "logs" / "agente_estado.json"
HISTORICO = SCRIPT_DIR / "logs" / "agente_historico.jsonl"
TRAVA = SCRIPT_DIR / "logs" / "agente_crediagora.lock"
CONFIG_PADRAO = SCRIPT_DIR / "agente_config.json"
DASHBOARD = SCRIPT_DIR / "logs" / "agente_dashboard.html"
RELATORIO = SCRIPT_DIR / "logs" / "agente_relatorio_final.json"

FLUXOS = {
    "completo": ["start.py", "--headless"],
    "ercard-carteira": ["ercard_carteira.py"],
    "retorno-ligacoes": [
        "start.py", "--exportacao", "retorno-ligacoes",
        "--sem-ercard", "--sem-gestor", "--headless",
    ],
}


def carregar_configuracao(caminho):
    caminho = Path(caminho)
    if not caminho.exists():
        return {}
    try:
        dados = json.loads(caminho.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as erro:
        raise RuntimeError(f"Configuração inválida: {caminho}: {erro}") from erro
    if not isinstance(dados, dict):
        raise RuntimeError("A configuração do agente deve ser um objeto JSON.")
    return dados


def adquirir_trava():
    TRAVA.parent.mkdir(parents=True, exist_ok=True)
    try:
        fd = os.open(TRAVA, os.O_CREAT | os.O_EXCL | os.O_WRONLY)
        os.write(fd, f"pid={os.getpid()} inicio={agora()}".encode("utf-8"))
        os.close(fd)
    except FileExistsError as erro:
        raise RuntimeError(f"Já existe uma execução do agente: {TRAVA}") from erro


def liberar_trava():
    try:
        TRAVA.unlink()
    except FileNotFoundError:
        pass


def validar_ambiente(fluxo):
    faltantes = [str(SCRIPT_DIR / nome) for nome in FLUXOS[fluxo] if nome.endswith(".py") and not (SCRIPT_DIR / nome).exists()]
    if faltantes:
        raise FileNotFoundError("Arquivos do fluxo não encontrados: " + ", ".join(faltantes))
    (SCRIPT_DIR / "logs").mkdir(parents=True, exist_ok=True)


def verificar_internet():
    try:
        with urllib.request.urlopen("https://crediagora.panoramaemprestimos.com.br", timeout=8):
            return True
    except Exception:
        return False


def notificar(config, mensagem):
    url = str(config.get("webhook_url", "")).strip()
    if not url:
        return
    dados = json.dumps({"text": mensagem}).encode("utf-8")
    requisicao = urllib.request.Request(url, data=dados, headers={"Content-Type": "application/json"})
    try:
        urllib.request.urlopen(requisicao, timeout=10).close()
    except Exception:
        pass


def validar_resultado(fluxo):
    if fluxo == "ercard-carteira":
        pasta = Path.home() / "Desktop" / "ATT manh\u00e3"
        arquivos = list(pasta.glob("*.csv"))
        if len(arquivos) < 3:
            return False
        for arquivo in arquivos:
            if arquivo.stat().st_size <= 0 or "\n" not in arquivo.read_text(encoding="utf-8-sig", errors="replace")[:4096]:
                return False
        return True
    if fluxo == "retorno-ligacoes":
        return (Path.home() / "TJI PROMOTORA DE VENDAS EIRELI" / "Crediagora-doc - dados" / "tabela fat" / "fat_retorno_ligacoes - teste automacao.xlsx").exists()
    return True


def criar_backups(config):
    destinos = config.get("backup_paths") or [
        str(Path.home() / "TJI PROMOTORA DE VENDAS EIRELI" / "Crediagora-doc - dados" / "tabela fat" / nome)
        for nome in (
            "fat_retorno_ligacoes - teste automacao.xlsx",
            "fat_vendas_Teste.xlsx",
            "fat_receita_gerada_CPC_Teste.xlsx",
        )
    ]
    if not isinstance(destinos, list):
        return []
    pasta = SCRIPT_DIR / "backups" / datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    copiados = []
    for item in destinos:
        origem = Path(os.path.expandvars(str(item))).expanduser()
        if origem.is_file():
            pasta.mkdir(parents=True, exist_ok=True)
            destino = pasta / origem.name
            shutil.copy2(origem, destino)
            copiados.append(str(destino))
    return copiados


def agora():
    return datetime.now().astimezone().isoformat(timespec="seconds")


def atualizar_dashboard():
    registros = []
    if HISTORICO.exists():
        for linha in HISTORICO.read_text(encoding="utf-8").splitlines()[-30:]:
            try:
                registros.append(json.loads(linha))
            except json.JSONDecodeError:
                pass
    linhas = "".join(
        f"<tr><td>{escape(str(r.get('fluxo','')))}</td><td class='status status-{escape(str(r.get('status','')))}'>{escape(str(r.get('status','')))}</td>"
        f"<td>{escape(str(r.get('inicio','')))}</td><td>{escape(str(r.get('duracao_minutos','')))}</td></tr>"
        for r in reversed(registros)
    )
    DASHBOARD.parent.mkdir(parents=True, exist_ok=True)
    DASHBOARD.write_text(
        "<!doctype html><html lang='pt-BR'><head><meta charset='utf-8'>"
        "<meta name='viewport' content='width=device-width,initial-scale=1'>"
        "<title>Agente Crediagora</title><style>"
        ":root{--bg:#f4f7fb;--card:#fff;--ink:#172033;--muted:#667085;--brand:#2457d6;--line:#e4e9f2}"
        "*{box-sizing:border-box}body{margin:0;background:var(--bg);color:var(--ink);font:14px Segoe UI,Arial,sans-serif}"
        ".wrap{max-width:1100px;margin:0 auto;padding:32px 20px}.top{display:flex;justify-content:space-between;align-items:center;margin-bottom:24px}"
        "h1{font-size:28px;margin:0}.sub{color:var(--muted);margin-top:6px}.badge{background:#e8efff;color:var(--brand);padding:8px 12px;border-radius:999px;font-weight:600}"
        ".card{background:var(--card);border:1px solid var(--line);border-radius:14px;box-shadow:0 4px 18px #182b4d0d;padding:20px}.table-wrap{overflow:auto}"
        ".history{width:100%;border-collapse:collapse;min-width:650px}.history th{text-align:left;color:var(--muted);font-size:12px;text-transform:uppercase;letter-spacing:.04em;padding:12px;border-bottom:2px solid var(--line)}"
        ".history td{padding:14px 12px;border-bottom:1px solid var(--line)}.history tr:hover{background:#f8faff}.status{font-weight:600}.status-concluido{color:#16834b}.status-erro{color:#c0392b}.status-executando{color:#b7791f}"
        "@media(max-width:700px){.top{display:block}.badge{display:inline-block;margin-top:14px}h1{font-size:23px}}"
        "</style></head><body><main class='wrap'><header class='top'><div><h1>Agente Crediagora</h1><div class='sub'>Histórico das automações</div></div><div class='badge'>Atualizado automaticamente</div></header><section class='card table-wrap'>"
        "<table class='history'><thead><tr><th>Fluxo</th><th>Status</th><th>Início</th><th>Duração (min)</th></tr></thead><tbody>" + linhas + "</tbody></table></section></main></body></html>", encoding="utf-8"
    )


def limpar_backups_antigos(config):
    dias = int(config.get("backup_retencao_dias", 30))
    pasta = SCRIPT_DIR / "backups"
    limite = time.time() - dias * 86400
    if dias > 0 and pasta.exists():
        for item in pasta.iterdir():
            if item.is_dir() and item.stat().st_mtime < limite:
                shutil.rmtree(item, ignore_errors=True)


def registrar_estado(dados):
    ESTADO.parent.mkdir(parents=True, exist_ok=True)
    ESTADO.write_text(json.dumps(dados, ensure_ascii=False, indent=2), encoding="utf-8")
    atualizar_dashboard()


def registrar_historico(dados):
    HISTORICO.parent.mkdir(parents=True, exist_ok=True)
    with HISTORICO.open("a", encoding="utf-8") as arquivo:
        arquivo.write(json.dumps(dados, ensure_ascii=False) + "\n")
    RELATORIO.write_text(json.dumps(dados, ensure_ascii=False, indent=2), encoding="utf-8")
    atualizar_dashboard()


def mostrar_status():
    if not ESTADO.exists():
        print("Nenhuma execução registrada.")
        return 0
    estado = json.loads(ESTADO.read_text(encoding="utf-8"))
    print(f"Fluxo: {estado.get('fluxo', '-')}")
    print(f"Status: {estado.get('status', '-')}")
    print(f"Início: {estado.get('inicio', '-')}")
    print(f"Fim: {estado.get('fim', '-')}")
    print(f"Tentativa: {estado.get('tentativa', '-')}")
    print(f"Duração (min): {estado.get('duracao_minutos', '-')}")
    print(f"Código de saída: {estado.get('codigo_saida', '-')}")
    return 0


def executar_fluxo(nome, tentativas, config):
    comando = [sys.executable, *FLUXOS[nome]]
    estado_processo = SCRIPT_DIR / "logs" / "processo_estado.json"
    if nome == "ercard-carteira":
        estado_fluxo = SCRIPT_DIR / "logs" / "ercard_carteira_estado.json"
        if estado_fluxo.exists():
            try:
                estado_fluxo_dados = json.loads(estado_fluxo.read_text(encoding="utf-8"))
                if estado_fluxo_dados.get("status") == "erro":
                    etapa = int(estado_fluxo_dados.get("etapa", 1))
                    if etapa in (1, 2, 3):
                        comando.extend(["--a-partir", str(etapa)])
            except (OSError, ValueError, json.JSONDecodeError):
                pass
    inicio = time.monotonic()
    base = {
        "fluxo": nome,
        "inicio": agora(),
        "status": "executando",
        "tentativa": 0,
        "comando": comando,
    }
    registrar_estado(base)
    base["backups"] = criar_backups(config)
    limpar_backups_antigos(config)
    timeout_minutos = max(1, int(config.get("timeout_minutos", 180)))

    ultimo_codigo = 1
    for tentativa in range(1, tentativas + 1):
        if not verificar_internet():
            print("[AGENTE] Internet indisponível; aguardando a próxima tentativa.", flush=True)
            ultimo_codigo = 3
            time.sleep(min(30, 5 * tentativa))
            continue
        ponto_retomada = {}
        if estado_processo.exists():
            try:
                ponto_retomada = json.loads(estado_processo.read_text(encoding="utf-8"))
            except (OSError, json.JSONDecodeError):
                ponto_retomada = {}
        base["tentativa"] = tentativa
        base["etapa_retomada"] = ponto_retomada.get("etapa")
        registrar_estado(base)
        etapa = ponto_retomada.get("etapa", "inicializacao")
        print(
            f"[AGENTE] Iniciando {nome}; tentativa {tentativa}/{tentativas}; "
            f"retomada na etapa {etapa}",
            flush=True,
        )
        try:
            processo = subprocess.run(
                comando,
                cwd=SCRIPT_DIR,
                timeout=timeout_minutos * 60,
            )
        except subprocess.TimeoutExpired:
            ultimo_codigo = 124
            print(f"[AGENTE] {nome} excedeu o limite de {timeout_minutos} minutos.", flush=True)
            continue
        ultimo_codigo = processo.returncode
        if ultimo_codigo == 0:
            if not validar_resultado(nome):
                ultimo_codigo = 2
                print(f"[AGENTE] {nome} terminou sem validar o arquivo final", flush=True)
                continue
            duracao = (time.monotonic() - inicio) / 60
            resultado = {
                **base,
                "fim": agora(),
                "status": "concluido",
                "codigo_saida": 0,
                "duracao_minutos": round(duracao, 2),
            }
            registrar_estado(resultado)
            registrar_historico(resultado)
            notificar(config, f"Crediagora: fluxo {nome} concluído em {duracao:.2f} minutos.")
            print(f"[AGENTE] {nome} concluído em {duracao:.2f} minutos", flush=True)
            return 0
        print(f"[AGENTE] {nome} terminou com código {ultimo_codigo}", flush=True)

    duracao = (time.monotonic() - inicio) / 60
    resultado = {
        **base,
        "fim": agora(),
        "status": "erro",
        "codigo_saida": ultimo_codigo,
        "duracao_minutos": round(duracao, 2),
    }
    registrar_estado(resultado)
    registrar_historico(resultado)
    notificar(config, f"Crediagora: fluxo {nome} terminou com erro (código {ultimo_codigo}).")
    return ultimo_codigo or 1


def main(argv=None):
    parser = argparse.ArgumentParser(description="Agente de automação Crediagora")
    parser.add_argument("--fluxo", choices=sorted(FLUXOS))
    parser.add_argument("--tentativas", type=int, default=2)
    parser.add_argument("--config", type=Path, default=CONFIG_PADRAO)
    parser.add_argument("--status", action="store_true", help="Mostra o status da última execução.")
    args = parser.parse_args(argv)
    if args.status:
        return mostrar_status()
    if not args.fluxo:
        parser.error("informe --fluxo ou use --status")
    config = carregar_configuracao(args.config)
    tentativas = int(config.get("tentativas", args.tentativas))
    if tentativas < 1:
        parser.error("--tentativas deve ser maior que zero")
    validar_ambiente(args.fluxo)
    adquirir_trava()
    try:
        return executar_fluxo(args.fluxo, tentativas, config)
    finally:
        liberar_trava()


if __name__ == "__main__":
    raise SystemExit(main())
