"""Gerador EasyBuilder Pro — G3 Automation.

Interface moderna em janela própria (Edge/Chrome em modo app) servida localmente.
Sem dependências externas. A geração dos arquivos está em gerador_core.py (lógica validada).
"""
from __future__ import annotations

import argparse
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import threading
import time
import webbrowser
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import unquote

import gerador_core as core

BASE = Path(__file__).resolve().parent
ROOT = BASE.parent
TEMPLATE = ROOT / "modelos" / "modelo_database.db"
STATE = {"xlsx": None, "ping": time.time(), "bye": 0.0}

PICK_CODE = (
    "import sys,tkinter as tk\nfrom tkinter import filedialog\n"
    "r=tk.Tk();r.withdraw();r.attributes('-topmost',True)\n"
    "p=filedialog.askdirectory(initialdir=sys.argv[2],title='Pasta de saída')\n"
    "open(sys.argv[1],'w',encoding='utf-8').write(p or '')\n"
)


def categorize(item: dict) -> str:
    s = item["source"]
    if s in ("Receita", "ID"):
        return "base"
    if item.get("kind") == "bool":
        return "bool"
    if re.match(r"^[^.]+\.Cota_\d+", s, re.I):
        return "cota"
    if s.startswith("IHM_RET.Tempos."):
        return "tempo"
    return "param"


def analyze(path: Path) -> list[dict]:
    try:
        rows, *_ = core.read_xlsx(path)
        items = core.source_items(rows)
    except Exception:
        raise ValueError("Não foi possível ler o arquivo. Use o .xlsx exportado da declaração de variáveis (Sysmac).")
    if not items:
        raise ValueError("Nenhuma variável de receita foi encontrada nesta planilha.")
    return [
        {"n": i, "name": it["source"], "type": "BOOL → WORD" if it.get("kind") == "bool" else it["type"],
         "cat": categorize(it), "rw": it.get("rw_base")}
        for i, it in enumerate(items, 1)
    ]


def pick_folder(initial: str) -> str:
    tmp = Path(tempfile.gettempdir()) / "g3_pick.txt"
    tmp.unlink(missing_ok=True)
    start = initial if Path(initial).is_dir() else str(Path.home())
    subprocess.run([sys.executable, "-c", PICK_CODE, str(tmp), start], timeout=300)
    return tmp.read_text(encoding="utf-8").strip() if tmp.exists() else ""


def open_folder(path: str) -> None:
    if os.name == "nt":
        os.startfile(path)  # type: ignore[attr-defined]
    else:
        subprocess.Popen(["open" if sys.platform == "darwin" else "xdg-open", path])


class Handler(BaseHTTPRequestHandler):
    def log_message(self, *args):  # sem console (pythonw)
        pass

    def reply(self, data, code=200, ctype="application/json"):
        body = data if isinstance(data, bytes) else json.dumps(data, ensure_ascii=False).encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", ctype + "; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        self.wfile.write(body)

    def do_GET(self):
        if self.path == "/":
            self.reply((ROOT / "ui" / "index.html").read_bytes(), ctype="text/html")
        elif self.path == "/api/ping":
            STATE["ping"], STATE["bye"] = time.time(), 0.0
            self.reply({"ok": True})
        elif self.path == "/api/config":
            self.reply({"out": str(Path.home() / "Downloads")})
        else:
            self.reply({"error": "não encontrado"}, 404)

    def do_POST(self):
        size = int(self.headers.get("Content-Length") or 0)
        body = self.rfile.read(size) if size else b""
        try:
            if self.path == "/api/bye":
                STATE["bye"] = time.time()
                return self.reply({"ok": True})
            if self.path == "/api/analyze":
                tmp = Path(tempfile.gettempdir()) / "g3_planilha.xlsx"
                tmp.write_bytes(body)
                items = analyze(tmp)
                STATE["xlsx"] = tmp
                return self.reply({"items": items})
            data = json.loads(body or b"{}")
            if self.path == "/api/pick":
                return self.reply({"path": pick_folder(data.get("path", ""))})
            if self.path == "/api/open":
                if Path(data.get("path", "")).is_dir():
                    open_folder(data["path"])
                return self.reply({"ok": True})
            if self.path == "/api/generate":
                if not STATE["xlsx"]:
                    raise ValueError("Selecione a planilha de variáveis primeiro.")
                out = Path(data.get("out") or "").expanduser()
                if not str(out).strip() or str(out) == ".":
                    raise ValueError("Informe a pasta de saída.")
                db, xlsx, count, macros = core.generate(STATE["xlsx"], TEMPLATE, out)
                files = [Path(p) for p in (db, xlsx, *macros)]
                return self.reply({"count": count, "out": str(out),
                                   "files": [{"name": f.name, "size": f.stat().st_size} for f in files]})
            self.reply({"error": "não encontrado"}, 404)
        except Exception as exc:
            self.reply({"error": str(exc)}, 400)


def find_browser() -> str | None:
    cands = []
    if os.name == "nt":
        for env in ("PROGRAMFILES(X86)", "PROGRAMFILES", "LOCALAPPDATA"):
            root = os.environ.get(env)
            if root:
                cands += [Path(root) / "Microsoft/Edge/Application/msedge.exe",
                          Path(root) / "Google/Chrome/Application/chrome.exe"]
    for name in ("msedge", "microsoft-edge", "chrome", "google-chrome", "chromium", "chromium-browser"):
        found = shutil.which(name)
        if found:
            cands.append(Path(found))
    return next((str(c) for c in cands if Path(c).exists()), None)


def watchdog():
    """Encerra o servidor quando a janela do app é fechada."""
    while True:
        time.sleep(1)
        now = time.time()
        if (STATE["bye"] and now - STATE["bye"] > 4) or now - STATE["ping"] > 120:
            os._exit(0)


def main():
    ap = argparse.ArgumentParser(description="Gerador EasyBuilder Pro — G3 Automation")
    ap.add_argument("--input", type=Path)
    ap.add_argument("--template", type=Path, default=TEMPLATE)
    ap.add_argument("--output", type=Path)
    ap.add_argument("--no-window", action="store_true", help="apenas sobe o servidor (teste)")
    args = ap.parse_args()
    if args.input and args.output:  # modo sem interface (igual à versão anterior)
        db, xlsx, count, macros = core.generate(args.input, args.template, args.output)
        print(f"OK {count} itens\n{db}\n{xlsx}")
        for m in macros:
            print(m)
        return
    server = ThreadingHTTPServer(("127.0.0.1", 0), Handler)
    url = f"http://127.0.0.1:{server.server_address[1]}/"
    if args.no_window:
        print(url, flush=True)
    else:
        threading.Thread(target=watchdog, daemon=True).start()
        exe = find_browser()
        if exe:
            subprocess.Popen([exe, f"--app={url}", "--window-size=1180,800", "--no-first-run",
                              "--no-default-browser-check", f"--user-data-dir={BASE / '.perfil'}"])
        else:
            webbrowser.open(url)
    server.serve_forever()


if __name__ == "__main__":
    try:
        main()
    except Exception as exc:  # pythonw não tem console: registra em arquivo
        (BASE / "erro.log").write_text(repr(exc), encoding="utf-8")
        raise
