import json
import subprocess
import sys
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path


BASE = Path(__file__).resolve().parent
BUILD_SCRIPT = BASE / "build_pl_dashboard.py"
SOURCE_FILE = Path(r"C:\Users\PauloMendonça\OneDrive - Redefrete\Área de Trabalho\Balanço\DashBoard_P&L\P&L.xlsx")
WORKBOOK = BASE / "PL.xlsx"
DATA_FILE = BASE / "pl_data.json"
NOTES_FILE = BASE / "pl_notes.json"
STATUS_LOG = BASE / "pl_server.status.log"


def status(message):
    with STATUS_LOG.open("a", encoding="utf-8") as handle:
        handle.write(f"{message}\n")


class PLHandler(SimpleHTTPRequestHandler):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=str(BASE), **kwargs)

    def log_message(self, format, *args):
        status(format % args)

    def end_headers(self):
        self.send_header("Cache-Control", "no-store, no-cache, must-revalidate, max-age=0")
        self.send_header("Pragma", "no-cache")
        super().end_headers()

    def do_POST(self):
        path = self.path.rstrip("/")
        if path == "/notes":
            self.save_notes()
            return
        if path != "/refresh":
            self.send_error(404, "Endpoint not found")
            return
        payload = self.refresh_payload()
        body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        self.send_response(200 if payload.get("ok") else 500)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def do_GET(self):
        if self.path.split("?", 1)[0].rstrip("/") == "/notes":
            payload = {}
            try:
                if NOTES_FILE.exists():
                    payload = json.loads(NOTES_FILE.read_text(encoding="utf-8"))
                    if not isinstance(payload, dict):
                        payload = {}
            except Exception as exc:
                status(f"notes read error: {exc!r}")
                payload = {}
            body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
            self.send_response(200)
            self.send_header("Content-Type", "application/json; charset=utf-8")
            self.send_header("Content-Length", str(len(body)))
            self.end_headers()
            self.wfile.write(body)
            return
        super().do_GET()

    def save_notes(self):
        try:
            length = int(self.headers.get("Content-Length", "0"))
            raw = self.rfile.read(length).decode("utf-8") if length else "{}"
            payload = json.loads(raw or "{}")
            if not isinstance(payload, dict):
                raise ValueError("Notes payload must be an object")
            NOTES_FILE.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
            body = json.dumps({"ok": True}, ensure_ascii=False).encode("utf-8")
            self.send_response(200)
        except Exception as exc:
            status(f"notes save error: {exc!r}")
            body = json.dumps({"ok": False, "error": repr(exc)}, ensure_ascii=False).encode("utf-8")
            self.send_response(500)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def refresh_payload(self):
        copied = False
        try:
            if SOURCE_FILE.exists() and DATA_FILE.exists():
                try:
                    current_mtime = SOURCE_FILE.stat().st_mtime_ns
                    meta = json.loads(DATA_FILE.read_text(encoding="utf-8")).get("meta", {})
                    if meta.get("sourceMtimeNs") == current_mtime:
                        return {
                            "ok": True,
                            "copied": False,
                            "skipped": True,
                            "sourceFile": str(SOURCE_FILE),
                            "stdout": "Base já estava atualizada.",
                            "stderr": "",
                        }
                except Exception as exc:
                    status(f"refresh cache check ignored: {exc!r}")
            result = subprocess.run(
                [sys.executable, str(BUILD_SCRIPT)],
                cwd=str(BASE.parent),
                capture_output=True,
                text=True,
            )
            return {
                "ok": result.returncode == 0,
                "copied": copied,
                "sourceFile": str(SOURCE_FILE),
                "stdout": result.stdout,
                "stderr": result.stderr,
            }
        except Exception as exc:
            status(f"refresh error: {exc!r}")
            return {
                "ok": False,
                "copied": copied,
                "sourceFile": str(SOURCE_FILE),
                "stdout": "",
                "stderr": repr(exc),
            }


def main():
    status("starting")
    server = ThreadingHTTPServer(("127.0.0.1", 8775), PLHandler)
    status("listening http://127.0.0.1:8775/dashboard_pl.html")
    print("P&L server running at http://127.0.0.1:8775/dashboard_pl.html", flush=True)
    server.serve_forever()


if __name__ == "__main__":
    main()
