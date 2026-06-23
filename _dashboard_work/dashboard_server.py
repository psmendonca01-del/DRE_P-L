import json
import shutil
import subprocess
import sys
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path


BASE = Path(__file__).resolve().parent
BUILD_SCRIPT = BASE / "build_dashboard.py"
SOURCE_DIR = Path("C:/Users/PauloMendonça/OneDrive - Redefrete/Área de Trabalho/Balanço/DashBoard")
SOURCE_COPIES = BASE / "source_copies"
STATUS_LOG = BASE / "server.status.log"
NOTES_FILE = BASE / "dashboard_notes.json"


def status(message):
    with STATUS_LOG.open("a", encoding="utf-8") as handle:
        handle.write(f"{message}\n")


class DashboardHandler(SimpleHTTPRequestHandler):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=str(BASE), **kwargs)

    def log_message(self, format, *args):
        status(format % args)

    def end_headers(self):
        self.send_header("Cache-Control", "no-store, no-cache, must-revalidate, max-age=0")
        self.send_header("Pragma", "no-cache")
        super().end_headers()

    def do_POST(self):
        if self.path.rstrip("/") == "/notes":
            try:
                length = int(self.headers.get("Content-Length", "0"))
                body = self.rfile.read(length)
                notes = json.loads(body.decode("utf-8") or "{}")
                if not isinstance(notes, dict):
                    raise ValueError("Notes payload must be an object")
                tmp = NOTES_FILE.with_suffix(".json.tmp")
                tmp.write_text(json.dumps(notes, ensure_ascii=False, indent=2), encoding="utf-8")
                tmp.replace(NOTES_FILE)
                payload = {"ok": True, "path": str(NOTES_FILE)}
                status(f"notes saved: {len(notes)} cards")
                response_body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
                self.send_response(200)
                self.send_header("Content-Type", "application/json; charset=utf-8")
                self.send_header("Content-Length", str(len(response_body)))
                self.end_headers()
                self.wfile.write(response_body)
            except Exception as error:
                payload = {"ok": False, "error": str(error)}
                response_body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
                self.send_response(500)
                self.send_header("Content-Type", "application/json; charset=utf-8")
                self.send_header("Content-Length", str(len(response_body)))
                self.end_headers()
                self.wfile.write(response_body)
            return

        if self.path.rstrip("/") != "/refresh":
            self.send_error(404, "Endpoint not found")
            return

        copied = []
        removed = []
        if SOURCE_DIR.exists():
            SOURCE_COPIES.mkdir(exist_ok=True)
            valid_sources = {
                source.name: source
                for source in sorted(SOURCE_DIR.glob("*.xlsx"))
                if not source.name.startswith("~$")
            }
            for target in sorted(SOURCE_COPIES.glob("*.xlsx")):
                if target.name not in valid_sources:
                    target.unlink()
                    removed.append(target.name)
            for source in valid_sources.values():
                target = SOURCE_COPIES / source.name
                shutil.copy2(source, target)
                copied.append(source.name)

        result = subprocess.run(
            [sys.executable, str(BUILD_SCRIPT)],
            cwd=str(BASE.parent),
            capture_output=True,
            text=True,
        )
        payload = {
            "ok": result.returncode == 0,
            "copied": copied,
            "removed": removed,
            "sourceDir": str(SOURCE_DIR),
            "stdout": result.stdout,
            "stderr": result.stderr,
        }
        body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        self.send_response(200 if result.returncode == 0 else 500)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)


def main():
    try:
        status("starting")
        server = ThreadingHTTPServer(("127.0.0.1", 8765), DashboardHandler)
        status("listening http://127.0.0.1:8765/dashboard_dre.html")
        print("Dashboard server running at http://127.0.0.1:8765/dashboard_dre.html", flush=True)
        server.serve_forever()
    except Exception as error:
        status(f"error: {error}")
        raise


if __name__ == "__main__":
    main()
