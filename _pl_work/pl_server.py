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
LEDGER_FILE = BASE / "pl_ledger.json"
NOTES_FILE = BASE / "pl_notes.json"
STATUS_LOG = BASE / "pl_server.status.log"
LEDGER_CACHE = {"mtime": None, "rows": []}


def norm(value):
    import unicodedata

    text = unicodedata.normalize("NFD", str(value or ""))
    text = "".join(ch for ch in text if unicodedata.category(ch) != "Mn")
    return "".join(ch for ch in text.upper() if ch.isalnum())


def display(value):
    return str(value or "")


def unit_value(row):
    return row.get("hub") or row.get("department") or row.get("sourceHub") or ""


def is_all_filter(value):
    return value in (None, "") or norm(value) in {"ALL", "TODOS"}


def same_token(left, right):
    left_norm = norm(left)
    right_norm = norm(right)
    if left_norm == right_norm:
        return True
    return left_norm.replace("C", "") == right_norm.replace("C", "")


def loose_norm(value):
    return norm(value).replace("C", "")


def line_bucket(row):
    account = norm(row.get("account"))
    category = norm(row.get("category"))
    cost_type = norm(row.get("costType"))
    if "RENDIMENTOSDEAPLICACOES" in account or "RENDIMENTOSDEAPLICACOES" in category:
        return "financialRevenue"
    if "DEPRECIACAO" in account or "DEPRECIACAO" in category:
        return "depreciation"
    if "IRRF" in account or "IRRF" in category:
        return "resultTaxes"
    if "IRPJ" in account or "IRPJ" in category:
        return "resultTaxes"
    if "CSLL" in account or "CSLL" in category:
        return "resultTaxes"
    if "RECEITABRUTA" in account:
        return "gross"
    if "IMPOSTOS" in account or "DEDUCOES" in account:
        return "deductions"
    if "CUSTODOSSERVICOS" in account:
        return "costsVariable" if "VARIAVEL" in cost_type else "costsFixed"
    if "OUTRASRECEITAS" in account:
        return "costsFixed"
    if "DESPESASADMINISTRATIVAS" in account or "DESPESASCOMPESSOAL" in account or "VENDASEMARKETING" in account or "OUTROSTRIBUTOS" in account:
        return "expensesTotal"
    if "RECEITASFINANCEIRAS" in account:
        return "financialResult"
    if "DESPESASFINANCEIRAS" in account:
        return "financialResult"
    return "other"


def load_ledger():
    try:
        mtime = LEDGER_FILE.stat().st_mtime_ns
        if LEDGER_CACHE["mtime"] != mtime:
            rows = json.loads(LEDGER_FILE.read_text(encoding="utf-8"))
            for row in rows:
                row["_line"] = line_bucket(row)
                row["_accountLoose"] = loose_norm(row.get("account"))
                row["_categoryLoose"] = loose_norm(row.get("category"))
                row["_clientNorm"] = norm(row.get("client"))
                row["_projectNorm"] = norm(row.get("project"))
                row["_unitNorm"] = norm(unit_value(row))
                row["_exptNorm"] = norm(row.get("expt"))
                row["_typeNorm"] = norm(row.get("vehicleType"))
                row["_fleetNorm"] = norm(row.get("fleetType"))
            LEDGER_CACHE["rows"] = rows
            LEDGER_CACHE["mtime"] = mtime
    except Exception as exc:
        status(f"ledger load error: {exc!r}")
        LEDGER_CACHE["rows"] = []
        LEDGER_CACHE["mtime"] = None
    return LEDGER_CACHE["rows"]


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
        if path == "/ledger":
            self.ledger_payload()
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

    def ledger_payload(self):
        try:
            length = int(self.headers.get("Content-Length", "0"))
            raw = self.rfile.read(length).decode("utf-8") if length else "{}"
            payload = json.loads(raw or "{}")
            filters = payload.get("filters", {})
            periods = set(filters.get("months") or [])
            line_key = payload.get("lineKey")
            account = payload.get("account")
            category = payload.get("category")
            account_loose = loose_norm(account)
            category_loose = loose_norm(category)
            filter_norms = {
                "client": norm(filters.get("client")),
                "project": norm(filters.get("project")),
                "unit": norm(filters.get("unit")),
                "expt": norm(filters.get("expt")),
                "type": norm(filters.get("type")),
                "fleet": norm(filters.get("fleet")),
            }
            rows = []
            for row in load_ledger():
                if periods and row.get("period") not in periods:
                    continue
                if line_key and row.get("_line") != line_key:
                    continue
                if account and row.get("_accountLoose") != account_loose:
                    continue
                if category and row.get("_categoryLoose") != category_loose:
                    continue
                if not is_all_filter(filters.get("client")) and row.get("_clientNorm") != filter_norms["client"]:
                    continue
                if not is_all_filter(filters.get("project")) and row.get("_projectNorm") != filter_norms["project"]:
                    continue
                if not is_all_filter(filters.get("unit")) and row.get("_unitNorm") != filter_norms["unit"]:
                    continue
                if not is_all_filter(filters.get("expt")) and row.get("_exptNorm") != filter_norms["expt"]:
                    continue
                if not is_all_filter(filters.get("type")) and row.get("_typeNorm") != filter_norms["type"]:
                    continue
                if not is_all_filter(filters.get("fleet")) and row.get("_fleetNorm") != filter_norms["fleet"]:
                    continue
                rows.append(row)
            grouped = {}
            for row in rows:
                key = (
                    row.get("period", ""),
                    row.get("source", ""),
                    row.get("sourceRow", ""),
                    row.get("party", ""),
                    row.get("invoice", ""),
                    row.get("account", ""),
                    row.get("category", ""),
                    row.get("originalValue", ""),
                )
                if key not in grouped:
                    grouped[key] = {
                        "period": row.get("period", ""),
                        "source": row.get("source", ""),
                        "party": row.get("party") or row.get("client") or "",
                        "invoice": row.get("invoice", ""),
                        "account": row.get("account", ""),
                        "category": row.get("category", ""),
                        "client": row.get("client", ""),
                        "project": row.get("project", ""),
                        "hub": row.get("hub", ""),
                        "expt": row.get("expt", ""),
                        "vehicleType": row.get("vehicleType", ""),
                        "fleetType": row.get("fleetType", ""),
                        "originalValue": float(row.get("originalValue") or row.get("value") or 0),
                        "value": 0.0,
                    }
                grouped[key]["value"] += float(row.get("value") or 0)
            rows = sorted(grouped.values(), key=lambda item: (str(item.get("period", "")), str(item.get("source", "")), str(item.get("invoice", ""))))
            original = sum(float(row.get("originalValue") or row.get("value") or 0) for row in rows)
            absorbed = sum(float(row.get("value") or 0) for row in rows)
            def common(key, label):
                values = sorted({display(row.get(key)) for row in rows if display(row.get(key))})
                if len(values) == 1:
                    return f"{label}: {values[0]}"
                if len(values) > 1:
                    return f"{label}: Diversos"
                return ""
            units = sorted({display(unit_value(row)) for row in rows if display(unit_value(row))})
            context_parts = [
                common("account", "Conta DRE"),
                common("category", "Categoria"),
                common("client", "Cliente"),
                common("project", "Projeto"),
                f"Unidade/Expt: {units[0]}" if len(units) == 1 else ("Unidade/Expt: Diversos" if len(units) > 1 else ""),
                common("vehicleType", "Tipo"),
                common("fleetType", "Frota"),
            ]
            body = json.dumps(
                {
                    "ok": True,
                    "rows": rows[:5000],
                    "count": len(rows),
                    "original": original,
                    "absorbed": absorbed,
                    "context": " | ".join(part for part in context_parts if part),
                },
                ensure_ascii=False,
                separators=(",", ":"),
            ).encode("utf-8")
            self.send_response(200)
        except Exception as exc:
            status(f"ledger error: {exc!r}")
            body = json.dumps({"ok": False, "error": repr(exc)}, ensure_ascii=False).encode("utf-8")
            self.send_response(500)
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
    server = ThreadingHTTPServer(("0.0.0.0", 8775), PLHandler)
    status("listening http://0.0.0.0:8775/dashboard_pl.html")
    print("P&L server running at http://0.0.0.0:8775/dashboard_pl.html", flush=True)
    server.serve_forever()


if __name__ == "__main__":
    main()
