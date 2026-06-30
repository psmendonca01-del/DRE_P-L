import json
import sqlite3
import subprocess
import sys
import time
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import parse_qs, urlparse


BASE = Path(__file__).resolve().parent
BUILD_SCRIPT = BASE / "build_pl_dashboard.py"
SOURCE_FILE = Path(r"C:\Users\PauloMendonça\OneDrive - Redefrete\Área de Trabalho\Balanço\DashBoard_P&L\P&L.xlsx")
BUDGET_FILE = Path(r"C:\Users\PauloMendonça\OneDrive - Redefrete\Área de Trabalho\Balanço\DashBoard_P&L\Budget.xlsx")
WORKBOOK = BASE / "PL.xlsx"
DATA_FILE = BASE / "pl_data.json"
LEDGER_FILE = BASE / "pl_ledger.json"
DB_FILE = BASE / "pl_database.sqlite"
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


def read_ledger_from_db(payload):
    if not DB_FILE.exists():
        return None
    filters = payload.get("filters", {})
    periods = set(filters.get("months") or [])
    line_key = payload.get("lineKey")
    account = payload.get("account")
    category = payload.get("category")
    where = []
    args = []
    if periods:
        where.append(f"period IN ({','.join('?' for _ in periods)})")
        args.extend(sorted(periods))
    if line_key:
        where.append("lineKey = ?")
        args.append(line_key)
    if account:
        where.append("accountLoose = ?")
        args.append(loose_norm(account))
    if category:
        where.append("categoryLoose = ?")
        args.append(loose_norm(category))
    filter_map = (
        ("client", "clientNorm"),
        ("project", "projectNorm"),
        ("unit", "unitNorm"),
        ("expt", "exptNorm"),
        ("type", "typeNorm"),
        ("fleet", "fleetNorm"),
    )
    for filter_name, column in filter_map:
        value = filters.get(filter_name)
        if not is_all_filter(value):
            where.append(f"{column} = ?")
            args.append(norm(value))
    group_fields = (
        "period",
        "source",
        "sourceRow",
        "party",
        "invoice",
        "account",
        "category",
        "client",
        "project",
        "hub",
        "expt",
        "vehicleType",
        "fleetType",
        "originalValue",
    )
    sql = (
        f"SELECT {', '.join(group_fields)}, SUM(value) AS value "
        "FROM ledger_rows"
    )
    if where:
        sql += " WHERE " + " AND ".join(where)
    sql += f" GROUP BY {', '.join(group_fields)}"
    rows = []
    with sqlite3.connect(str(DB_FILE)) as conn:
        conn.row_factory = sqlite3.Row
        for row in conn.execute(sql, args):
            rows.append(dict(row))
    return rows


def grouped_ledger_payload(rows):
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
    return sorted(grouped.values(), key=lambda item: (str(item.get("period", "")), str(item.get("source", "")), str(item.get("invoice", ""))))


def _decode_meta_value(value):
    try:
        return json.loads(value)
    except Exception:
        return value


def _compact_row(row):
    out = {}
    for key, value in dict(row).items():
        if key == "id":
            continue
        if value in (None, ""):
            continue
        if isinstance(value, (int, float)) and abs(value) < 0.0001:
            continue
        out[key] = value
    return out


def read_data_from_db():
    if not DB_FILE.exists():
        return None
    with sqlite3.connect(str(DB_FILE)) as conn:
        conn.row_factory = sqlite3.Row
        meta = {
            row["key"]: _decode_meta_value(row["value"])
            for row in conn.execute("SELECT key, value FROM meta")
        }
        tables = {
            "unifiedRows": "unified_rows",
            "operationRows": "operation_rows",
            "budgetRows": "budget_rows",
        }
        data = {"meta": meta, "finance": [], "operations": []}
        row_fields = (
            "source",
            "period",
            "tipo",
            "grupo",
            "account",
            "category",
            "client",
            "rateio",
            "project",
            "sourceHub",
            "hub",
            "expt",
            "city",
            "department",
            "vehicleType",
            "fleetType",
            "fleetOwner",
            "costType",
            "scenario",
            "campaign",
            "value",
            "routes",
            "loaded",
            "delivered",
            "evidenced",
        )
        sql = f"SELECT {', '.join(row_fields)} FROM {{table}}"
        for data_key, table in tables.items():
            data[data_key] = [
                _compact_row(row)
                for row in conn.execute(sql.format(table=table))
            ]
        return data


def status(message):
    try:
        with STATUS_LOG.open("a", encoding="utf-8") as handle:
            handle.write(f"{message}\n")
    except OSError:
        pass


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
        parsed = urlparse(self.path)
        path = parsed.path.rstrip("/")
        if path == "/notes":
            self.save_notes()
            return
        if path == "/ledger":
            self.ledger_payload()
            return
        if path != "/refresh":
            self.send_error(404, "Endpoint not found")
            return
        payload = self.refresh_payload(parse_qs(parsed.query))
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
            raw_rows = read_ledger_from_db(payload)
            if raw_rows is None:
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
                raw_rows = []
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
                    raw_rows.append(row)
            rows = grouped_ledger_payload(raw_rows)
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
        path = self.path.split("?", 1)[0].rstrip("/")
        if path == "/data":
            try:
                payload = read_data_from_db()
                if payload is None:
                    payload = json.loads(DATA_FILE.read_text(encoding="utf-8"))
                body = json.dumps(payload, ensure_ascii=False, separators=(",", ":")).encode("utf-8")
                self.send_response(200)
            except Exception as exc:
                status(f"data read error: {exc!r}")
                body = json.dumps({"ok": False, "error": repr(exc)}, ensure_ascii=False).encode("utf-8")
                self.send_response(500)
            self.send_header("Content-Type", "application/json; charset=utf-8")
            self.send_header("Content-Length", str(len(body)))
            self.end_headers()
            self.wfile.write(body)
            return
        if path == "/notes":
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

    def refresh_budget_payload(self):
        started = time.perf_counter()
        try:
            if not DATA_FILE.exists():
                return {
                    "ok": False,
                    "copied": False,
                    "budgetOnly": True,
                    "sourceFile": str(BUDGET_FILE),
                    "stdout": "",
                    "stderr": "pl_data.json não encontrado. Faça uma atualização completa primeiro.",
                }
            if not BUDGET_FILE.exists():
                return {
                    "ok": False,
                    "copied": False,
                    "budgetOnly": True,
                    "sourceFile": str(BUDGET_FILE),
                    "stdout": "",
                    "stderr": f"Budget não encontrado: {BUDGET_FILE}",
                }

            sys.path.insert(0, str(BASE))
            import build_pl_dashboard as builder
            import pl_database

            data = json.loads(DATA_FILE.read_text(encoding="utf-8"))
            actual_periods = data.get("meta", {}).get("periods", [])
            budget_rows, budget_campaigns = builder.load_budget_rows(actual_periods)
            budget_periods = sorted({item["period"] for item in budget_rows})
            budget_scenarios = sorted({item.get("scenario") for item in budget_rows if item.get("scenario")})

            data["budgetRows"] = budget_rows
            meta = data.setdefault("meta", {})
            meta["budgetPeriods"] = budget_periods
            meta["budgetScenarios"] = budget_scenarios
            meta["budgetRows"] = len(budget_rows)
            meta["budgetCampaigns"] = len(budget_campaigns)
            meta["budgetSourceFile"] = str(BUDGET_FILE)
            meta["budgetSourceMtimeNs"] = BUDGET_FILE.stat().st_mtime_ns
            meta["budgetRefreshedAt"] = time.strftime("%Y-%m-%d %H:%M:%S")

            DATA_FILE.write_text(
                json.dumps(data, ensure_ascii=False, separators=(",", ":")),
                encoding="utf-8",
            )
            pl_database.update_budget_rows(data)
            elapsed = time.perf_counter() - started
            return {
                "ok": True,
                "copied": False,
                "budgetOnly": True,
                "sourceFile": str(BUDGET_FILE),
                "stdout": f"Budget atualizado: {len(budget_rows)} linhas em {elapsed:.1f}s.",
                "stderr": "",
            }
        except Exception as exc:
            status(f"budget refresh error: {exc!r}")
            return {
                "ok": False,
                "copied": False,
                "budgetOnly": True,
                "sourceFile": str(BUDGET_FILE),
                "stdout": "",
                "stderr": repr(exc),
            }

    def refresh_payload(self, params=None):
        params = params or {}
        if (params.get("view") or [""])[0] == "budget":
            return self.refresh_budget_payload()

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
