import json
import os
import sqlite3
import unicodedata
from pathlib import Path


BASE = Path(__file__).resolve().parent
DB_FILE = BASE / "pl_database.sqlite"


ROW_TABLES = {
    "unifiedRows": "unified_rows",
    "operationRows": "operation_rows",
    "budgetRows": "budget_rows",
}


TEXT_FIELDS = (
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
)


NUMERIC_FIELDS = (
    "value",
    "routes",
    "loaded",
    "delivered",
    "evidenced",
)

LEDGER_TEXT_FIELDS = TEXT_FIELDS + (
    "sourceRow",
    "party",
    "invoice",
    "rateioOriginProject",
    "rateioOriginDepartment",
    "lineKey",
    "accountLoose",
    "categoryLoose",
    "clientNorm",
    "projectNorm",
    "unitNorm",
    "exptNorm",
    "typeNorm",
    "fleetNorm",
)

LEDGER_NUMERIC_FIELDS = NUMERIC_FIELDS + ("originalValue",)


def norm(value):
    text = unicodedata.normalize("NFD", str(value or ""))
    text = "".join(ch for ch in text if unicodedata.category(ch) != "Mn")
    return "".join(ch for ch in text.upper() if ch.isalnum())


def loose_norm(value):
    return norm(value).replace("C", "")


def unit_value(row):
    return row.get("hub") or row.get("department") or row.get("sourceHub") or ""


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
    if (
        "DESPESASADMINISTRATIVAS" in account
        or "DESPESASCOMPESSOAL" in account
        or "VENDASEMARKETING" in account
        or "OUTROSTRIBUTOS" in account
    ):
        return "expensesTotal"
    if "RECEITASFINANCEIRAS" in account:
        return "financialResult"
    if "DESPESASFINANCEIRAS" in account:
        return "financialResult"
    return "other"


def _connect(db_file=DB_FILE):
    conn = sqlite3.connect(str(db_file))
    conn.execute("PRAGMA journal_mode=WAL")
    conn.execute("PRAGMA synchronous=NORMAL")
    conn.execute("PRAGMA temp_store=MEMORY")
    return conn


def _create_row_table(conn, table):
    columns = ["id INTEGER PRIMARY KEY"]
    columns.extend(f"{name} TEXT" for name in TEXT_FIELDS)
    columns.extend(f"{name} REAL" for name in NUMERIC_FIELDS)
    columns.append("payload TEXT NOT NULL")
    conn.execute(f"DROP TABLE IF EXISTS {table}")
    conn.execute(f"CREATE TABLE {table} ({', '.join(columns)})")


def _insert_rows(conn, table, rows):
    fields = list(TEXT_FIELDS) + list(NUMERIC_FIELDS) + ["payload"]
    placeholders = ",".join("?" for _ in fields)
    sql = f"INSERT INTO {table} ({', '.join(fields)}) VALUES ({placeholders})"
    payloads = []
    for row in rows:
        values = [row.get(name, "") for name in TEXT_FIELDS]
        values.extend(float(row.get(name) or 0) for name in NUMERIC_FIELDS)
        values.append(json.dumps(row, ensure_ascii=False, separators=(",", ":")))
        payloads.append(values)
    if payloads:
        conn.executemany(sql, payloads)


def _create_ledger_table(conn):
    columns = ["id INTEGER PRIMARY KEY"]
    columns.extend(f"{name} TEXT" for name in LEDGER_TEXT_FIELDS)
    columns.extend(f"{name} REAL" for name in LEDGER_NUMERIC_FIELDS)
    columns.append("payload TEXT NOT NULL")
    conn.execute("DROP TABLE IF EXISTS ledger_rows")
    conn.execute(f"CREATE TABLE ledger_rows ({', '.join(columns)})")


def _ledger_row(row):
    out = dict(row)
    out["lineKey"] = line_bucket(row)
    out["accountLoose"] = loose_norm(row.get("account"))
    out["categoryLoose"] = loose_norm(row.get("category"))
    out["clientNorm"] = norm(row.get("client"))
    out["projectNorm"] = norm(row.get("project"))
    out["unitNorm"] = norm(unit_value(row))
    out["exptNorm"] = norm(row.get("expt"))
    out["typeNorm"] = norm(row.get("vehicleType"))
    out["fleetNorm"] = norm(row.get("fleetType"))
    return out


def _insert_ledger_rows(conn, rows):
    fields = list(LEDGER_TEXT_FIELDS) + list(LEDGER_NUMERIC_FIELDS) + ["payload"]
    placeholders = ",".join("?" for _ in fields)
    sql = f"INSERT INTO ledger_rows ({', '.join(fields)}) VALUES ({placeholders})"
    payloads = []
    for row in rows:
        prepared = _ledger_row(row)
        values = [prepared.get(name, "") for name in LEDGER_TEXT_FIELDS]
        values.extend(float(prepared.get(name) or 0) for name in LEDGER_NUMERIC_FIELDS)
        values.append(json.dumps(row, ensure_ascii=False, separators=(",", ":")))
        payloads.append(values)
    if payloads:
        conn.executemany(sql, payloads)


def _create_indexes(conn, table):
    for fields in (
        ("period",),
        ("client",),
        ("project",),
        ("hub",),
        ("expt",),
        ("vehicleType",),
        ("fleetType",),
        ("account",),
        ("category",),
        ("period", "client", "project", "hub"),
    ):
        index_name = f"idx_{table}_{'_'.join(fields)}"
        conn.execute(f"CREATE INDEX {index_name} ON {table} ({', '.join(fields)})")


def _create_ledger_indexes(conn):
    for fields in (
        ("period",),
        ("lineKey",),
        ("accountLoose",),
        ("categoryLoose",),
        ("clientNorm",),
        ("projectNorm",),
        ("unitNorm",),
        ("exptNorm",),
        ("typeNorm",),
        ("fleetNorm",),
        ("period", "lineKey"),
        ("period", "lineKey", "accountLoose"),
        ("period", "lineKey", "accountLoose", "categoryLoose"),
        ("period", "clientNorm", "projectNorm", "unitNorm"),
    ):
        index_name = f"idx_ledger_rows_{'_'.join(fields)}"
        conn.execute(f"CREATE INDEX {index_name} ON ledger_rows ({', '.join(fields)})")


def _write_meta(conn, meta):
    conn.execute("DROP TABLE IF EXISTS meta")
    conn.execute("CREATE TABLE meta (key TEXT PRIMARY KEY, value TEXT)")
    conn.executemany(
        "INSERT INTO meta (key, value) VALUES (?, ?)",
        [(key, json.dumps(value, ensure_ascii=False)) for key, value in sorted(meta.items())],
    )


def write_database(data, ledger=None, db_file=DB_FILE):
    db_file = Path(db_file)
    db_file.parent.mkdir(parents=True, exist_ok=True)
    if os.environ.get("PL_DB_IN_PLACE") == "1" and db_file.exists():
        conn = _connect(db_file)
        try:
            _write_meta(conn, data.get("meta", {}))
            for data_key, table in ROW_TABLES.items():
                rows = data.get(data_key, [])
                _create_row_table(conn, table)
                _insert_rows(conn, table, rows)
                _create_indexes(conn, table)
            if ledger is not None:
                _create_ledger_table(conn)
                _insert_ledger_rows(conn, ledger)
                _create_ledger_indexes(conn)
            conn.commit()
        finally:
            conn.close()
        return db_file

    tmp_file = db_file.with_suffix(".sqlite.tmp")
    if tmp_file.exists():
        tmp_file.unlink()

    conn = _connect(tmp_file)
    try:
        _write_meta(conn, data.get("meta", {}))
        for data_key, table in ROW_TABLES.items():
            rows = data.get(data_key, [])
            _create_row_table(conn, table)
            _insert_rows(conn, table, rows)
            _create_indexes(conn, table)
        if ledger is not None:
            _create_ledger_table(conn)
            _insert_ledger_rows(conn, ledger)
            _create_ledger_indexes(conn)
        conn.commit()
    finally:
        conn.close()

    if db_file.exists():
        db_file.unlink()
    tmp_file.replace(db_file)
    return db_file


def update_budget_rows(data, db_file=DB_FILE):
    db_file = Path(db_file)
    if not db_file.exists():
        return write_database(data, db_file)

    conn = _connect(db_file)
    try:
        _write_meta(conn, data.get("meta", {}))
        table = ROW_TABLES["budgetRows"]
        _create_row_table(conn, table)
        _insert_rows(conn, table, data.get("budgetRows", []))
        _create_indexes(conn, table)
        conn.commit()
    finally:
        conn.close()
    return db_file
