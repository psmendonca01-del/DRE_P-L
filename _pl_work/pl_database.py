import json
import sqlite3
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


def _write_meta(conn, meta):
    conn.execute("DROP TABLE IF EXISTS meta")
    conn.execute("CREATE TABLE meta (key TEXT PRIMARY KEY, value TEXT)")
    conn.executemany(
        "INSERT INTO meta (key, value) VALUES (?, ?)",
        [(key, json.dumps(value, ensure_ascii=False)) for key, value in sorted(meta.items())],
    )


def write_database(data, db_file=DB_FILE):
    db_file = Path(db_file)
    db_file.parent.mkdir(parents=True, exist_ok=True)
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
