"""
Database layer — tự động chọn:
  - SQLite  khi chạy local (không có DATABASE_URL)
  - PostgreSQL khi chạy trên Render/cloud (có DATABASE_URL)
"""
import os, json
from datetime import datetime

# Load .env nếu có (local dev)
try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass

DATABASE_URL = os.getenv("DATABASE_URL", "")
USE_POSTGRES  = DATABASE_URL.startswith("postgres")

# ── SQLite helpers ────────────────────────────────────────────────
def _sqlite_conn():
    import sqlite3
    db_path = os.path.join(os.path.dirname(__file__), "..", "machine_log.db")
    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    return conn

# ── PostgreSQL helpers ────────────────────────────────────────────
def _pg_conn():
    import psycopg2, psycopg2.extras
    url = DATABASE_URL
    # Render đôi khi trả "postgres://" — psycopg2 cần "postgresql://"
    if url.startswith("postgres://"):
        url = "postgresql://" + url[len("postgres://"):]
    conn = psycopg2.connect(url, cursor_factory=psycopg2.extras.RealDictCursor)
    return conn

def _get_conn():
    return _pg_conn() if USE_POSTGRES else _sqlite_conn()

# ── Placeholder khác nhau giữa SQLite và PostgreSQL ───────────────
PH = "%s" if USE_POSTGRES else "?"   # SQLite dùng ?, Postgres dùng %s

# ── CREATE TABLE ──────────────────────────────────────────────────
_CREATE = """
CREATE TABLE IF NOT EXISTS logs (
    id                      SERIAL PRIMARY KEY,
    created_at              TIMESTAMP DEFAULT NOW(),
    project                 TEXT    NOT NULL,
    mo_color                TEXT,
    date                    TEXT    NOT NULL,
    machine_line            TEXT,
    sam                     REAL,
    standard_working_time   REAL,
    hc_operators            REAL,
    actual_output_day_hc    REAL,
    actual_output_day_mc    REAL,
    auto_output_h_hc        REAL,
    actual_output_h_mc      REAL,
    manual_output_day_hc    REAL,
    manual_output_h_hc      REAL,
    machine_working_time    REAL,
    changeover_time         REAL,
    breakdown_time          REAL,
    idle_time               REAL,
    defects_json            TEXT,
    target_output_day       REAL,
    auto_target_h_hc        REAL,
    eff_pct                 REAL,
    mc_utilization_pct      REAL,
    total_time_min          REAL,
    total_defect            REAL,
    defect_rate_pct         REAL
)
"""

_CREATE_SQLITE = _CREATE.replace("SERIAL PRIMARY KEY", "INTEGER PRIMARY KEY AUTOINCREMENT") \
                         .replace("TIMESTAMP DEFAULT NOW()", "TEXT DEFAULT (datetime('now'))")


def init_db():
    conn = _get_conn()
    try:
        cur = conn.cursor()
        sql = _CREATE if USE_POSTGRES else _CREATE_SQLITE
        cur.execute(sql)
        conn.commit()
    finally:
        conn.close()


# ── INSERT ────────────────────────────────────────────────────────
_INSERT = f"""
    INSERT INTO logs (
        project, mo_color, date, machine_line, sam,
        standard_working_time, hc_operators,
        actual_output_day_hc, actual_output_day_mc,
        auto_output_h_hc, actual_output_h_mc,
        manual_output_day_hc, manual_output_h_hc,
        machine_working_time, changeover_time,
        breakdown_time, idle_time, defects_json,
        target_output_day, auto_target_h_hc, eff_pct,
        mc_utilization_pct, total_time_min,
        total_defect, defect_rate_pct
    ) VALUES ({",".join([PH]*25)})
"""
_INSERT_PG = _INSERT + " RETURNING id"


def save_record(entry, computed: dict) -> int:
    params = (
        entry.project, entry.mo_color, entry.date, entry.machine_line,
        entry.sam, entry.standard_working_time, entry.hc_operators,
        entry.actual_output_day_hc, entry.actual_output_day_mc,
        entry.auto_output_h_hc, entry.actual_output_h_mc,
        entry.manual_output_day_hc, entry.manual_output_h_hc,
        entry.machine_working_time, entry.changeover_time,
        entry.breakdown_time, entry.idle_time,
        json.dumps([d.dict() for d in entry.defects]),
        computed["target_output_day"], computed["auto_target_h_hc"],
        computed["eff_pct"], computed["mc_utilization_pct"],
        computed["total_time_min"], computed["total_defect"],
        computed["defect_rate_pct"],
    )
    conn = _get_conn()
    try:
        cur = conn.cursor()
        if USE_POSTGRES:
            cur.execute(_INSERT_PG, params)
            row_id = cur.fetchone()["id"]
        else:
            cur.execute(_INSERT, params)
            row_id = cur.lastrowid
        conn.commit()
        return row_id
    finally:
        conn.close()


# ── SELECT ────────────────────────────────────────────────────────
def get_records(project=None, record_id=None, limit=200):
    conn = _get_conn()
    try:
        cur = conn.cursor()
        if record_id:
            cur.execute(f"SELECT * FROM logs WHERE id={PH}", (record_id,))
        elif project:
            cur.execute(
                f"SELECT * FROM logs WHERE project={PH} ORDER BY date DESC LIMIT {PH}",
                (project, limit)
            )
        else:
            cur.execute(f"SELECT * FROM logs ORDER BY date DESC LIMIT {PH}", (limit,))

        rows = cur.fetchall()
        result = []
        for row in rows:
            d = dict(row)
            d["defects"] = json.loads(d.get("defects_json") or "[]")
            del d["defects_json"]
            # Chuyển datetime object → string nếu cần
            if isinstance(d.get("created_at"), datetime):
                d["created_at"] = d["created_at"].isoformat()
            result.append(d)
        return result
    finally:
        conn.close()
