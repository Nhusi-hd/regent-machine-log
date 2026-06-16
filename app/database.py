"""
Database layer — tự động chọn:
  - SQLite  khi chạy local (không có DATABASE_URL)
  - PostgreSQL khi chạy trên Render/cloud (có DATABASE_URL)

Logic UPSERT: nếu đã có record cùng (project + date) → UPDATE thay vì INSERT mới.
Đảm bảo mỗi (project, date) chỉ có đúng 1 record trong DB.
"""
import os, json
from datetime import datetime

try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass

DATABASE_URL = os.getenv("DATABASE_URL", "")
USE_POSTGRES = DATABASE_URL.startswith("postgres")

def _sqlite_conn():
    import sqlite3
    db_path = os.path.join(os.path.dirname(__file__), "..", "machine_log.db")
    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    return conn

def _pg_conn():
    import psycopg2, psycopg2.extras
    url = DATABASE_URL
    if url.startswith("postgres://"):
        url = "postgresql://" + url[len("postgres://"):]
    conn = psycopg2.connect(url, cursor_factory=psycopg2.extras.RealDictCursor)
    return conn

def _get_conn():
    return _pg_conn() if USE_POSTGRES else _sqlite_conn()

PH = "%s" if USE_POSTGRES else "?"

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
    defect_rate_pct         REAL,
    UNIQUE(project, date)
)
"""
_CREATE_SQLITE = _CREATE \
    .replace("SERIAL PRIMARY KEY", "INTEGER PRIMARY KEY AUTOINCREMENT") \
    .replace("TIMESTAMP DEFAULT NOW()", "TEXT DEFAULT (datetime('now'))")

# Thêm UNIQUE constraint nếu table cũ chưa có (migration nhẹ)
_ADD_UNIQUE_SQLITE = """
CREATE UNIQUE INDEX IF NOT EXISTS idx_project_date ON logs(project, date)
"""
_ADD_UNIQUE_PG = """
DO $$ BEGIN
    IF NOT EXISTS (
        SELECT 1 FROM pg_indexes
        WHERE tablename='logs' AND indexname='idx_project_date'
    ) THEN
        CREATE UNIQUE INDEX idx_project_date ON logs(project, date);
    END IF;
END $$;
"""

def init_db():
    conn = _get_conn()
    try:
        cur = conn.cursor()
        sql = _CREATE if USE_POSTGRES else _CREATE_SQLITE
        cur.execute(sql)
        # Thêm unique index cho table cũ chưa có constraint
        try:
            cur.execute(_ADD_UNIQUE_PG if USE_POSTGRES else _ADD_UNIQUE_SQLITE)
        except Exception:
            pass
        conn.commit()
    finally:
        conn.close()


# ── UPSERT: insert hoặc update nếu đã có cùng (project, date) ────
_FIELDS = [
    "mo_color", "machine_line", "sam", "standard_working_time", "hc_operators",
    "actual_output_day_hc", "actual_output_day_mc", "auto_output_h_hc",
    "actual_output_h_mc", "manual_output_day_hc", "manual_output_h_hc",
    "machine_working_time", "changeover_time", "breakdown_time", "idle_time",
    "defects_json", "target_output_day", "auto_target_h_hc", "eff_pct",
    "mc_utilization_pct", "total_time_min", "total_defect", "defect_rate_pct",
]

def save_record(entry, computed: dict) -> int:
    """
    Upsert an toàn: SELECT trước → INSERT nếu chưa có, UPDATE nếu đã có.
    Không phụ thuộc vào UNIQUE constraint trong DB.
    """
    defects_json = json.dumps([d.dict() for d in entry.defects])
    values = [
        entry.mo_color, entry.machine_line, entry.sam,
        entry.standard_working_time, entry.hc_operators,
        entry.actual_output_day_hc, entry.actual_output_day_mc,
        entry.auto_output_h_hc, entry.actual_output_h_mc,
        entry.manual_output_day_hc, entry.manual_output_h_hc,
        entry.machine_working_time, entry.changeover_time,
        entry.breakdown_time, entry.idle_time,
        defects_json,
        computed["target_output_day"], computed["auto_target_h_hc"],
        computed["eff_pct"], computed["mc_utilization_pct"],
        computed["total_time_min"], computed["total_defect"],
        computed["defect_rate_pct"],
    ]

    conn = _get_conn()
    try:
        cur = conn.cursor()

        # ── Bước 1: Kiểm tra record đã tồn tại chưa ──────────────
        cur.execute(
            f"SELECT id FROM logs WHERE project={PH} AND date={PH} LIMIT 1",
            (entry.project, entry.date)
        )
        existing = cur.fetchone()

        if existing is None:
            # ── Bước 2a: INSERT mới ────────────────────────────────
            all_fields = ["project", "date"] + _FIELDS
            sql_insert = f"""
                INSERT INTO logs ({", ".join(all_fields)})
                VALUES ({", ".join([PH] * len(all_fields))})
            """
            cur.execute(sql_insert, [entry.project, entry.date] + values)

            if USE_POSTGRES:
                cur.execute("SELECT lastval()")
                row_id = cur.fetchone()[0]
            else:
                row_id = cur.lastrowid

        else:
            # ── Bước 2b: UPDATE record hiện có ────────────────────
            existing_id = existing["id"] if isinstance(existing, dict) else existing[0]
            set_clause  = ", ".join(f"{f} = {PH}" for f in _FIELDS)
            sql_update  = f"""
                UPDATE logs
                SET {set_clause}
                WHERE id = {PH}
            """
            cur.execute(sql_update, values + [existing_id])
            row_id = existing_id

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
                f"SELECT * FROM logs WHERE project={PH} ORDER BY date ASC LIMIT {PH}",
                (project, limit)
            )
        else:
            cur.execute(
                f"SELECT * FROM logs ORDER BY date ASC LIMIT {PH}", (limit,))

        rows = cur.fetchall()
        result = []
        for row in rows:
            d = dict(row)
            d["defects"] = json.loads(d.get("defects_json") or "[]")
            del d["defects_json"]
            if isinstance(d.get("created_at"), datetime):
                d["created_at"] = d["created_at"].isoformat()
            result.append(d)
        return result
    finally:
        conn.close()
