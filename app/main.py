from fastapi import FastAPI, HTTPException
from fastapi.responses import FileResponse, JSONResponse
from fastapi.middleware.cors import CORSMiddleware
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel
from typing import Optional, List
import os, uuid
from datetime import datetime
from .excel import build_excel
from .database import init_db, save_record, get_records

app = FastAPI(title="Regent Machine Log API", version="1.0.0")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

BASE_DIR    = os.path.dirname(os.path.dirname(__file__))
STATIC_DIR  = os.path.join(BASE_DIR, "static")
EXPORTS_DIR = os.path.join(BASE_DIR, "exports")
os.makedirs(EXPORTS_DIR, exist_ok=True)

@app.on_event("startup")
def startup():
    init_db()

# ── Pydantic models ──────────────────────────────────────────────
class DefectItem(BaseModel):
    name: str
    qty: float = 0

class LogEntry(BaseModel):
    project: str
    mo_color: str
    date: str
    machine_line: Optional[str] = ""
    sam: float
    standard_working_time: float
    hc_operators: float
    actual_output_day_hc: float = 0
    actual_output_day_mc: float = 0
    auto_output_h_hc: float = 0
    actual_output_h_mc: float = 0
    manual_output_day_hc: float = 0
    manual_output_h_hc: float = 0
    machine_working_time: float = 0
    changeover_time: float = 0
    breakdown_time: float = 0
    idle_time: float = 0
    defects: List[DefectItem] = []

# ── Computed ─────────────────────────────────────────────────────
def compute(entry: LogEntry) -> dict:
    sam, swt, hc = entry.sam, entry.standard_working_time, entry.hc_operators
    target_day     = round((swt * 60) / sam, 0) if sam > 0 else 0
    auto_target_h  = round(60 / sam / hc, 2)    if sam > 0 and hc > 0 else 0
    eff            = round(entry.actual_output_day_hc / target_day * 100, 1) \
                     if target_day > 0 and entry.actual_output_day_hc > 0 else 0
    total_time     = (entry.machine_working_time + entry.changeover_time +
                      entry.breakdown_time + entry.idle_time)
    mc_util        = round(entry.machine_working_time / total_time * 100, 1) \
                     if total_time > 0 else 0
    total_defect   = sum(d.qty for d in entry.defects)
    defect_rate    = round(total_defect / entry.actual_output_day_mc * 100, 2) \
                     if entry.actual_output_day_mc > 0 else 0
    return dict(target_output_day=int(target_day), auto_target_h_hc=auto_target_h,
                eff_pct=eff, mc_utilization_pct=mc_util, total_time_min=total_time,
                total_defect=total_defect, defect_rate_pct=defect_rate)

# ── API routes ───────────────────────────────────────────────────
@app.post("/api/log")
def create_log(entry: LogEntry):
    computed = compute(entry)
    record_id = save_record(entry, computed)
    return {"id": record_id, "computed": computed, "status": "saved"}

@app.post("/api/log/export")
def export_log(entry: LogEntry):
    computed = compute(entry)
    filename = f"MachineLog_{entry.project.replace(' ','_')}_{entry.date}_{uuid.uuid4().hex[:6]}.xlsx"
    filepath = os.path.join(EXPORTS_DIR, filename)
    build_excel(entry, computed, filepath)
    save_record(entry, computed)
    return FileResponse(path=filepath, filename=filename,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

@app.get("/api/logs")
def list_logs(project: Optional[str] = None, limit: int = 200):
    return get_records(project=project, limit=limit)

@app.get("/api/logs/{record_id}")
def get_log(record_id: int):
    rows = get_records(record_id=record_id)
    if not rows:
        raise HTTPException(status_code=404, detail="Record not found")
    return rows[0]

@app.get("/health")
def health():
    return {"status": "ok", "timestamp": datetime.utcnow().isoformat()}

# ── Serve HTML pages ──────────────────────────────────────────────
@app.get("/")
def serve_index():
    return FileResponse(os.path.join(STATIC_DIR, "index.html"))

@app.get("/dashboard")
def serve_dashboard():
    return FileResponse(os.path.join(STATIC_DIR, "dashboard.html"))

# Mount static (CSS/JS assets nếu cần sau này)
app.mount("/static", StaticFiles(directory=STATIC_DIR), name="static")
