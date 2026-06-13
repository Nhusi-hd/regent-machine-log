from fastapi import FastAPI, HTTPException, Query
from fastapi.responses import FileResponse, JSONResponse
from fastapi.middleware.cors import CORSMiddleware
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel
from typing import Optional, List
import os, uuid
from datetime import datetime
from .excel import build_excel
from .excel_monthly import build_monthly_excel
from .database import init_db, save_record, get_records
from .gsheets import push_to_sheet, write_monthly_sheet, GOOGLE_SHEETS_ENABLED

app = FastAPI(title="Regent Machine Log API", version="2.0.0")

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
    mo_color: Optional[str] = ""
    date: str
    machine_line: Optional[str] = ""
    sam: float = 0
    standard_working_time: float = 0
    hc_operators: float = 0
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

class MonthlyExportRequest(BaseModel):
    year: int
    month: int                       # 1–12
    projects: Optional[List[str]] = None   # None = tất cả dự án

class SheetPushRequest(BaseModel):
    entry: LogEntry
    spreadsheet_id: str              # ID của Google Sheet


# ── Computed ─────────────────────────────────────────────────────
def compute(entry: LogEntry) -> dict:
    sam, swt, hc = entry.sam, entry.standard_working_time, entry.hc_operators
    target_day    = round((swt * 60) / sam, 0)   if sam > 0 else 0
    auto_target_h = round(60 / sam / hc, 2)       if sam > 0 and hc > 0 else 0
    eff           = round(entry.actual_output_day_hc / target_day * 100, 1) \
                    if target_day > 0 and entry.actual_output_day_hc > 0 else 0
    total_time    = (entry.machine_working_time + entry.changeover_time +
                     entry.breakdown_time + entry.idle_time)
    mc_util       = round(entry.machine_working_time / total_time * 100, 1) \
                    if total_time > 0 else 0
    total_defect  = sum(d.qty for d in entry.defects)
    defect_rate   = round(total_defect / entry.actual_output_day_mc * 100, 2) \
                    if entry.actual_output_day_mc > 0 else 0
    return dict(
        target_output_day=int(target_day),
        auto_target_h_hc=auto_target_h,
        eff_pct=eff,
        mc_utilization_pct=mc_util,
        total_time_min=total_time,
        total_defect=total_defect,
        defect_rate_pct=defect_rate,
    )


# ════════════════════════════════════════════════════════════════
# API ROUTES
# ════════════════════════════════════════════════════════════════

# ── 1. Lưu record ────────────────────────────────────────────────
@app.post("/api/log")
def create_log(entry: LogEntry):
    computed  = compute(entry)
    record_id = save_record(entry, computed)
    return {"id": record_id, "computed": computed, "status": "saved"}


# ── 2. Lưu + xuất Excel từng ngày ────────────────────────────────
@app.post("/api/log/export")
def export_log(entry: LogEntry):
    computed = compute(entry)
    filename = (f"MachineLog_{entry.project.replace(' ','_')}"
                f"_{entry.date}_{uuid.uuid4().hex[:6]}.xlsx")
    filepath = os.path.join(EXPORTS_DIR, filename)
    build_excel(entry, computed, filepath)
    save_record(entry, computed)
    return FileResponse(
        path=filepath, filename=filename,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


# ── 3. Xuất Excel tháng (multi-project, multi-date-column) ───────
@app.post("/api/export/monthly")
def export_monthly(req: MonthlyExportRequest):
    if not (1 <= req.month <= 12):
        raise HTTPException(400, "month phải từ 1–12")

    # Lấy records trong tháng từ DB
    import calendar
    _, last_day = calendar.monthrange(req.year, req.month)
    date_from   = f"{req.year}-{req.month:02d}-01"
    date_to     = f"{req.year}-{req.month:02d}-{last_day:02d}"

    all_records = get_records(limit=5000)
    records     = [
        r for r in all_records
        if date_from <= r["date"] <= date_to
        and (req.projects is None or r["project"] in req.projects)
    ]

    if not records:
        raise HTTPException(404, f"Không có dữ liệu tháng {req.month}/{req.year}")

    import calendar as cal
    month_name = cal.month_abbr[req.month]
    filename   = f"MachineLog_Monthly_{month_name}{req.year}_{uuid.uuid4().hex[:4]}.xlsx"
    filepath   = os.path.join(EXPORTS_DIR, filename)
    build_monthly_excel(records, req.year, req.month, filepath)

    return FileResponse(
        path=filepath, filename=filename,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


# ── 4. Đẩy 1 record lên Google Sheet ────────────────────────────
@app.post("/api/log/push-sheet")
def push_sheet(req: SheetPushRequest):
    if not GOOGLE_SHEETS_ENABLED:
        raise HTTPException(503,
            "Google Sheets chưa được cấu hình. "
            "Set biến môi trường GOOGLE_SERVICE_ACCOUNT_JSON và SPREADSHEET_ID.")
    computed = compute(req.entry)
    save_record(req.entry, computed)
    try:
        result = push_to_sheet(req.entry, computed, req.spreadsheet_id)
        return {"status": "pushed", "computed": computed, **result}
    except Exception as e:
        raise HTTPException(500, f"Lỗi Google Sheets: {str(e)}")


# ── 5. Lưu + xuất Excel + đẩy Sheet cùng lúc ───────────────────
@app.post("/api/log/save-all")
def save_all(req: SheetPushRequest):
    """Lưu DB + xuất .xlsx + đẩy Google Sheet trong 1 request."""
    entry    = req.entry
    computed = compute(entry)

    # Lưu DB
    record_id = save_record(entry, computed)

    # Xuất xlsx
    filename = (f"MachineLog_{entry.project.replace(' ','_')}"
                f"_{entry.date}_{uuid.uuid4().hex[:6]}.xlsx")
    filepath = os.path.join(EXPORTS_DIR, filename)
    build_excel(entry, computed, filepath)

    # Đẩy Google Sheet (nếu đã cấu hình)
    sheet_result = None
    if GOOGLE_SHEETS_ENABLED and req.spreadsheet_id:
        try:
            sheet_result = push_to_sheet(entry, computed, req.spreadsheet_id)
        except Exception as e:
            sheet_result = {"error": str(e)}

    return FileResponse(
        path=filepath, filename=filename,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={
            "X-Record-ID":    str(record_id),
            "X-Sheet-Status": "pushed" if sheet_result and "url" in sheet_result else "skipped",
            "X-Sheet-URL":    sheet_result.get("url", "") if sheet_result else "",
        },
    )


# ── 6. Kiểm tra trạng thái Google Sheet ──────────────────────────
@app.get("/api/sheets/status")
def sheets_status():
    return {
        "enabled": GOOGLE_SHEETS_ENABLED,
        "spreadsheet_id": os.getenv("SPREADSHEET_ID", ""),
        "message": (
            "Google Sheets đã kết nối" if GOOGLE_SHEETS_ENABLED
            else "Chưa cấu hình — set GOOGLE_SERVICE_ACCOUNT_JSON và SPREADSHEET_ID"
        ),
    }


# ── 7. Lấy danh sách records ─────────────────────────────────────
@app.get("/api/logs")
def list_logs(project: Optional[str] = None, limit: int = 200):
    return get_records(project=project, limit=limit)

@app.get("/api/logs/{record_id}")
def get_log(record_id: int):
    rows = get_records(record_id=record_id)
    if not rows:
        raise HTTPException(404, "Record not found")
    return rows[0]

@app.get("/health")
def health():
    return {
        "status": "ok",
        "timestamp": datetime.utcnow().isoformat(),
        "google_sheets": GOOGLE_SHEETS_ENABLED,
    }



# ── 10. Ghi toàn bộ tháng lên Google Sheet theo đúng template ────
@app.post("/api/sheets/write-monthly")
def write_monthly_to_sheet(
    year:           int,
    month:          int,
    spreadsheet_id: Optional[str] = Query(default=None),
    project:        Optional[str] = Query(default=None),
):
    """
    Lấy data từ DB theo tháng → ghi lên Google Sheet
    đúng format template (rows=chỉ số, cols=ngày, màu vàng/xanh).
    """
    from .gsheets import write_monthly_sheet, GOOGLE_SHEETS_ENABLED, SPREADSHEET_ID
    if not GOOGLE_SHEETS_ENABLED:
        raise HTTPException(503, "Google Sheets chưa được cấu hình")
    import calendar as cal
    _, last_day = cal.monthrange(year, month)
    date_from   = f"{year}-{month:02d}-01"
    date_to     = f"{year}-{month:02d}-{last_day:02d}"
    all_records = get_records(limit=5000)
    records     = [
        r for r in all_records
        if date_from <= r["date"] <= date_to
        and (project is None or r["project"] == project)
    ]
    if not records:
        raise HTTPException(404, f"Không có dữ liệu tháng {month}/{year}")
    sid    = spreadsheet_id or SPREADSHEET_ID
    result = write_monthly_sheet(records, year, month, sid)
    return result

# ── Serve HTML ────────────────────────────────────────────────────
@app.get("/")
def serve_index():
    return FileResponse(os.path.join(STATIC_DIR, "index.html"))

@app.get("/dashboard")
def serve_dashboard():
    return FileResponse(os.path.join(STATIC_DIR, "dashboard.html"))

app.mount("/static", StaticFiles(directory=STATIC_DIR), name="static")


# ── 8. Sync toàn bộ DB lên Google Sheet ──────────────────────────
@app.post("/api/sheets/sync-all")
def sync_all_to_sheet(
    spreadsheet_id: Optional[str] = Query(default=None),
    project: Optional[str] = Query(default=None),
):
    """Đẩy toàn bộ records trong DB lên Google Sheet (batch)."""
    from .gsheets import push_batch_to_sheet, GOOGLE_SHEETS_ENABLED, SPREADSHEET_ID
    if not GOOGLE_SHEETS_ENABLED:
        raise HTTPException(503, "Google Sheets chưa được cấu hình")
    sid     = spreadsheet_id or SPREADSHEET_ID
    records = get_records(project=project, limit=5000)
    if not records:
        raise HTTPException(404, "Không có records nào trong DB")
    from .gsheets import push_batch_to_sheet
    result  = push_batch_to_sheet(records, sid)
    return result


# ── 9. Lấy thông tin Google Sheet ────────────────────────────────
@app.get("/api/sheets/info")
def sheet_info(spreadsheet_id: Optional[str] = Query(default=None)):
    """Lấy thông tin Sheet: tên file, tabs, số records mỗi tab."""
    from .gsheets import get_sheet_info, GOOGLE_SHEETS_ENABLED, SPREADSHEET_ID
    if not GOOGLE_SHEETS_ENABLED:
        raise HTTPException(503, "Google Sheets chưa được cấu hình")
    sid = spreadsheet_id or SPREADSHEET_ID
    return get_sheet_info(sid)
