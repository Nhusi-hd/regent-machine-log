"""
gsheets.py — Kết nối Google Sheets API v4
─────────────────────────────────────────
Cách hoạt động:
  • Mỗi dự án → 1 tab riêng trong Google Sheet
  • Mỗi lần nhấn "Lưu & Gửi Sheet" → append 1 dòng mới vào tab đó
  • Header row tự tạo (màu navy) nếu tab chưa tồn tại
  • Hàng EFF/Defect rate tự tô màu xanh (công thức)

Biến môi trường cần set trên Render:
  GOOGLE_SERVICE_ACCOUNT_JSON  → nội dung file JSON service account
  SPREADSHEET_ID               → ID của Google Sheet (lấy từ URL)
"""

import os, json

# ── Kiểm tra đã cấu hình chưa ────────────────────────────────────
GOOGLE_SERVICE_ACCOUNT_JSON = os.getenv("GOOGLE_SERVICE_ACCOUNT_JSON", "")
SPREADSHEET_ID              = os.getenv("SPREADSHEET_ID", "")
GOOGLE_SHEETS_ENABLED       = bool(GOOGLE_SERVICE_ACCOUNT_JSON and SPREADSHEET_ID)

# ── Màu sắc (RGB dạng dict cho Sheets API) ───────────────────────
def _rgb(hex_color: str) -> dict:
    """Chuyển hex string → dict {red, green, blue} (0.0–1.0)."""
    h = hex_color.lstrip("#")
    return {
        "red":   int(h[0:2], 16) / 255,
        "green": int(h[2:4], 16) / 255,
        "blue":  int(h[4:6], 16) / 255,
    }

NAVY_RGB   = _rgb("1A2744")
YELLOW_RGB = _rgb("FFD700")
BLUE_RGB   = _rgb("00B0F0")
WHITE_RGB  = _rgb("FFFFFF")
GREEN_RGB  = _rgb("00B050")
GRAY_RGB   = _rgb("F2F2F2")

# ── Danh sách cột ────────────────────────────────────────────────
# (header_label, db_field_or_computed, is_formula)
COLUMNS = [
    ("Ngày",                   "date",                    False),
    ("Dự Án",                  "project",                 False),
    ("MO/Color",               "mo_color",                False),
    ("Máy/Line",               "machine_line",            False),
    ("SAM",                    "sam",                     False),
    ("SWT (h)",                "standard_working_time",   False),
    ("HC vận hành",            "hc_operators",            False),
    ("Target/day (pcs)",       "target_output_day",       True),
    ("Auto Target/h/HC",       "auto_target_h_hc",        True),
    ("Actual HC/day (pcs)",    "actual_output_day_hc",    False),
    ("Actual MC/day (pcs)",    "actual_output_day_mc",    False),
    ("Auto/h/HC (pcs)",        "auto_output_h_hc",        False),
    ("Actual/h/MC (pcs)",      "actual_output_h_mc",      False),
    ("Manual/day/HC (pcs)",    "manual_output_day_hc",    False),
    ("Manual/h/HC (pcs)",      "manual_output_h_hc",      False),
    ("Working time (min)",     "machine_working_time",    False),
    ("Change-over (min)",      "changeover_time",         False),
    ("Breakdown (min)",        "breakdown_time",          False),
    ("Idle (min)",             "idle_time",               False),
    ("Total time (min)",       "total_time_min",          True),
    ("EFF (%)",                "eff_pct",                 True),
    ("MC Utilization (%)",     "mc_utilization_pct",      True),
    ("Total Defect (pcs)",     "total_defect",            True),
    ("Defect rate (%)",        "defect_rate_pct",         True),
    ("Defect detail",          "_defects_str",            False),
]

HEADERS = [col[0] for col in COLUMNS]


# ── Khởi tạo Google Sheets service ───────────────────────────────
def _get_service():
    try:
        from google.oauth2.service_account import Credentials
        from googleapiclient.discovery import build
    except ImportError:
        raise RuntimeError(
            "Thiếu thư viện Google. Chạy: "
            "pip install google-auth google-api-python-client"
        )
    if not GOOGLE_SERVICE_ACCOUNT_JSON:
        raise RuntimeError("Chưa set GOOGLE_SERVICE_ACCOUNT_JSON")

    info  = json.loads(GOOGLE_SERVICE_ACCOUNT_JSON)
    creds = Credentials.from_service_account_info(
        info,
        scopes=["https://www.googleapis.com/auth/spreadsheets"],
    )
    return build("sheets", "v4", credentials=creds, cache_discovery=False)


# ── Lấy sheet_id từ tab name ─────────────────────────────────────
def _get_sheet_id(service, spreadsheet_id: str, tab_name: str) -> int:
    meta = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
    for s in meta["sheets"]:
        if s["properties"]["title"] == tab_name:
            return s["properties"]["sheetId"]
    return 0


# ── Format header row (màu navy + chữ trắng) ─────────────────────
def _format_header(service, spreadsheet_id: str, sheet_id: int, num_cols: int):
    """Tô màu header row 1 và các ô công thức."""
    col_formats = []

    for i, (_, _, is_formula) in enumerate(COLUMNS):
        bg = BLUE_RGB if is_formula else NAVY_RGB
        col_formats.append({
            "repeatCell": {
                "range": {
                    "sheetId":          sheet_id,
                    "startRowIndex":    0,
                    "endRowIndex":      1,
                    "startColumnIndex": i,
                    "endColumnIndex":   i + 1,
                },
                "cell": {
                    "userEnteredFormat": {
                        "backgroundColor": bg,
                        "textFormat": {
                            "foregroundColor": WHITE_RGB,
                            "bold": True,
                            "fontSize": 10,
                            "fontFamily": "Calibri",
                        },
                        "horizontalAlignment": "CENTER",
                        "verticalAlignment":   "MIDDLE",
                    }
                },
                "fields": "userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)",
            }
        })

    # Freeze row 1
    col_formats.append({
        "updateSheetProperties": {
            "properties": {
                "sheetId": sheet_id,
                "gridProperties": {"frozenRowCount": 1},
            },
            "fields": "gridProperties.frozenRowCount",
        }
    })

    # Auto-resize tất cả cột
    col_formats.append({
        "autoResizeDimensions": {
            "dimensions": {
                "sheetId":    sheet_id,
                "dimension":  "COLUMNS",
                "startIndex": 0,
                "endIndex":   num_cols,
            }
        }
    })

    service.spreadsheets().batchUpdate(
        spreadsheetId=spreadsheet_id,
        body={"requests": col_formats},
    ).execute()


# ── Format data row (màu vàng/xanh theo loại ô) ──────────────────
def _format_data_row(service, spreadsheet_id: str, sheet_id: int, row_index: int):
    """Tô màu dòng data vừa append: vàng=nhập tay, xanh=công thức."""
    requests = []
    for i, (_, _, is_formula) in enumerate(COLUMNS):
        bg = BLUE_RGB if is_formula else YELLOW_RGB
        requests.append({
            "repeatCell": {
                "range": {
                    "sheetId":          sheet_id,
                    "startRowIndex":    row_index,
                    "endRowIndex":      row_index + 1,
                    "startColumnIndex": i,
                    "endColumnIndex":   i + 1,
                },
                "cell": {
                    "userEnteredFormat": {
                        "backgroundColor": bg,
                        "textFormat": {
                            "bold":       is_formula,
                            "fontSize":   10,
                            "fontFamily": "Calibri",
                            "foregroundColor": _rgb("0070C0") if is_formula else _rgb("000000"),
                        },
                        "horizontalAlignment": "CENTER" if is_formula else "LEFT",
                    }
                },
                "fields": "userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)",
            }
        })

    service.spreadsheets().batchUpdate(
        spreadsheetId=spreadsheet_id,
        body={"requests": requests},
    ).execute()


# ── Tạo tab mới nếu chưa có ──────────────────────────────────────
def _ensure_tab(service, spreadsheet_id: str, tab_name: str) -> tuple:
    """
    Đảm bảo tab tồn tại với header đã được format.
    Trả về (tab_name, sheet_id, is_new).
    """
    meta     = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
    existing = {s["properties"]["title"]: s["properties"]["sheetId"]
                for s in meta["sheets"]}

    if tab_name in existing:
        return tab_name, existing[tab_name], False

    # Tạo tab mới
    result = service.spreadsheets().batchUpdate(
        spreadsheetId=spreadsheet_id,
        body={"requests": [{"addSheet": {"properties": {"title": tab_name}}}]},
    ).execute()
    sheet_id = result["replies"][0]["addSheet"]["properties"]["sheetId"]

    # Ghi header row
    service.spreadsheets().values().update(
        spreadsheetId=spreadsheet_id,
        range=f"'{tab_name}'!A1",
        valueInputOption="RAW",
        body={"values": [HEADERS]},
    ).execute()

    # Format header
    _format_header(service, spreadsheet_id, sheet_id, len(HEADERS))

    return tab_name, sheet_id, True


# ── Build row data từ entry + computed ───────────────────────────
def _build_row(entry, computed: dict) -> list:
    defects_str = "; ".join(
        f"{d.name}: {int(d.qty)}" for d in (entry.defects or []) if d.name
    )
    # Gộp tất cả giá trị vào 1 dict để lookup theo field name
    merged = {
        "date":                  entry.date,
        "project":               entry.project,
        "mo_color":              entry.mo_color or "",
        "machine_line":          entry.machine_line or "",
        "sam":                   entry.sam,
        "standard_working_time": entry.standard_working_time,
        "hc_operators":          entry.hc_operators,
        "actual_output_day_hc":  entry.actual_output_day_hc,
        "actual_output_day_mc":  entry.actual_output_day_mc,
        "auto_output_h_hc":      entry.auto_output_h_hc,
        "actual_output_h_mc":    entry.actual_output_h_mc,
        "manual_output_day_hc":  entry.manual_output_day_hc,
        "manual_output_h_hc":    entry.manual_output_h_hc,
        "machine_working_time":  entry.machine_working_time,
        "changeover_time":       entry.changeover_time,
        "breakdown_time":        entry.breakdown_time,
        "idle_time":             entry.idle_time,
        "_defects_str":          defects_str,
        **computed,
    }
    return [merged.get(field, "") for _, field, _ in COLUMNS]


# ── Lấy số dòng hiện tại trong tab ──────────────────────────────
def _get_last_row(service, spreadsheet_id: str, tab_name: str) -> int:
    result = service.spreadsheets().values().get(
        spreadsheetId=spreadsheet_id,
        range=f"'{tab_name}'!A:A",
    ).execute()
    return len(result.get("values", []))


# ════════════════════════════════════════════════════════════════
# PUBLIC FUNCTIONS
# ════════════════════════════════════════════════════════════════

def push_to_sheet(entry, computed: dict, spreadsheet_id: str) -> dict:
    """
    Đẩy 1 record lên Google Sheet.
    - Tự tạo tab nếu chưa có
    - Append dòng mới
    - Tô màu vàng/xanh theo loại ô
    Trả về {"url", "tab", "row", "sheet_url"}
    """
    service                    = _get_service()
    tab_name                   = entry.project[:31]
    tab, sheet_id, _           = _ensure_tab(service, spreadsheet_id, tab_name)
    row_data                   = _build_row(entry, computed)

    # Append data
    service.spreadsheets().values().append(
        spreadsheetId=spreadsheet_id,
        range=f"'{tab}'!A1",
        valueInputOption="USER_ENTERED",
        insertDataOption="INSERT_ROWS",
        body={"values": [row_data]},
    ).execute()

    # Lấy index dòng vừa append để format màu
    last_row = _get_last_row(service, spreadsheet_id, tab)
    try:
        _format_data_row(service, spreadsheet_id, sheet_id, last_row - 1)
    except Exception:
        pass   # format lỗi không block việc lưu

    sheet_url = (
        f"https://docs.google.com/spreadsheets/d/{spreadsheet_id}"
        f"/edit#gid={sheet_id}"
    )
    return {
        "url":       sheet_url,
        "tab":       tab,
        "row":       last_row,
        "sheet_url": sheet_url,
    }


def push_batch_to_sheet(records: list, spreadsheet_id: str) -> dict:
    """
    Đẩy nhiều records cùng lúc (dùng cho sync hàng loạt).
    records: list of dict từ DB (đã có computed fields).
    """
    service  = _get_service()
    pushed   = 0
    errors   = []

    # Nhóm theo project
    proj_map: dict = {}
    for r in records:
        proj_map.setdefault(r["project"], []).append(r)

    for proj_name, proj_records in proj_map.items():
        tab_name          = proj_name[:31]
        tab, sheet_id, _ = _ensure_tab(service, spreadsheet_id, tab_name)

        rows_to_append = []
        for r in proj_records:
            # Tạo fake entry object từ dict
            class _FakeEntry:
                pass
            e = _FakeEntry()
            for k, v in r.items():
                setattr(e, k, v)
            e.defects = [
                type("D", (), {"name": d.get("name",""), "qty": d.get("qty",0)})()
                for d in (r.get("defects") or [])
            ]
            computed = {k: r.get(k) for k in [
                "target_output_day","auto_target_h_hc","eff_pct",
                "mc_utilization_pct","total_time_min","total_defect","defect_rate_pct"
            ]}
            rows_to_append.append(_build_row(e, computed))

        try:
            service.spreadsheets().values().append(
                spreadsheetId=spreadsheet_id,
                range=f"'{tab}'!A1",
                valueInputOption="USER_ENTERED",
                insertDataOption="INSERT_ROWS",
                body={"values": rows_to_append},
            ).execute()
            pushed += len(rows_to_append)
        except Exception as ex:
            errors.append(f"{proj_name}: {str(ex)}")

    return {
        "pushed": pushed,
        "errors": errors,
        "spreadsheet_url": f"https://docs.google.com/spreadsheets/d/{spreadsheet_id}/edit",
    }


def get_sheet_info(spreadsheet_id: str) -> dict:
    """Lấy thông tin Sheet: tên, danh sách tabs, số dòng mỗi tab."""
    service = _get_service()
    meta    = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
    tabs    = []
    for s in meta["sheets"]:
        title = s["properties"]["title"]
        # Đếm dòng
        try:
            res = service.spreadsheets().values().get(
                spreadsheetId=spreadsheet_id,
                range=f"'{title}'!A:A",
            ).execute()
            rows = max(0, len(res.get("values", [])) - 1)  # trừ header
        except Exception:
            rows = 0
        tabs.append({"name": title, "records": rows})

    return {
        "title":           meta["properties"]["title"],
        "spreadsheet_id":  spreadsheet_id,
        "url":             f"https://docs.google.com/spreadsheets/d/{spreadsheet_id}/edit",
        "tabs":            tabs,
        "total_records":   sum(t["records"] for t in tabs),
    }
