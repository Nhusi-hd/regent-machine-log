"""
gsheets.py — Ghi Google Sheet đúng format template gốc:

Layout:
  Cột A  = Section label  ("Overall" merge rows 3→21, "Defect (pcs)" merge rows defect)
  Cột B  = Row label      (tên chỉ số)
  Cột C+ = Ngày           (05-Jun, 06-Jun, ...)

Màu:
  Xanh lá  (#00B050) = không dùng (đã bỏ hàng Tên Dự Án)
  Xanh blue (#00B0F0) = công thức (Target, EFF, Defect rate, MC util, Total defect)
  Vàng     (#FFD700) = nhập tay
  Xanh cyan (#00B0F0 đậm) = Total defect row
  Navy     (#1A2744) = header Date row
"""

import os, json, calendar
from datetime import date

# ── Env vars ──────────────────────────────────────────────────────
GOOGLE_SERVICE_ACCOUNT_JSON = os.getenv("GOOGLE_SERVICE_ACCOUNT_JSON", "")
SPREADSHEET_ID              = os.getenv("SPREADSHEET_ID", "")
GOOGLE_SHEETS_ENABLED       = bool(GOOGLE_SERVICE_ACCOUNT_JSON and SPREADSHEET_ID)

# ── Màu → RGB dict cho Sheets API ────────────────────────────────
def _c(h):
    h = h.lstrip("#")
    return {"red": int(h[0:2],16)/255, "green": int(h[2:4],16)/255, "blue": int(h[4:6],16)/255}

C_NAVY      = _c("1A2744")
C_WHITE     = _c("FFFFFF")
C_YELLOW    = _c("FFD700")
C_BLUE      = _c("00B0F0")
C_GREEN     = _c("00B050")
C_CYAN      = _c("00CCEE")
C_GRAY      = _c("F2F2F2")
C_LIGHT     = _c("DDEEFF")
C_OVERALL   = _c("E8F0FE")   # màu section Overall (xanh lavender nhạt)
C_DEFECT_S  = _c("DCE6F1")   # màu section Defect (xanh nhạt)
C_BLUE_TXT  = _c("0070C0")
C_BLACK     = _c("000000")

# ── Row definitions (đúng thứ tự template) ───────────────────────
# (label_cot_B,  db_field,               mau_label, is_formula)
OVERALL_ROWS = [
    ("MO/Color",                 "mo_color",             "YELLOW", False),
    ("SAM (testing)",            "sam",                  "YELLOW", False),
    ("Target output / day (pcs)","target_output_day",    "BLUE",   True),
    ("Auto target output/h/HC",  "auto_target_h_hc",     "BLUE",   True),
    ("Actual ouput/day/HC(pcs)", "actual_output_day_hc", "YELLOW", False),
    ("Actual ouput/day/MC(pcs)", "actual_output_day_mc", "YELLOW", False),
    ("Auto output/h/HC (pcs)",   "auto_output_h_hc",     "YELLOW", False),
    ("Actual output/h/MC (pcs)", "actual_output_h_mc",   "YELLOW", False),
    ("Defect rate (%)",          "defect_rate_pct",      "BLUE",   True),
    ("EFF(%)",                   "eff_pct",              "YELLOW", False),  # vàng theo template
    ("Mc utilization (%)",       "mc_utilization_pct",   "BLUE",   True),
    ("Manual output/day/HC (pcs)","manual_output_day_hc","YELLOW", False),
    ("Manual output/h/HC (pcs)", "manual_output_h_hc",   "YELLOW", False),
    ("Machine working time (min)","machine_working_time", "YELLOW", False),
    ("Change-over time(min)",    "changeover_time",      "YELLOW", False),
    ("Breakdown time(min)",      "breakdown_time",       "YELLOW", False),
    ("Idle time(min)",           "idle_time",            "YELLOW", False),
]

COLOR_MAP = {
    "GREEN":  C_GREEN,
    "YELLOW": C_YELLOW,
    "BLUE":   C_BLUE,
}

def _get_service():
    try:
        from google.oauth2.service_account import Credentials
        from googleapiclient.discovery import build
    except ImportError:
        raise RuntimeError("pip install google-auth google-api-python-client")
    if not GOOGLE_SERVICE_ACCOUNT_JSON:
        raise RuntimeError("Chưa set GOOGLE_SERVICE_ACCOUNT_JSON")
    info  = json.loads(GOOGLE_SERVICE_ACCOUNT_JSON)
    creds = Credentials.from_service_account_info(
        info, scopes=["https://www.googleapis.com/auth/spreadsheets"])
    return build("sheets", "v4", credentials=creds, cache_discovery=False)


def _get_sheet_id(service, spreadsheet_id, tab_name):
    meta = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
    for s in meta["sheets"]:
        if s["properties"]["title"] == tab_name:
            return s["properties"]["sheetId"]
    return None


def _working_days(year, month):
    """Trả về list date là Mon–Sat trong tháng."""
    _, last = calendar.monthrange(year, month)
    return [date(year, month, d) for d in range(1, last+1)
            if date(year, month, d).weekday() < 6]


# ── Batch update helper ───────────────────────────────────────────
def _fmt(requests, service, spreadsheet_id):
    if requests:
        service.spreadsheets().batchUpdate(
            spreadsheetId=spreadsheet_id,
            body={"requests": requests}
        ).execute()


# ── Cell format builder ───────────────────────────────────────────
def _cell_fmt(sheet_id, r1, c1, r2, c2, bg=None, bold=False,
              fg=None, size=10, halign="CENTER", border=True,
              num_format=None):
    """
    num_format: dict với 'type' và 'pattern', ví dụ:
      {"type": "PERCENT", "pattern": "0.0%"}   → hiển thị 85.0%
      {"type": "NUMBER",  "pattern": "#,##0"}  → hiển thị 1,200
    """
    fmt = {
        "textFormat": {
            "bold": bold,
            "fontSize": size,
            "fontFamily": "Calibri",
        },
        "horizontalAlignment": halign,
        "verticalAlignment":   "MIDDLE",
    }
    if bg:         fmt["backgroundColor"] = bg
    if fg:         fmt["textFormat"]["foregroundColor"] = fg
    if num_format: fmt["numberFormat"] = num_format

    fields_list = [
        "backgroundColor" if bg else "",
        "textFormat",
        "horizontalAlignment",
        "verticalAlignment",
        "numberFormat" if num_format else "",
    ]
    fields_str = ",".join(f for f in fields_list if f)

    req = {
        "repeatCell": {
            "range": {"sheetId": sheet_id,
                      "startRowIndex": r1, "endRowIndex": r2,
                      "startColumnIndex": c1, "endColumnIndex": c2},
            "cell": {"userEnteredFormat": fmt},
            "fields": f"userEnteredFormat({fields_str})",
        }
    }
    return req


def _merge(sheet_id, r1, c1, r2, c2):
    return {"mergeCells": {
        "range": {"sheetId": sheet_id,
                  "startRowIndex": r1, "endRowIndex": r2,
                  "startColumnIndex": c1, "endColumnIndex": c2},
        "mergeType": "MERGE_ALL"
    }}


def _border_req(sheet_id, r1, c1, r2, c2):
    s = {"style": "SOLID", "color": _c("BFBFBF"), "width": 1}
    return {"updateBorders": {
        "range": {"sheetId": sheet_id,
                  "startRowIndex": r1, "endRowIndex": r2,
                  "startColumnIndex": c1, "endColumnIndex": c2},
        "top": s, "bottom": s, "left": s, "right": s,
        "innerHorizontal": s, "innerVertical": s,
    }}


def _col_width(sheet_id, col_idx, width_px):
    return {"updateDimensionProperties": {
        "range": {"sheetId": sheet_id, "dimension": "COLUMNS",
                  "startIndex": col_idx, "endIndex": col_idx+1},
        "properties": {"pixelSize": width_px},
        "fields": "pixelSize",
    }}

def _row_height(sheet_id, row_idx, height_px):
    return {"updateDimensionProperties": {
        "range": {"sheetId": sheet_id, "dimension": "ROWS",
                  "startIndex": row_idx, "endIndex": row_idx+1},
        "properties": {"pixelSize": height_px},
        "fields": "pixelSize",
    }}


# ════════════════════════════════════════════════════════════════
# MAIN: Tạo/cập nhật Sheet theo đúng template
# ════════════════════════════════════════════════════════════════
def write_monthly_sheet(records: list, year: int, month: int,
                        spreadsheet_id: str, tab_suffix: str = "") -> dict:
    """
    Ghi dữ liệu tháng lên Google Sheet đúng format template:
    - Mỗi dự án = 1 tab
    - Rows = chỉ số, Cols = ngày làm việc
    - Màu vàng/xanh đúng template

    records: list dict từ DB (đã có computed fields + defects list)
    """
    service      = _get_service()
    working_days = _working_days(year, month)
    month_abbr   = calendar.month_abbr[month]

    # Nhóm theo dự án
    proj_map: dict = {}
    for r in records:
        proj_map.setdefault(r["project"], []).append(r)

    result_tabs = []

    for proj_name, proj_records in proj_map.items():
        date_map = {r["date"]: r for r in proj_records}

        # ── Tạo/lấy tab ───────────────────────────────────────────
        tab_name = (proj_name[:28] + tab_suffix)[:31]
        meta     = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
        existing = {s["properties"]["title"]: s["properties"]["sheetId"]
                    for s in meta["sheets"]}

        if tab_name in existing:
            # Xóa nội dung cũ để ghi lại
            sheet_id = existing[tab_name]
            service.spreadsheets().values().clear(
                spreadsheetId=spreadsheet_id,
                range=f"'{tab_name}'",
                body={}
            ).execute()
            # Bỏ merge cũ
            _fmt([{"unmergeCells": {"range": {
                "sheetId": sheet_id,
                "startRowIndex": 0, "endRowIndex": 200,
                "startColumnIndex": 0, "endColumnIndex": 50,
            }}}], service, spreadsheet_id)
        else:
            res = service.spreadsheets().batchUpdate(
                spreadsheetId=spreadsheet_id,
                body={"requests": [{"addSheet": {
                    "properties": {"title": tab_name}
                }}]}
            ).execute()
            sheet_id = res["replies"][0]["addSheet"]["properties"]["sheetId"]

        # ── Chuẩn bị dữ liệu ──────────────────────────────────────
        # Thu thập tất cả tên lỗi
        defect_names = []
        for rec in proj_records:
            for d in (rec.get("defects") or []):
                if d.get("name") and d["name"] not in defect_names:
                    defect_names.append(d["name"])

        num_days    = len(working_days)
        num_cols    = 2 + num_days      # A, B, ngày...
        last_col    = num_cols          # 1-based
        OVERALL_LEN = len(OVERALL_ROWS)

        # Row indices (0-based)
        R_TITLE     = 0
        R_DATE_HDR  = 1
        R_OVERALL_S = 2                             # MO/Color (bắt đầu từ đây)
        R_OVERALL_E = R_OVERALL_S + OVERALL_LEN    # exclusive
        R_NAME_HDR  = R_OVERALL_E                  # "Name" row
        R_DEFECT_S  = R_NAME_HDR + 1
        R_DEFECT_E  = R_DEFECT_S + max(len(defect_names), 1)
        R_TOTAL     = R_DEFECT_E
        TOTAL_ROWS  = R_TOTAL + 1

        # ── Ghi giá trị (values) ──────────────────────────────────
        all_values = []

        # Row 0: Title
        title_row = [f"DAILY MACHINE PERFORMANCE LOG — {proj_name.upper()} — {month_abbr}-{year}"]
        title_row += [""] * (num_cols - 1)
        all_values.append(title_row)

        # Row 1: Date header
        date_hdr = ["", "Date"]
        for d in working_days:
            date_hdr.append(f"{d.day:02d}-{month_abbr}")
        all_values.append(date_hdr)

        # Rows Overall
        for label, field, color, is_formula in OVERALL_ROWS:
            row = ["", label]
            for d in working_days:
                date_str = d.strftime("%Y-%m-%d")
                rec      = date_map.get(date_str)
                val      = ""
                if rec:
                    v = rec.get(field)
                    if v is not None:
                        if field in ("eff_pct","defect_rate_pct","mc_utilization_pct"):
                            # Chia 100 → dạng thập phân 0.xx để Google Sheets
                            # hiển thị đúng khi format PERCENT (tránh nhân 100 lần)
                            val = round(float(v) / 100, 4)
                        elif field in ("target_output_day","actual_output_day_hc",
                                       "actual_output_day_mc","manual_output_day_hc",
                                       "machine_working_time","changeover_time",
                                       "breakdown_time","idle_time"):
                            val = int(v)
                        else:
                            val = v
                row.append(val)
            all_values.append(row)

        # Row "Name" (defect header)
        name_row = ["", "Name"] + [""] * num_days
        all_values.append(name_row)

        # Defect rows
        if defect_names:
            for dname in defect_names:
                row = ["", dname]
                for d in working_days:
                    date_str = d.strftime("%Y-%m-%d")
                    rec      = date_map.get(date_str)
                    qty      = ""
                    if rec:
                        for def_item in (rec.get("defects") or []):
                            if def_item.get("name") == dname:
                                qty = int(def_item.get("qty", 0)) or ""
                                break
                    row.append(qty)
                all_values.append(row)
        else:
            all_values.append(["", ""] + [""] * num_days)

        # Total defect row
        total_row = ["Total defect", ""]
        for d in working_days:
            date_str = d.strftime("%Y-%m-%d")
            rec      = date_map.get(date_str)
            total_row.append(int(rec.get("total_defect", 0)) if rec else "")
        all_values.append(total_row)

        # Ghi tất cả giá trị 1 lần
        service.spreadsheets().values().update(
            spreadsheetId=spreadsheet_id,
            range=f"'{tab_name}'!A1",
            valueInputOption="USER_ENTERED",
            body={"values": all_values},
        ).execute()

        # ── Formatting requests ────────────────────────────────────
        reqs = []

        # Kích thước cột
        reqs.append(_col_width(sheet_id, 0, 90))   # A: section label
        reqs.append(_col_width(sheet_id, 1, 220))  # B: row label
        for ci in range(num_days):
            reqs.append(_col_width(sheet_id, ci+2, 80))

        # Row heights
        reqs.append(_row_height(sheet_id, R_TITLE, 28))
        reqs.append(_row_height(sheet_id, R_DATE_HDR, 22))
        for ri in range(R_OVERALL_S, R_TOTAL + 1):
            reqs.append(_row_height(sheet_id, ri, 18))

        # Title row: merge + navy
        reqs.append(_merge(sheet_id, R_TITLE, 0, R_TITLE+1, num_cols))
        reqs.append(_cell_fmt(sheet_id, R_TITLE, 0, R_TITLE+1, num_cols,
            bg=C_NAVY, bold=True, fg=C_WHITE, size=12, halign="CENTER"))

        # Date header row: navy
        reqs.append(_cell_fmt(sheet_id, R_DATE_HDR, 0, R_DATE_HDR+1, num_cols,
            bg=C_NAVY, bold=True, fg=C_WHITE, size=10))
        reqs.append(_cell_fmt(sheet_id, R_DATE_HDR, 1, R_DATE_HDR+1, 2,
            bg=C_NAVY, bold=True, fg=C_WHITE, halign="CENTER"))

        # "Overall" label: merge A3:A(R_OVERALL_E+1), section color
        reqs.append(_merge(sheet_id, R_OVERALL_S, 0, R_OVERALL_E, 1))
        reqs.append(_cell_fmt(sheet_id, R_OVERALL_S, 0, R_OVERALL_E, 1,
            bg=C_OVERALL, bold=False, fg=C_BLACK, size=10, halign="CENTER"))
        # Ghi chữ "Overall" vào cell đầu của merge
        service.spreadsheets().values().update(
            spreadsheetId=spreadsheet_id,
            range=f"'{tab_name}'!A{R_OVERALL_S+1}",
            valueInputOption="RAW",
            body={"values": [["Overall"]]},
        ).execute()

        # Tô màu cột B (label) theo template — data cells C+ để TRẮNG
        PCT_FIELDS = {"eff_pct", "defect_rate_pct", "mc_utilization_pct"}
        for ri, (label, field, color, is_formula) in enumerate(OVERALL_ROWS):
            row_idx = R_OVERALL_S + ri
            bg      = COLOR_MAP[color]
            fg      = C_BLUE_TXT if color == "BLUE" else C_BLACK
            is_pct  = field in PCT_FIELDS
            num_fmt = {"type": "PERCENT", "pattern": "0.0%"} if is_pct else None

            # ── Cột B: tô màu theo template (vàng/xanh/xanh lá) ──
            reqs.append(_cell_fmt(sheet_id, row_idx, 1, row_idx+1, 2,
                bg=bg, bold=(color=="BLUE"), fg=fg, halign="LEFT"))

            # ── Data cells C+ : nền TRẮNG, chỉ thêm numberFormat nếu là % ──
            reqs.append(_cell_fmt(sheet_id, row_idx, 2, row_idx+1, num_cols,
                bg=C_WHITE, bold=False, fg=C_BLACK, halign="CENTER",
                num_format=num_fmt))

        # "Name" row header
        reqs.append(_cell_fmt(sheet_id, R_NAME_HDR, 0, R_NAME_HDR+1, num_cols,
            bg=C_GRAY, bold=False, fg=C_BLACK, halign="CENTER"))

        # "Defect (pcs)" section label: merge A + xanh lavender
        reqs.append(_merge(sheet_id, R_NAME_HDR, 0, R_TOTAL, 1))
        reqs.append(_cell_fmt(sheet_id, R_NAME_HDR, 0, R_TOTAL, 1,
            bg=C_DEFECT_S, bold=False, fg=C_BLACK, halign="CENTER"))
        service.spreadsheets().values().update(
            spreadsheetId=spreadsheet_id,
            range=f"'{tab_name}'!A{R_NAME_HDR+1}",
            valueInputOption="RAW",
            body={"values": [["Defect (pcs)"]]},
        ).execute()

        # Defect rows — cột B label tô vàng, data C+ trắng
        if R_DEFECT_S < R_DEFECT_E:
            reqs.append(_cell_fmt(sheet_id, R_DEFECT_S, 1, R_DEFECT_E, 2,
                bg=C_YELLOW, bold=False, fg=C_BLACK, halign="LEFT"))
            reqs.append(_cell_fmt(sheet_id, R_DEFECT_S, 2, R_DEFECT_E, num_cols,
                bg=C_WHITE, bold=False, fg=C_BLACK, halign="CENTER"))

        # Total defect row — label A+B merge tô xanh, data C+ trắng
        reqs.append(_merge(sheet_id, R_TOTAL, 0, R_TOTAL+1, 2))
        reqs.append(_cell_fmt(sheet_id, R_TOTAL, 0, R_TOTAL+1, 2,
            bg=C_BLUE, bold=True, fg=C_BLUE_TXT, halign="CENTER"))
        reqs.append(_cell_fmt(sheet_id, R_TOTAL, 2, R_TOTAL+1, num_cols,
            bg=C_WHITE, bold=False, fg=C_BLACK, halign="CENTER"))

        # Borders toàn bộ
        reqs.append(_border_req(sheet_id, R_TITLE, 0, R_TOTAL+1, num_cols))

        # Freeze 2 dòng header (KHÔNG freeze cột — xung đột với merged cells ở cột A+B)
        reqs.append({"updateSheetProperties": {
            "properties": {"sheetId": sheet_id,
                           "gridProperties": {"frozenRowCount": 2,
                                              "frozenColumnCount": 0}},
            "fields": "gridProperties.frozenRowCount,gridProperties.frozenColumnCount"
        }})

        # Landscape + fit to page
        reqs.append({"updateSpreadsheetProperties": {
            "properties": {"title": f"Machine Log {month_abbr}-{year}"},
            "fields": "title"
        }})

        _fmt(reqs, service, spreadsheet_id)

        sheet_url = (f"https://docs.google.com/spreadsheets/d/{spreadsheet_id}"
                     f"/edit#gid={sheet_id}")
        result_tabs.append({"project": proj_name, "tab": tab_name,
                             "url": sheet_url, "sheet_id": sheet_id})

    return {
        "spreadsheet_url": f"https://docs.google.com/spreadsheets/d/{spreadsheet_id}/edit",
        "tabs": result_tabs,
        "month": f"{month_abbr}-{year}",
        "total_projects": len(result_tabs),
    }


# ── Push 1 record (realtime khi nhấn Lưu) ────────────────────────
def push_to_sheet(entry, computed: dict, spreadsheet_id: str) -> dict:
    """
    Đẩy 1 record vào đúng vị trí cột ngày trong Sheet.
    Nếu Sheet chưa có layout → tạo mới theo template.
    """
    # Lấy tháng/năm từ entry.date
    y, m, d = map(int, entry.date.split("-"))
    # Fake record dict
    rec = {
        "project":               entry.project,
        "mo_color":              entry.mo_color or "",
        "date":                  entry.date,
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
        "defects":               [{"name": d.name, "qty": d.qty} for d in entry.defects],
        **computed,
    }
    # Lấy tất cả records cùng project + tháng từ... chỉ record này
    result = write_monthly_sheet([rec], y, m, spreadsheet_id)
    # Trả về URL tab của project
    tab_info = result["tabs"][0] if result["tabs"] else {}
    return {
        "url":       tab_info.get("url", ""),
        "tab":       tab_info.get("tab", ""),
        "sheet_url": tab_info.get("url", ""),
        "row":       None,
    }



def push_batch_to_sheet(records: list, spreadsheet_id: str) -> dict:
    """
    Đẩy nhiều records cùng lúc lên Google Sheet theo đúng format template.
    records: list of dict từ DB (đã có computed fields + defects list).
    Nhóm theo project + tháng, mỗi nhóm gọi write_monthly_sheet.
    """
    if not records:
        return {"pushed": 0, "errors": [], "spreadsheet_url": ""}

    # Nhóm theo (project, year, month)
    from collections import defaultdict
    groups = defaultdict(list)
    for r in records:
        try:
            y, m, _ = r["date"].split("-")
            groups[(r["project"], int(y), int(m))].append(r)
        except Exception:
            continue

    pushed = 0
    errors = []
    for (proj, year, month), recs in groups.items():
        try:
            result = write_monthly_sheet(recs, year, month, spreadsheet_id)
            pushed += len(recs)
        except Exception as ex:
            errors.append(f"{proj} {month}/{year}: {str(ex)}")

    return {
        "pushed":           pushed,
        "errors":           errors,
        "spreadsheet_url":  f"https://docs.google.com/spreadsheets/d/{spreadsheet_id}/edit",
    }

def get_sheet_info(spreadsheet_id: str) -> dict:
    service = _get_service()
    meta    = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
    tabs    = []
    for s in meta["sheets"]:
        title = s["properties"]["title"]
        try:
            res  = service.spreadsheets().values().get(
                spreadsheetId=spreadsheet_id,
                range=f"'{title}'!A:A").execute()
            rows = max(0, len(res.get("values", [])) - 2)
        except Exception:
            rows = 0
        tabs.append({"name": title, "records": rows})
    return {
        "title":          meta["properties"]["title"],
        "spreadsheet_id": spreadsheet_id,
        "url":            f"https://docs.google.com/spreadsheets/d/{spreadsheet_id}/edit",
        "tabs":           tabs,
        "total_records":  sum(t["records"] for t in tabs),
    }
