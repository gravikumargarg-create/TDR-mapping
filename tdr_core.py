"""
TDR Data Excel script - Sheet/file selection and TDR → BAN extraction.
- Asks user which sheet to use from TDR Data.xlsx (default: same folder as script).
- Supports multiple Excel files and/or multiple sheets if needed.
- Headings are not static; identifies sections by TDR number (TDR-######,
  TDR ######, TDR_###### or TDR######) and collects 9-digit BAN IDs per section.
"""

import os
import re
import shutil
import sys
from copy import copy
from datetime import datetime, timedelta

try:
    import openpyxl
    from openpyxl import load_workbook, Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
except ImportError:
    print("openpyxl is required. Install with: pip install openpyxl")
    sys.exit(1)

# Pattern: TDR number – hyphen (TDR-204035), space (TDR 203410), underscore (TDR_203410), or none
TDR_PATTERN = re.compile(r"TDR[\s\-_]*(\d{5,6})", re.IGNORECASE)
# Pattern: 9-digit BAN ID (standalone number in string)
BAN_PATTERN = re.compile(r"\b(\d{9})\b")


# Script folder; reports in report subfolder; old reports moved to report/archive
# When TDR_WEB_REPORT_FOLDER env is set (e.g. by web/Cloud Function), use that for output
BASE_FOLDER = os.path.dirname(os.path.abspath(__file__))
REPORT_FOLDER = os.environ.get("TDR_WEB_REPORT_FOLDER") or os.path.join(BASE_FOLDER, "report")
ARCHIVE_FOLDER = os.path.join(REPORT_FOLDER, "archive")
# Console styling (ANSI on Windows 10+)
if sys.platform == "win32":
    try:
        import ctypes
        kernel32 = ctypes.windll.kernel32
        kernel32.SetConsoleMode(kernel32.GetStdHandle(-11), 7)
    except Exception:
        pass
C = type("C", (), {"RESET": "\033[0m", "BOLD": "\033[1m", "GREEN": "\033[92m", "RED": "\033[91m", "YELLOW": "\033[93m", "CYAN": "\033[96m", "DIM": "\033[2m"})()
BOX = {"TL": "╔", "TR": "╗", "BL": "╚", "BR": "╝", "H": "═", "V": "║"}
W = 56
DEFAULT_TDR_EXCEL = os.path.join(BASE_FOLDER, "TDR Data.xlsx")
DEFAULT_LVT_REPORT = os.path.join(BASE_FOLDER, "LVT_RUN_25Feb_Report.xlsx")
LVT_SHEET_NAME = "BAN Wise Result"
# Failures sheet: match any variation of "BAN Wise Failures" (spaces, underscores, case)
LVT_FAILURES_SHEET_NORMALIZED = "banwisefailures"

# Comments for TDR Info column F: check_id (str) -> comment text. Add more as needed.
CHECK_ID_COMMENTS = {
    "3106": "BANs can be used after RS-122932 fix.",
}


def _normalize_sheet_name(name):
    """Return name with spaces/underscores removed and lowercased for matching."""
    if not name:
        return ""
    return "".join(c for c in str(name).lower().strip() if c not in " \t_")


def _find_failures_sheet_in_workbook(wb):
    """Return the first sheet name in wb that normalizes to BAN Wise Failures, or None."""
    target = _normalize_sheet_name(LVT_FAILURES_SHEET_NORMALIZED)
    for sheet_name in wb.sheetnames:
        if _normalize_sheet_name(sheet_name) == target:
            return sheet_name
    return None


def detect_excel_roles(wb):
    """
    Detect which roles this workbook can serve (TDR Data, LVT Report, Device Details).
    Returns dict: tdr_sheets=[], lvt_sheets=[], device_sheets=[] (sheet names).
    """
    result = {"tdr_sheets": [], "lvt_sheets": [], "device_sheets": []}
    lvt_target = _normalize_sheet_name(LVT_SHEET_NAME)  # "banwiseresult"
    for name in wb.sheetnames:
        try:
            ws = wb[name]
        except Exception:
            continue
        try:
            if extract_tdr_ban_mapping(ws):
                result["tdr_sheets"].append(name)
        except Exception:
            pass
        if _normalize_sheet_name(name) == lvt_target:
            result["lvt_sheets"].append(name)
        try:
            first_row = next(ws.iter_rows(min_row=1, max_row=1, values_only=True), None)
            if first_row and _find_column_index_in_row(list(first_row), DEVICE_DETAILS_CUSTOMER_ID_HEADERS):
                result["device_sheets"].append(name)
        except Exception:
            pass
    return result
# Folder to scan for LVT Excel files (user picks file and optionally sheet)
TDR_DELIVERY_FOLDER = os.path.join(BASE_FOLDER, "TDR deliver")


def archive_old_reports(report_folder, max_age_days=1):
    """
    Move report files and TDR folders in report_folder that are older than max_age_days
    into report_folder/archive. Creates archive folder if needed.
    """
    if not os.path.isdir(report_folder):
        return
    archive_dir = os.path.join(report_folder, "archive")
    os.makedirs(archive_dir, exist_ok=True)
    cutoff = datetime.now() - timedelta(days=max_age_days)
    cutoff_ts = cutoff.timestamp()
    for name in os.listdir(report_folder):
        if name == "archive":
            continue
        path = os.path.join(report_folder, name)
        try:
            mtime = os.path.getmtime(path)
        except OSError:
            continue
        if mtime >= cutoff_ts:
            continue
        dest = os.path.join(archive_dir, name)
        try:
            if os.path.exists(dest):
                base, ext = os.path.splitext(name) if os.path.isfile(path) else (name, "")
                if ext:
                    n = 1
                    while os.path.exists(dest):
                        dest = os.path.join(archive_dir, f"{base}_{n}{ext}")
                        n += 1
                else:
                    n = 1
                    while os.path.exists(dest):
                        dest = os.path.join(archive_dir, f"{base}_{n}")
                        n += 1
            shutil.move(path, dest)
        except (OSError, shutil.Error):
            pass


def get_excel_path():
    """Ask user for TDR Data Excel path; use default if empty."""
    print(f"\nTDR Data Excel (default: {DEFAULT_TDR_EXCEL})")
    path = input("  Path [Enter for default]: ").strip()
    if not path:
        path = DEFAULT_TDR_EXCEL
    if not os.path.isfile(path):
        print(f"  Error: File not found: {path}")
        return None
    return path


def _excel_extensions():
    return (".xlsx", ".xlsm")


def _list_lvt_excel_files(folder):
    """Return list of (full_path, filename) for files in folder whose name starts with 'LVT' (case-insensitive)."""
    if not os.path.isdir(folder):
        return []
    result = []
    for name in os.listdir(folder):
        if not name.upper().startswith("LVT"):
            continue
        if name.lower().endswith(_excel_extensions()):
            result.append((os.path.join(folder, name), name))
    return sorted(result, key=lambda x: x[1])


def ask_single_sheet_choice(sheet_names, label="workbook"):
    """Ask user to pick one sheet by number or name. Returns sheet name or None."""
    print(f"\nSheets in {label}:")
    for i, name in enumerate(sheet_names, 1):
        print(f"  {i}. {name}")
    print("  Enter sheet number or sheet name to use for BAN-wise list.")
    choice = input("  Your choice: ").strip()
    if not choice:
        return None
    if choice.isdigit():
        idx = int(choice)
        if 1 <= idx <= len(sheet_names):
            return sheet_names[idx - 1]
        print(f"  Invalid number: {idx}")
        return None
    if choice in sheet_names:
        return choice
    print(f"  Sheet not found: {choice!r}")
    return None


def get_lvt_report_file_and_sheet():
    """
    Scan TDR_DELIVERY_FOLDER for Excel files whose name starts with 'LVT'.
    Let user select one file; then if sheet 'BAN Wise Result' exists use it,
    else list all sheets and ask which sheet to use for BAN-wise list.
    Returns (file_path, sheet_name) or None.
    """
    print(f"\nLVT Report – selecting file from delivery path: {TDR_DELIVERY_FOLDER}")
    lvt_files = _list_lvt_excel_files(TDR_DELIVERY_FOLDER)
    if not lvt_files:
        print("  No Excel file starting with 'LVT' found in this folder.")
        print("  Status column will show 'Not found' for all BANs.")
        return None

    print("  LVT Excel files found:")
    for i, (path, name) in enumerate(lvt_files, 1):
        print(f"  {i}. {name}")
    print("  Enter file number or filename to use for BAN-wise list.")
    choice = input("  Your choice: ").strip()
    if not choice:
        print("  No selection. Skipping LVT report.")
        return None

    selected_path = None
    if choice.isdigit():
        idx = int(choice)
        if 1 <= idx <= len(lvt_files):
            selected_path = lvt_files[idx - 1][0]
        else:
            print(f"  Invalid number: {idx}")
            return None
    else:
        for path, name in lvt_files:
            if name == choice or name.lower() == choice.lower():
                selected_path = path
                break
        if not selected_path:
            print(f"  File not found: {choice!r}")
            return None

    # Open workbook and decide sheet
    try:
        wb = load_workbook(selected_path, read_only=True)
        sheet_names = wb.sheetnames
        wb.close()
    except Exception as e:
        print(f"  Error opening file: {e}")
        return None

    if LVT_SHEET_NAME in sheet_names:
        print(f"  Using sheet: {LVT_SHEET_NAME}")
        return (selected_path, LVT_SHEET_NAME)

    print(f"  Expected sheet {LVT_SHEET_NAME!r} not found in this workbook.")
    chosen = ask_single_sheet_choice(sheet_names, os.path.basename(selected_path))
    if not chosen:
        return None
    return (selected_path, chosen)


def get_sheet_names(excel_path):
    """Return list of sheet names in the workbook."""
    wb = load_workbook(excel_path, read_only=True)
    names = wb.sheetnames
    wb.close()
    return names


def ask_sheet_choice(sheet_names, excel_label="TDR Data"):
    """Ask user to pick one or more sheets; return list of selected sheet names."""
    print(f"\nSheets in {excel_label}:")
    for i, name in enumerate(sheet_names, 1):
        print(f"  {i}. {name}")
    print("  Enter sheet number(s) separated by comma (e.g. 4 or 1,3,5), or sheet name(s).")
    choice = input("  Your choice: ").strip()
    if not choice:
        return None
    selected = []
    # Allow numbers
    for part in (c.strip() for c in choice.split(",")):
        if part.isdigit():
            idx = int(part)
            if 1 <= idx <= len(sheet_names):
                selected.append(sheet_names[idx - 1])
            else:
                print(f"  Invalid number: {idx}")
        else:
            # Treat as sheet name
            if part in sheet_names:
                selected.append(part)
            else:
                print(f"  Sheet not found: {part!r}")
    return selected if selected else None


def ask_more_sources():
    """Ask if user needs to add more Excel files or more sheets."""
    print("\nDo you need to use another Excel file or more sheets from the same file? (y/n)")
    return input("  Your choice [n]: ").strip().lower() in ("y", "yes")


def _extract_tdr_from_cell(value):
    """If cell value contains TDR-######, return e.g. 'TDR-204035'; else None."""
    if value is None:
        return None
    s = str(value).strip()
    m = TDR_PATTERN.search(s)
    if m:
        return f"TDR-{m.group(1)}"
    return None


def _extract_bans_from_cell(value):
    """Extract all 9-digit BAN IDs from a cell (number or string with newlines). Returns set of strings."""
    bans = set()
    if value is None:
        return bans
    if isinstance(value, (int, float)):
        v = int(value)
        if 100_000_000 <= v <= 999_999_999:
            bans.add(str(v))
        return bans
    s = str(value).strip()
    for m in BAN_PATTERN.finditer(s):
        bans.add(m.group(1))
    return bans


def extract_tdr_ban_mapping(ws):
    """
    Scan sheet: find rows where any cell contains TDR-###### (section header).
    For each section (from one TDR row to just before the next), collect all
    9-digit BAN IDs from any cell. Return dict: TDR_number -> list of BAN IDs (unique).
    """
    # Find all TDR header rows: (row_1_index, tdr_id)
    tdr_rows = []  # list of (row_idx, tdr_id), 1-based
    for row in ws.iter_rows():
        row_idx = None
        for cell in row:
            if hasattr(cell, "row"):
                row_idx = cell.row
                break
        if not row_idx:
            continue
        for cell in row:
            tdr_id = _extract_tdr_from_cell(cell.value)
            if tdr_id:
                tdr_rows.append((row_idx, tdr_id))
                break

    if not tdr_rows:
        return {}

    # For each consecutive pair of TDR rows, collect BANs in that block
    result = {}
    for i, (start_row, tdr_id) in enumerate(tdr_rows):
        end_row = tdr_rows[i + 1][0] if i + 1 < len(tdr_rows) else (ws.max_row + 1)
        bans = set()
        for row in ws.iter_rows(min_row=start_row, max_row=end_row - 1):
            for cell in row:
                bans |= _extract_bans_from_cell(cell.value)
        result[tdr_id] = sorted(bans)  # sorted for consistent output

    return result


def get_tdr_section_ranges(ws):
    """
    Find all TDR section header rows and return (tdr_id, start_row, end_row) for each section.
    Section includes rows from start_row through end_row - 1 (end_row is exclusive).
    """
    tdr_rows = []
    for row in ws.iter_rows():
        row_idx = next((c.row for c in row if hasattr(c, "row")), None)
        if not row_idx:
            continue
        for cell in row:
            tdr_id = _extract_tdr_from_cell(cell.value)
            if tdr_id:
                tdr_rows.append((row_idx, tdr_id))
                break
    if not tdr_rows:
        return []
    result = []
    for i, (start_row, tdr_id) in enumerate(tdr_rows):
        end_row_exclusive = tdr_rows[i + 1][0] if i + 1 < len(tdr_rows) else (ws.max_row + 1)
        result.append((tdr_id, start_row, end_row_exclusive))
    return result


def _copy_cell_style(src_cell, dest_cell):
    """Copy font, fill, border, alignment, number_format from src_cell to dest_cell."""
    try:
        if getattr(src_cell, "font", None):
            dest_cell.font = copy(src_cell.font)
        if getattr(src_cell, "fill", None):
            dest_cell.fill = copy(src_cell.fill)
        if getattr(src_cell, "border", None):
            dest_cell.border = copy(src_cell.border)
        if getattr(src_cell, "alignment", None):
            dest_cell.alignment = copy(src_cell.alignment)
        if getattr(src_cell, "number_format", None):
            dest_cell.number_format = src_cell.number_format
        if getattr(src_cell, "protection", None):
            dest_cell.protection = copy(src_cell.protection)
    except Exception:
        pass


def _copy_sheet_range_to_workbook(src_ws, start_row, end_row_exclusive, sheet_title="Data"):
    """
    Copy all rows from src_ws (start_row through end_row_exclusive - 1) and all columns
    into a new Workbook with one sheet. Copies values, column widths, row heights, and cell formatting
    (font, fill, border, alignment, number_format). Returns the new Workbook.
    """
    wb = Workbook()
    dest_ws = wb.active
    dest_ws.title = sheet_title[:31] if sheet_title else "Data"
    max_col = src_ws.max_column
    for r in range(start_row, end_row_exclusive):
        dest_r = r - start_row + 1
        for c in range(1, max_col + 1):
            src_cell = src_ws.cell(row=r, column=c)
            dest_cell = dest_ws.cell(row=dest_r, column=c, value=src_cell.value)
            _copy_cell_style(src_cell, dest_cell)
        # Copy row height
        if r in src_ws.row_dimensions and src_ws.row_dimensions[r].height is not None:
            dest_ws.row_dimensions[dest_r].height = src_ws.row_dimensions[r].height
    for c in range(1, max_col + 1):
        letter = get_column_letter(c)
        if letter in src_ws.column_dimensions:
            dest_ws.column_dimensions[letter].width = src_ws.column_dimensions[letter].width
    return wb


TDR_HEADER_MAP = [
    ("no_of_ban", ("number of ban", "no. of ban", "no of ban", "bans needed")),
    ("account_type", ("account type", "account segment")),
    ("sub_type", ("sub type", "subtype", "account subtype")),
    ("source_plan", ("source price plan", "source protection", "uscc plan")),
    ("line_type", ("line type", "t-mobile")),
    ("target_soc", ("target plan soc", "target plan boc", "target feature code")),
    ("target_plan_name", ("target plan name")),
    ("no_of_lines", ("no of lines")),
    ("owner", ("owner")),
    ("comment", ("comment")),
    ("bans_cell", ("od", "cid", "bans")),
]


def _find_tdr_section_header_row(ws, start_row, end_row):
    for r in range(start_row, end_row):
        row_tuple = next(ws.iter_rows(min_row=r, max_row=r, values_only=True), None)
        if not row_tuple:
            continue
        row_vals = list(row_tuple)
        has_no_ban = _find_column_index_in_row(row_vals, ("number of ban", "no. of ban", "no of ban", "bans needed")) is not None
        has_account = _find_column_index_in_row(row_vals, ("account type", "account segment", "account subtype")) is not None
        if not (has_no_ban and has_account):
            continue
        col_map = []
        for field_name, keywords in TDR_HEADER_MAP:
            idx = _find_column_index_in_row(row_vals, keywords)
            if idx is not None:
                col_map.append((field_name, idx))
        if any(f[0] == "no_of_ban" for f in col_map) and any(f[0] == "bans_cell" for f in col_map):
            return (r, col_map)
    return (None, [])


def _extract_bans_list_from_cell(value):
    if value is None:
        return []
    return list(BAN_PATTERN.findall(str(value).strip()))


def _get_merged_ranges(ws):
    """Return list of (min_row, min_col, max_row, max_col) for each merged range. Requires ws not in read_only mode."""
    try:
        return [
            (r.min_row, r.min_col, r.max_row, r.max_col)
            for r in ws.merged_cells.ranges
        ]
    except Exception:
        return []


def _get_cell_value_respecting_merge(ws, row, col, merged_ranges):
    """Return cell value; if (row, col) is inside a merged range, return the top-left cell value."""
    for (min_row, min_col, max_row, max_col) in merged_ranges:
        if min_row <= row <= max_row and min_col <= col <= max_col:
            return ws.cell(row=min_row, column=min_col).value
    return ws.cell(row=row, column=col).value


def extract_tdr_sections_with_rows(ws):
    tdr_rows = []
    for row in ws.iter_rows():
        row_idx = next((c.row for c in row if hasattr(c, "row")), None)
        if not row_idx:
            continue
        for cell in row:
            tdr_id = _extract_tdr_from_cell(cell.value)
            if tdr_id:
                tdr_rows.append((row_idx, tdr_id))
                break
    if not tdr_rows:
        return {}
    merged_ranges = _get_merged_ranges(ws)
    result = {}
    for i, (start_row, tdr_id) in enumerate(tdr_rows):
        end_row = tdr_rows[i + 1][0] if i + 1 < len(tdr_rows) else (ws.max_row + 1)
        header_row_idx, col_map = _find_tdr_section_header_row(ws, start_row, end_row)
        if header_row_idx is None or not col_map:
            continue
        col_by_field = dict(col_map)
        data_start = header_row_idx + 1
        rows_out = []
        for r in range(data_start, end_row):
            row_dict = {}
            for field_name, col_idx in col_by_field.items():
                val = _get_cell_value_respecting_merge(ws, r, col_idx, merged_ranges)
                if field_name == "bans_cell":
                    row_dict["bans_list"] = _extract_bans_list_from_cell(val)
                else:
                    row_dict[field_name] = val
            if "bans_list" not in row_dict:
                row_dict["bans_list"] = []
            no_ban = row_dict.get("no_of_ban")
            if no_ban is not None:
                try:
                    row_dict["no_of_ban"] = int(no_ban)
                except (TypeError, ValueError):
                    row_dict["no_of_ban"] = no_ban
            rows_out.append(row_dict)
        if rows_out:
            result[tdr_id] = rows_out
    return result


def _copy_sheet_into_workbook(src_ws, dest_wb, sheet_name):
    """Copy all cell values from src_ws into a new sheet in dest_wb with the given name."""
    dest_ws = dest_wb.create_sheet(title=sheet_name)
    for row in src_ws.iter_rows(values_only=True):
        dest_ws.append(row)
    return dest_ws


# Device Details sheet: columns to store as text to avoid long-number truncation/scientific notation
DEVICE_DETAILS_SHEET_NAME = "Device Details"
DEVICE_DETAILS_CUSTOMER_ID_HEADERS = ("customer_id", "customer id", "ban", "bans", "lgc_customer id")
DEVICE_DETAILS_TEXT_COLUMNS = (
    "customer_id", "customer id", "msisdn", "imei", "esn", "eid", "uiccid", "uimsi", "timsi",
    "device_model", "device_lock_status",
)


def _load_device_details(path, sheet_name=None):
    """
    Load device details Excel. Returns (headers, list of row lists, customer_id_col_1based) or None.
    Uses first sheet with CUSTOMER_ID (or similar) in header if sheet_name not given.
    """
    if not path or not os.path.isfile(path):
        return None
    try:
        wb = load_workbook(path, read_only=True, data_only=True)
        sheet_to_use = sheet_name
        if not sheet_to_use:
            for name in wb.sheetnames:
                ws = wb[name]
                first_row = next(ws.iter_rows(min_row=1, max_row=1, values_only=True), None)
                if first_row and _find_column_index_in_row(list(first_row), DEVICE_DETAILS_CUSTOMER_ID_HEADERS):
                    sheet_to_use = name
                    break
        if not sheet_to_use:
            wb.close()
            return None
        ws = wb[sheet_to_use]
        rows = list(ws.iter_rows(values_only=True))
        wb.close()
        if not rows:
            return None
        headers = [str(h).strip() if h is not None else "" for h in rows[0]]
        cust_col = _find_column_index_in_row(headers, DEVICE_DETAILS_CUSTOMER_ID_HEADERS)
        if cust_col is None:
            cust_col = 1
        return (headers, rows[1:], cust_col)
    except Exception:
        return None


def _add_device_details_sheet_to_workbook(wb, device_data, bans_set):
    """
    Add sheet 'Device Details' to wb with rows from device_data where CUSTOMER_ID is in bans_set.
    device_data = (headers, list of row tuples, customer_id_col_1based).
    Format long-number columns as text (@) to avoid scattering/truncation.
    """
    if not device_data or not bans_set:
        return
    headers, rows, cust_col = device_data
    cust_idx = cust_col - 1
    # Normalize BANs for match
    bans_normalized = set()
    for b in bans_set:
        n = _normalize_ban(b)
        if n:
            bans_normalized.add(n)
        if isinstance(b, str):
            bans_normalized.add(b.strip())
        else:
            bans_normalized.add(str(b).strip())
    matching = []
    for row in rows:
        if not row or len(row) <= cust_idx:
            continue
        val = row[cust_idx]
        n = _normalize_ban(val)
        if n and n in bans_normalized:
            matching.append(row)
            continue
        s = (str(val).strip() if val is not None else "")
        if s in bans_normalized:
            matching.append(row)
    # Create sheet
    ws = wb.create_sheet(title=DEVICE_DETAILS_SHEET_NAME[:31])
    header_row = list(headers)
    ws.append(header_row)
    for row in matching:
        # Write as strings for long-number columns to avoid Excel reformatting
        out = []
        for i, v in enumerate(row):
            if i >= len(header_row):
                out.append(v)
                continue
            h = (header_row[i] or "").lower().replace(" ", "_")
            if any(t in h for t in ("customer_id", "msisdn", "imei", "esn", "eid", "uiccid", "uimsi", "timsi")):
                out.append(str(v) if v is not None else "")
            else:
                out.append(v)
        ws.append(out)
    # Set text format for long-number columns
    for col_idx, h in enumerate(header_row, 1):
        hl = (h or "").lower().replace(" ", "_")
        if any(t in hl for t in ("customer_id", "msisdn", "imei", "esn", "eid", "uiccid", "uimsi", "timsi")):
            for r in range(2, ws.max_row + 1):
                ws.cell(row=r, column=col_idx).number_format = "@"
    # Header styling
    thin_border = Border(left=Side(style="thin"), right=Side(style="thin"), top=Side(style="thin"), bottom=Side(style="thin"))
    header_fill = PatternFill(start_color="2F5496", end_color="2F5496", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF")
    for col in range(1, len(header_row) + 1):
        c = ws.cell(row=1, column=col)
        c.fill = header_fill
        c.font = header_font
        c.border = thin_border
    for r in range(2, ws.max_row + 1):
        for col in range(1, len(header_row) + 1):
            ws.cell(row=r, column=col).border = thin_border


def _normalize_ban(value):
    """Return 9-digit BAN as string, or None if not a valid BAN."""
    if value is None:
        return None
    if isinstance(value, (int, float)):
        v = int(value)
        if 100_000_000 <= v <= 999_999_999:
            return str(v)
        return None
    s = str(value).strip()
    if s.isdigit() and len(s) == 9:
        return s
    m = BAN_PATTERN.search(s)
    return m.group(1) if m else None


def _find_column_index_in_row(row_values, header_keywords):
    """Return 1-based column index of first cell in row that contains any keyword (case-insensitive)."""
    keywords = [k.lower() for k in header_keywords]
    for col_idx, cell_value in enumerate(row_values or (), 1):
        if cell_value is None:
            continue
        val = str(cell_value).strip().lower()
        if any(kw in val for kw in keywords):
            return col_idx
    return None


def _find_header_row(ws, max_look=10):
    """Find first row that has both BAN-like and Status-like headers. Returns (1-based row index, row values) or (None, None)."""
    ban_kw = ("ban", "bans", "ban id", "account")
    status_kw = ("status", "result", "lvt", "passed", "delivered", "outcome", "state", "execution")
    for r in range(1, min(ws.max_row + 1, max_look + 1)):
        row_tuple = next(ws.iter_rows(min_row=r, max_row=r, values_only=True), None)
        if not row_tuple:
            continue
        row_vals = list(row_tuple)
        has_ban = _find_column_index_in_row(row_vals, ban_kw) is not None
        has_status = _find_column_index_in_row(row_vals, status_kw) is not None
        if has_ban and has_status:
            return (r, row_vals)
    return (None, None)


def _build_ban_to_status_from_sheet(ws):
    """Build dict BAN (str) -> Status from BAN Wise Result sheet. Returns dict."""
    ban_to_status = {}

    def collect(ban_col, status_col, data_start_row):
        for row in ws.iter_rows(min_row=data_start_row, values_only=True):
            if not row or len(row) < 2:
                continue
            ban_val = row[ban_col - 1] if ban_col <= len(row) else None
            status_val = row[status_col - 1] if status_col <= len(row) else None
            ban_strs = set()
            if ban_val is not None:
                s = str(ban_val).strip()
                for m in BAN_PATTERN.finditer(s):
                    ban_strs.add(m.group(1))
                if not ban_strs:
                    single = _normalize_ban(ban_val)
                    if single:
                        ban_strs.add(single)
            status_str = (str(status_val).strip() if status_val else "") if status_val is not None else ""
            for ban_str in ban_strs:
                ban_to_status[ban_str] = status_str

    # Try header-based detection first
    header_row_idx, header_row_vals = _find_header_row(ws)
    if header_row_idx is not None and header_row_vals:
        ban_col = _find_column_index_in_row(header_row_vals, ("ban", "bans", "ban id", "account"))
        status_col = _find_column_index_in_row(header_row_vals, ("status", "result", "lvt", "passed", "delivered", "outcome", "state"))
        if ban_col is not None and status_col is not None:
            collect(ban_col, status_col, header_row_idx + 1)
            return ban_to_status

    # Fallback: BAN Wise Result is often Column A = BAN, Column B = Status (no headers or different headers)
    data_start = 1
    first_row = next(ws.iter_rows(min_row=1, max_row=1, values_only=True), None)
    if first_row and len(first_row) >= 1:
        first_cell = first_row[0]
        if first_cell is not None and _normalize_ban(first_cell) is None:
            s = str(first_cell).strip().lower()
            if s and not s.isdigit():
                data_start = 2
    collect(1, 2, data_start)  # column A = BAN, column B = Status
    return ban_to_status


def _fill_tdr_info_status_column(tdr_ws, ban_to_status):
    """Fill column C (Status) in TDR Info sheet by looking up BAN in ban_to_status. Uses 'Not found' when BAN not in BAN Wise Result."""
    tdr_ws.cell(row=1, column=3, value="Status")
    filled = 0
    for r in range(2, tdr_ws.max_row + 1):
        ban_cell = tdr_ws.cell(row=r, column=2).value
        ban_str = _normalize_ban(ban_cell)
        status = ban_to_status.get(ban_str, "Not found") if ban_str else "Not found"
        if status and status != "Not found":
            filled += 1
        tdr_ws.cell(row=r, column=3, value=status)
    return filled


def _find_failures_header_row(ws, max_look=10):
    """Find row with customer/BAN, description, and check_id. Returns (1-based row index, row values) or (None, None)."""
    cust_kw = ("lgc_customer", "customer id", "customer", "ban", "bans", "account")
    desc_kw = ("description", "failure", "desc")
    check_kw = ("check_id", "check id", "checkid")
    for r in range(1, min(ws.max_row + 1, max_look + 1)):
        row_tuple = next(ws.iter_rows(min_row=r, max_row=r, values_only=True), None)
        if not row_tuple:
            continue
        row_vals = list(row_tuple)
        if (_find_column_index_in_row(row_vals, cust_kw) is not None and
                _find_column_index_in_row(row_vals, desc_kw) is not None and
                _find_column_index_in_row(row_vals, check_kw) is not None):
            return (r, row_vals)
    return (None, None)


def _build_ban_to_failures_from_sheet(ws):
    """Build dict BAN (str) -> list of (description, check_id) from BAN Wise Failures sheet."""
    from collections import defaultdict
    ban_to_failures = defaultdict(list)

    cust_kw = ("lgc_customer", "customer id", "customer", "ban", "bans", "account")
    desc_kw = ("description", "failure", "desc")
    check_kw = ("check_id", "check id", "checkid")

    header_row_idx, header_row_vals = _find_failures_header_row(ws)
    if header_row_idx is None or not header_row_vals:
        return dict(ban_to_failures)

    cust_col = _find_column_index_in_row(header_row_vals, cust_kw)
    desc_col = _find_column_index_in_row(header_row_vals, desc_kw)
    check_col = _find_column_index_in_row(header_row_vals, check_kw)
    if not all((cust_col, desc_col, check_col)):
        return dict(ban_to_failures)

    for row in ws.iter_rows(min_row=header_row_idx + 1, values_only=True):
        if not row or len(row) < max(cust_col, desc_col, check_col):
            continue
        cust_val = row[cust_col - 1]
        desc_val = row[desc_col - 1]
        check_val = row[check_col - 1]
        ban_strs = set()
        if cust_val is not None:
            s = str(cust_val).strip()
            for m in BAN_PATTERN.finditer(s):
                ban_strs.add(m.group(1))
            if not ban_strs:
                single = _normalize_ban(cust_val)
                if single:
                    ban_strs.add(single)
        desc_str = (str(desc_val).strip() if desc_val is not None else "") or ""
        if check_val is None:
            check_str = ""
        else:
            s = str(check_val).strip()
            check_str = str(int(float(check_val))) if s.replace(".", "").isdigit() else s
        for ban_str in ban_strs:
            ban_to_failures[ban_str].append((desc_str, check_str))

    return dict(ban_to_failures)


def _fill_tdr_info_failure_columns(tdr_ws, ban_to_failures):
    """Fill columns D (Failure Description), E (Check ID), F (Comments). For Passed status use N/A; else look up BAN in ban_to_failures."""
    tdr_ws.cell(row=1, column=4, value="Failure Description")
    tdr_ws.cell(row=1, column=5, value="Check ID")
    tdr_ws.cell(row=1, column=6, value="Comments")
    for r in range(2, tdr_ws.max_row + 1):
        status_cell = tdr_ws.cell(row=r, column=3).value
        status_str = (str(status_cell).strip().lower() if status_cell else "") or ""
        if status_str == "passed":
            tdr_ws.cell(row=r, column=4, value="N/A")
            tdr_ws.cell(row=r, column=5, value="N/A")
            tdr_ws.cell(row=r, column=6, value="")
            continue
        ban_cell = tdr_ws.cell(row=r, column=2).value
        ban_str = _normalize_ban(ban_cell)
        if not ban_str or ban_str not in ban_to_failures:
            tdr_ws.cell(row=r, column=4, value="Not found")
            tdr_ws.cell(row=r, column=5, value="Not found")
            tdr_ws.cell(row=r, column=6, value="")
            continue
        pairs = ban_to_failures[ban_str]
        descriptions = [p[0] for p in pairs if p[0]]
        check_ids = [p[1] for p in pairs if p[1]]
        tdr_ws.cell(row=r, column=4, value="; ".join(descriptions) if descriptions else "Not found")
        tdr_ws.cell(row=r, column=5, value=", ".join(check_ids) if check_ids else "Not found")
        # Column F: comments for known check IDs (e.g. 3106 -> "BANs can be used after RS-122932 fix.")
        comments = []
        for cid in (check_ids or []):
            cid_str = str(cid).strip()
            if cid_str and cid_str in CHECK_ID_COMMENTS:
                comments.append(CHECK_ID_COMMENTS[cid_str])
        tdr_ws.cell(row=r, column=6, value="; ".join(comments) if comments else "")


def _format_tdr_info_sheet(ws):
    """Apply formatting to TDR Info sheet: header style, column widths, status colors, borders."""
    thin_border = Border(left=Side(style="thin"), right=Side(style="thin"), top=Side(style="thin"), bottom=Side(style="thin"))
    header_fill = PatternFill(start_color="2F5496", end_color="2F5496", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF")
    passed_fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
    failed_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
    notfound_fill = PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid")
    num_cols = max(6, ws.max_column)
    for col in range(1, num_cols + 1):
        c = ws.cell(row=1, column=col)
        c.fill = header_fill
        c.font = header_font
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = thin_border
    ws.column_dimensions["A"].width = 18
    ws.column_dimensions["B"].width = 14
    ws.column_dimensions["C"].width = 14
    ws.column_dimensions["D"].width = 48
    ws.column_dimensions["E"].width = 14
    ws.column_dimensions["F"].width = 36
    for r in range(2, ws.max_row + 1):
        for col in range(1, num_cols + 1):
            cell = ws.cell(row=r, column=col)
            cell.border = thin_border
            cell.alignment = Alignment(wrap_text=True, vertical="top")
            if col == 3:
                val = (cell.value or "").strip().lower()
                if val == "passed":
                    cell.fill = passed_fill
                elif val == "failed":
                    cell.fill = failed_fill
                elif val == "not found":
                    cell.fill = notfound_fill
    ws.freeze_panes = "A2"


def _build_tdr_summary(all_rows, ban_to_status):
    """Return list of (TDR, total, passed, failed, not_found, tdr_status). TDR status: Passed, Failed, or Partial (some not found/OOS)."""
    from collections import defaultdict
    counts = defaultdict(lambda: {"passed": 0, "failed": 0, "not_found": 0})
    for tdr, ban in all_rows:
        ban_str = _normalize_ban(ban)
        status = ban_to_status.get(ban_str, "Not found") if ban_str else "Not found"
        s = str(status).strip().lower()
        if status == "Not found" or s == "not found":
            counts[tdr]["not_found"] += 1
        elif s == "passed":
            counts[tdr]["passed"] += 1
        elif s == "failed":
            counts[tdr]["failed"] += 1
    result = []
    for tdr in sorted(counts.keys()):
        c = counts[tdr]
        total = c["passed"] + c["failed"] + c["not_found"]
        tdr_status = "Failed" if c["failed"] > 0 else ("Partial" if c["not_found"] > 0 else "Passed")
        result.append((tdr, total, c["passed"], c["failed"], c["not_found"], tdr_status))
    return result


def _add_tdr_summary_sheet(wb, tdr_summary_rows):
    """Add 'TDR Summary' sheet with TDR-wise counts and status."""
    ws = wb.create_sheet(title="TDR Summary", index=1)
    ws.append(["TDR", "Total BANs", "Passed", "Failed", "Not found", "TDR Status"])
    for row in tdr_summary_rows:
        ws.append(list(row))
    thin_border = Border(left=Side(style="thin"), right=Side(style="thin"), top=Side(style="thin"), bottom=Side(style="thin"))
    header_fill = PatternFill(start_color="2F5496", end_color="2F5496", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF")
    passed_fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
    failed_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
    partial_fill = PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid")
    for col in range(1, 7):
        c = ws.cell(row=1, column=col)
        c.fill = header_fill
        c.font = header_font
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = thin_border
    for i, w in enumerate([18, 12, 10, 10, 12, 14], 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    for r in range(2, ws.max_row + 1):
        for col in range(1, 7):
            cell = ws.cell(row=r, column=col)
            cell.border = thin_border
            if col == 6:
                val = (cell.value or "").strip()
                if val == "Passed":
                    cell.fill = passed_fill
                elif val == "Failed":
                    cell.fill = failed_fill
                elif val == "Partial":
                    cell.fill = partial_fill
    ws.freeze_panes = "A2"


def _safe_tdr_filename(tdr_id):
    """Return a safe filename for the TDR (e.g. TDR-204119 -> TDR-204119.xlsx)."""
    safe = re.sub(r'[\\/:*?"<>|]', "_", str(tdr_id).strip())
    return f"{safe}.xlsx" if safe else "TDR.xlsx"


def _row_val(v):
    if v is None:
        return ""
    return str(v).strip() if v != "" else ""


def _write_one_excel_per_tdr(all_rows, ban_to_status, output_folder, tdr_sections_data=None):
    from collections import defaultdict
    os.makedirs(output_folder, exist_ok=True)
    written = []
    all_tdrs = sorted(set(t for t, _ in all_rows))
    if tdr_sections_data:
        out_headers = ["SCN No", "NO. of BAN", "Account Type", "Sub Type", "USCC plan", "T-Mobile Plan SOC &Name", "USCC Plan", "Target Plan SOC &Name", "No of Lines", "BANs", "Comment"]
        for tdr in all_tdrs:
            if tdr not in tdr_sections_data:
                continue
            section_rows = tdr_sections_data[tdr]
            wb = Workbook()
            ws = wb.active
            ws.title = tdr[:31]
            ws.append(out_headers)
            for i, row_dict in enumerate(section_rows, 1):
                scn = f"SCN{i:02d}"
                no_ban = row_dict.get("no_of_ban", "")
                bans_list = row_dict.get("bans_list") or []
                provided_count = len(bans_list)
                needed = no_ban if isinstance(no_ban, int) else (int(no_ban) if str(no_ban).strip().isdigit() else 0)
                comment = _row_val(row_dict.get("comment"))
                if provided_count >= needed or needed == 0:
                    comment = ""
                bans_str = "\n".join(bans_list) if bans_list else ("N/A" if needed else "0")
                if not bans_str and needed:
                    bans_str = "N/A"
                ws.append([scn, no_ban, _row_val(row_dict.get("account_type")), _row_val(row_dict.get("sub_type")), _row_val(row_dict.get("source_plan")), _row_val(row_dict.get("line_type")), _row_val(row_dict.get("target_soc")), _row_val(row_dict.get("target_plan_name")), _row_val(row_dict.get("no_of_lines")), bans_str, comment])
            _format_tdr_per_sheet_wide(ws)
            path = os.path.join(output_folder, _safe_tdr_filename(tdr))
            try:
                wb.save(path)
                written.append(path)
            except PermissionError:
                pass
        for tdr in all_tdrs:
            if tdr in tdr_sections_data:
                continue
            by_tdr = defaultdict(list)
            for t, ban in all_rows:
                if t != tdr:
                    continue
                ban_str = _normalize_ban(ban)
                status = ban_to_status.get(ban_str, "Not found") if ban_str else "Not found"
                by_tdr[tdr].append((tdr, ban, status))
            if not by_tdr.get(tdr):
                continue
            wb = Workbook()
            ws = wb.active
            ws.title = tdr[:31]
            ws.append(["TDR", "BAN", "Status"])
            for r in by_tdr[tdr]:
                ws.append(list(r))
            _format_tdr_info_sheet(ws)
            path = os.path.join(output_folder, _safe_tdr_filename(tdr))
            try:
                wb.save(path)
                written.append(path)
            except PermissionError:
                pass
        return written
    by_tdr = defaultdict(list)
    for tdr, ban in all_rows:
        ban_str = _normalize_ban(ban)
        status = ban_to_status.get(ban_str, "Not found") if ban_str else "Not found"
        by_tdr[tdr].append((tdr, ban, status))
    for tdr in sorted(by_tdr.keys()):
        rows = by_tdr[tdr]
        wb = Workbook()
        ws = wb.active
        ws.title = tdr[:31]
        ws.append(["TDR", "BAN", "Status"])
        for r in rows:
            ws.append(list(r))
        _format_tdr_info_sheet(ws)
        path = os.path.join(output_folder, _safe_tdr_filename(tdr))
        try:
            wb.save(path)
            written.append(path)
        except PermissionError:
            pass
    return written


def _format_tdr_per_sheet_wide(ws):
    thin_border = Border(left=Side(style="thin"), right=Side(style="thin"), top=Side(style="thin"), bottom=Side(style="thin"))
    header_fill = PatternFill(start_color="2F5496", end_color="2F5496", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF")
    for col in range(1, 12):
        c = ws.cell(row=1, column=col)
        c.fill = header_fill
        c.font = header_font
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = thin_border
    widths = [10, 12, 12, 10, 22, 22, 14, 24, 12, 18, 30]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = min(w, 50)
    for r in range(2, ws.max_row + 1):
        for col in range(1, 12):
            ws.cell(row=r, column=col).border = thin_border
    ws.freeze_panes = "A2"


def run_extraction_and_report(all_sources, output_excel=None, lvt_report_path=None, lvt_sheet_name=None, device_details_path=None, device_details_sheet_name=None):
    """
    For each (file, sheets) in all_sources, extract TDR→BAN mapping and print.
    Saves one Excel with TDR Info sheet + TDR Summary when output_excel is set (Status/failures filled from LVT; BAN Wise Result sheet not copied).
    lvt_sheet_name: sheet to use in LVT workbook for BAN-wise list; defaults to LVT_SHEET_NAME if None.
    device_details_path: optional Excel with CUSTOMER_ID and device columns; adds "Device Details" sheet to each TDR-wise Excel for matching BANs.
    """
    wb = None
    all_rows = []
    tdr_sections_data = {}
    tdr_excel_folder = os.path.join(REPORT_FOLDER, datetime.now().strftime("%Y%m%d") + "_TDR")
    per_tdr_files = set()  # paths written (one Excel per TDR; same TDR in multiple sheets overwrites)
    os.makedirs(tdr_excel_folder, exist_ok=True)

    for excel_path, sheet_names in all_sources:
        # Load without read_only so merged_cells are available for extract_tdr_sections_with_rows
        wb = load_workbook(excel_path, read_only=False, data_only=True)
        for sheet_name in sheet_names:
            if sheet_name not in wb.sheetnames:
                continue
            ws = wb[sheet_name]
            mapping = extract_tdr_ban_mapping(ws)
            sections_rows = extract_tdr_sections_with_rows(ws)
            for tdr_id, rows in sections_rows.items():
                tdr_sections_data[tdr_id] = rows
            if not mapping:
                continue
            for tdr, bans in mapping.items():
                for ban in bans:
                    all_rows.append((tdr, ban))
            # Copy each TDR section from this sheet into a separate Excel (all columns, all rows for that TDR)
            for tdr_id, start_row, end_row_exclusive in get_tdr_section_ranges(ws):
                try:
                    copy_wb = _copy_sheet_range_to_workbook(
                        ws, start_row, end_row_exclusive, sheet_title=tdr_id
                    )
                    path = os.path.join(tdr_excel_folder, _safe_tdr_filename(tdr_id))
                    copy_wb.save(path)
                    per_tdr_files.add(path)
                except PermissionError:
                    pass
        if wb:
            wb.close()
            wb = None

    # Add "Device Details" sheet to each TDR-wise Excel if user provided device details file
    if all_rows and device_details_path and os.path.isfile(device_details_path):
        device_data = _load_device_details(device_details_path, device_details_sheet_name)
        if device_data:
            for tdr_id in sorted(set(t for t, _ in all_rows)):
                path = os.path.join(tdr_excel_folder, _safe_tdr_filename(tdr_id))
                if not os.path.isfile(path):
                    continue
                try:
                    wb = load_workbook(path, read_only=False)
                    bans_for_tdr = {ban for t, ban in all_rows if t == tdr_id}
                    _add_device_details_sheet_to_workbook(wb, device_data, bans_for_tdr)
                    wb.save(path)
                    wb.close()
                except (PermissionError, Exception):
                    pass

    # Single Excel: TDR Info sheet + TDR Summary only (no BAN Wise Result copy; Status filled from LVT)
    if output_excel and all_rows:
        out_wb = Workbook()
        out_ws = out_wb.active
        out_ws.title = "TDR Info"
        out_ws.append(["TDR", "BAN", "Status"])
        for row in all_rows:
            out_ws.append(list(row))  # Status filled in column C below
        ban_to_status = {}
        ban_to_failures = {}
        sheet_to_use = (lvt_sheet_name or LVT_SHEET_NAME) if lvt_report_path else None
        if lvt_report_path and os.path.isfile(lvt_report_path) and sheet_to_use:
            try:
                lvt_wb = load_workbook(lvt_report_path, read_only=True, data_only=True)
                if sheet_to_use in lvt_wb.sheetnames:
                    ban_to_status = _build_ban_to_status_from_sheet(lvt_wb[sheet_to_use])
                failures_sheet_name = _find_failures_sheet_in_workbook(lvt_wb)
                if failures_sheet_name is not None:
                    ban_to_failures = _build_ban_to_failures_from_sheet(lvt_wb[failures_sheet_name])
                lvt_wb.close()
            except Exception:
                pass
        summary = {"total": len(all_rows), "passed": 0, "failed": 0, "not_found": 0}
        _fill_tdr_info_status_column(out_ws, ban_to_status)
        _fill_tdr_info_failure_columns(out_ws, ban_to_failures)
        for _tdr, ban in all_rows:
            ban_str = _normalize_ban(ban)
            status = ban_to_status.get(ban_str, "Not found") if ban_str else "Not found"
            if status == "Not found":
                summary["not_found"] += 1
            elif str(status).strip().lower() == "passed":
                summary["passed"] += 1
            elif str(status).strip().lower() == "failed":
                summary["failed"] += 1
        tdr_summary_rows = _build_tdr_summary(all_rows, ban_to_status)
        _add_tdr_summary_sheet(out_wb, tdr_summary_rows)
        summary["tdr_passed"] = sum(1 for _r in tdr_summary_rows if _r[5] == "Passed")
        summary["tdr_failed"] = sum(1 for _r in tdr_summary_rows if _r[5] == "Failed")
        summary["tdr_partial"] = sum(1 for _r in tdr_summary_rows if _r[5] == "Partial")
        _format_tdr_info_sheet(out_ws)
        summary["per_tdr_folder"] = tdr_excel_folder
        summary["per_tdr_count"] = len(per_tdr_files)
        try:
            out_wb.save(output_excel)
            return (output_excel, summary)
        except PermissionError:
            print(f"\n  Permission denied: {output_excel}")
            print("  Close the file if it is open in Excel, then run again.")
            return (None, None)
    return (None, None)


def _box_title(title):
    print(f"{C.CYAN}{C.BOLD}{BOX['TL']}{BOX['H'] * (W - 2)}{BOX['TR']}{C.RESET}")
    print(f"{C.CYAN}{BOX['V']}{C.RESET} {C.BOLD}{title:<{W - 4}}{C.RESET} {C.CYAN}{BOX['V']}{C.RESET}")
    print(f"{C.CYAN}{BOX['BL']}{BOX['H'] * (W - 2)}{BOX['BR']}{C.RESET}")


def main():
    _box_title("TDR Data Excel – Sheet & file selection")
    print(f"\n{C.DIM}Working folder:{C.RESET} {BASE_FOLDER}")

    # 1) TDR Data Excel path
    excel_path = get_excel_path()
    if not excel_path:
        return

    # 2) Which sheet(s) to use from this file
    sheet_names = get_sheet_names(excel_path)
    if not sheet_names:
        print("  No sheets found in workbook.")
        return

    selected_sheets = ask_sheet_choice(sheet_names, os.path.basename(excel_path))
    if not selected_sheets:
        print("  No sheet selected. Exiting.")
        return

    # 3) Optional: more Excel files or more sheets
    all_sources = [(excel_path, selected_sheets)]

    while ask_more_sources():
        more_path = input("  Path to another Excel file: ").strip()
        if not more_path or not os.path.isfile(more_path):
            print("  Invalid or missing file. Skipping.")
            continue
        more_sheets = get_sheet_names(more_path)
        if not more_sheets:
            print("  No sheets in that workbook. Skipping.")
            continue
        more_selected = ask_sheet_choice(more_sheets, os.path.basename(more_path))
        if more_selected:
            all_sources.append((more_path, more_selected))

    # 4) LVT report – user selects file from TDR deliver folder, then sheet if needed
    lvt_choice = get_lvt_report_file_and_sheet()
    lvt_report_path = lvt_choice[0] if lvt_choice else None
    lvt_sheet_name = lvt_choice[1] if lvt_choice else None

    # Extract TDR → BAN mapping and save single Excel to report folder
    os.makedirs(REPORT_FOLDER, exist_ok=True)
    archive_old_reports(REPORT_FOLDER, max_age_days=1)
    _ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    out_excel = os.path.join(REPORT_FOLDER, f"TDR_BAN_Report_{_ts}.xlsx")
    saved_path, summary = run_extraction_and_report(
        all_sources,
        output_excel=out_excel,
        lvt_report_path=lvt_report_path,
        lvt_sheet_name=lvt_sheet_name,
    )
    if saved_path:
        total = summary.get("total", 0) if summary else 0
        pct = lambda n: f"({100 * n / total:.1f}%)" if total else "(-)"
        print(f"\n{C.GREEN}{C.BOLD}{BOX['TL']}{BOX['H'] * (W - 2)}{BOX['TR']}{C.RESET}")
        print(f"{C.GREEN}{BOX['V']}{C.RESET} {C.BOLD}Report status:{C.RESET} Successful")
        print(f"{C.GREEN}{BOX['V']}{C.RESET} {C.BOLD}Report file:{C.RESET} {os.path.basename(saved_path)}")
        print(f"{C.GREEN}{BOX['V']}{C.RESET} {C.BOLD}Report saved to path:{C.RESET} {REPORT_FOLDER}")
        if summary:
            print(f"{C.GREEN}{BOX['V']}{C.RESET}")
            print(f"{C.GREEN}{BOX['V']}{C.RESET} {C.BOLD}Summary{C.RESET}  Total BANs: {C.BOLD}{total}{C.RESET}")
            print(f"{C.GREEN}{BOX['V']}{C.RESET} {C.DIM}Below are the BANs status:{C.RESET}")
            print(f"{C.GREEN}{BOX['V']}{C.RESET}   {C.GREEN}Passed:{C.RESET}   {summary['passed']:>3}  {pct(summary['passed'])}")
            print(f"{C.GREEN}{BOX['V']}{C.RESET}   {C.RED}Failed:{C.RESET}   {summary['failed']:>3}  {pct(summary['failed'])}")
            print(f"{C.GREEN}{BOX['V']}{C.RESET}   {C.YELLOW}Not found:{C.RESET} {summary['not_found']:>3}  {pct(summary['not_found'])}")
            if "tdr_passed" in summary:
                print(f"{C.GREEN}{BOX['V']}{C.RESET} {C.DIM}TDR-wise:{C.RESET} {C.GREEN}Passed:{C.RESET} {summary['tdr_passed']}  {C.RED}Failed:{C.RESET} {summary['tdr_failed']}  {C.YELLOW}Partial:{C.RESET} {summary['tdr_partial']} (Partial = some BANs not in LVT / OOS)")
            if summary.get("per_tdr_count"):
                print(f"{C.GREEN}{BOX['V']}{C.RESET} {C.DIM}One Excel per TDR:{C.RESET} {summary['per_tdr_count']} file(s) saved to {summary.get('per_tdr_folder', '')}")
        print(f"{C.GREEN}{BOX['BL']}{BOX['H'] * (W - 2)}{BOX['BR']}{C.RESET}")
    return all_sources


if __name__ == "__main__":
    main()
