from pathlib import Path
from zipfile import BadZipFile
from collections import defaultdict
import csv
import html
import os
import secrets
import sqlite3
import string
import uuid

from flask import Flask, redirect, request, send_file, session
from werkzeug.security import check_password_hash, generate_password_hash

import xlrd
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

try:
    import psycopg
    from psycopg.rows import dict_row
except ImportError:
    psycopg = None
    dict_row = None


H_ALIGN = {1: "left", 2: "center", 3: "right", 4: "fill", 5: "justify", 6: "centerContinuous", 7: "distributed"}
V_ALIGN = {0: "top", 1: "center", 2: "bottom", 3: "justify", 4: "distributed"}
NO_BORDER = Border(
    left=Side(border_style=None),
    right=Side(border_style=None),
    top=Side(border_style=None),
    bottom=Side(border_style=None),
)
ALLOWED_EXTENSIONS = {"xls", "xlsx"}
MAX_UPLOAD_SIZE = 10 * 1024 * 1024
DATABASE_URL = os.environ.get("DATABASE_URL")


def clean(val):
    if isinstance(val, str):
        return val.replace("*", "").replace("?", "").strip()
    return val


def get_extension(filename):
    if not filename or "." not in filename:
        return ""
    return filename.rsplit(".", 1)[-1].lower()


def detect_excel_signature(file_storage):
    pos = file_storage.stream.tell()
    header = file_storage.stream.read(8)
    file_storage.stream.seek(pos)
    if header.startswith(b"PK\x03\x04"):
        return "xlsx"
    if header.startswith(b"\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1"):
        return "xls"
    return ""


def validate_upload(file_storage):
    if not file_storage or file_storage.filename == "":
        return "missing", None

    ext = get_extension(file_storage.filename)
    if ext not in ALLOWED_EXTENSIONS:
        return "unsupported", None

    detected = detect_excel_signature(file_storage)
    if detected != ext:
        return "invalid_excel", None

    file_storage.stream.seek(0, os.SEEK_END)
    size = file_storage.stream.tell()
    file_storage.stream.seek(0)
    if size <= 0:
        return "empty", None
    if size > MAX_UPLOAD_SIZE:
        return "too_large", None

    return None, ext


def idx_to_hex(cmap, idx):
    if idx in (0, 64, 65, 32767, None):
        return None
    rgb = cmap.get(idx)
    if rgb and None not in rgb:
        return "{:02X}{:02X}{:02X}".format(*rgb)
    return None


def widen_known_columns(ws):
    for col, extra in (("G", 4.5), ("S", 3.0), ("AU", 3.0)):
        current = ws.column_dimensions[col].width or 8.43
        ws.column_dimensions[col].width = current + extra


def process_legacy_xls(input_path, output_path):
    wb_in = xlrd.open_workbook(input_path, formatting_info=True)
    cmap = wb_in.colour_map
    wb_out = Workbook()
    wb_out.remove(wb_out.active)

    for s_idx in range(wb_in.nsheets):
        ws_in = wb_in.sheet_by_index(s_idx)
        ws_out = wb_out.create_sheet(title=ws_in.name)
        ws_out.sheet_view.rightToLeft = True
        ws_out.sheet_view.showGridLines = False

        for c in range(ws_in.ncols):
            ltr = get_column_letter(c + 1)
            ci = ws_in.colinfo_map.get(c)
            width = (ci.width / 256.0) if (ci and ci.width) else 1.0
            ws_out.column_dimensions[ltr].width = max(width, 0.5)

        for r in range(ws_in.nrows):
            ri = ws_in.rowinfo_map.get(r)
            orig_pts = (ri.height / 20.0) if (ri and ri.height) else 12.75
            has_content = any(v != "" for v in ws_in.row_values(r))
            if has_content:
                new_h = orig_pts if orig_pts >= 25 else (18 if orig_pts >= 15 else 15)
            else:
                new_h = 2 if orig_pts <= 8 else (orig_pts if orig_pts <= 50 else 12)
            ws_out.row_dimensions[r + 1].height = new_h

        for r in range(ws_in.nrows):
            for c in range(ws_in.ncols):
                val = clean(ws_in.cell_value(r, c))
                xf = wb_in.xf_list[ws_in.cell_xf_index(r, c)]
                fi = wb_in.font_list[xf.font_index]
                cell = ws_out.cell(row=r + 1, column=c + 1, value=val)
                fc = idx_to_hex(cmap, fi.colour_index)
                cell.font = Font(
                    bold=bool(fi.bold),
                    italic=bool(fi.italic),
                    size=fi.height / 20,
                    name=fi.name or "Arial",
                    color=fc or "000000",
                )
                bg = idx_to_hex(cmap, xf.background.pattern_colour_index)
                if bg and xf.background.fill_pattern != 0:
                    cell.fill = PatternFill(fill_type="solid", fgColor=bg)
                cell.alignment = Alignment(
                    horizontal=H_ALIGN.get(xf.alignment.hor_align, "general"),
                    vertical=V_ALIGN.get(xf.alignment.vert_align, "bottom"),
                    wrapText=bool(xf.alignment.text_wrapped),
                    shrinkToFit=bool(xf.alignment.shrink_to_fit),
                    readingOrder=2,
                )
                cell.border = NO_BORDER

        for r1, r2, c1, c2 in ws_in.merged_cells:
            if r2 > r1 or c2 > c1:
                try:
                    ws_out.merge_cells(start_row=r1 + 1, start_column=c1 + 1, end_row=r2, end_column=c2)
                except Exception:
                    pass

    for ws in wb_out.worksheets:
        widen_known_columns(ws)

    wb_out.save(output_path)


def process_xlsx(input_path, output_path):
    wb = load_workbook(input_path)
    for ws in wb.worksheets:
        ws.sheet_view.rightToLeft = True
        ws.sheet_view.showGridLines = False
        for row in ws.iter_rows():
            for cell in row:
                if isinstance(cell.value, str):
                    cell.value = clean(cell.value)
        widen_known_columns(ws)
    wb.save(output_path)


def process_spreadsheet(input_path, output_path, extension):
    if extension == "xls":
        process_legacy_xls(input_path, output_path)
        return
    if extension == "xlsx":
        process_xlsx(input_path, output_path)
        return
    raise ValueError("Unsupported file type")


def run_attendance_cleanup(input_path, output_path, extension, options=None):
    process_spreadsheet(input_path, output_path, extension)


PAYABLE_HOUR_LABELS = {"׳¨׳’׳™׳׳•׳×", "100%", "125%", "150%", "175%", "200%"}


def parse_numeric_rate(value):
    if value in ("", None):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip().replace(",", ".")
    if not text:
        return None
    return float(text)


def parse_hours_value(value):
    if value in ("", None):
        return None
    if isinstance(value, (int, float)):
        numeric = float(value)
        if 0 <= numeric <= 1:
            return numeric * 24
        return numeric
    text = str(value).strip()
    if not text:
        return None
    if ":" in text:
        hours_text, minutes_text = text.split(":", 1)
        return int(hours_text) + (int(minutes_text) / 60.0)
    return float(text.replace(",", "."))


def format_hours(hours_value):
    if hours_value in (None, ""):
        return ""
    total_minutes = int(round(float(hours_value) * 60))
    hours, minutes = divmod(total_minutes, 60)
    return f"{hours:02d}:{minutes:02d}"


def parse_hours_or_zero(value):
    parsed = parse_hours_value(value)
    return 0.0 if parsed is None else parsed


def parse_float_or_none(value):
    text = str(value).strip()
    if not text:
        return None
    return float(text.replace(",", "."))


def safe_sheet_title(title, fallback):
    cleaned = "".join(ch for ch in str(title) if ch not in '[]:*?/\\')[:31].strip()
    return cleaned or fallback


def get_sheet_cell(sheet, row_index, col_index, default=""):
    if row_index >= sheet.nrows or col_index >= sheet.ncols:
        return default
    value = sheet.cell_value(row_index, col_index)
    if isinstance(value, str):
        return value.strip()
    return value


def find_row_label_value(sheet, row_index, label):
    if row_index >= sheet.nrows:
        return ""
    values = [sheet.cell_value(row_index, c) for c in range(sheet.ncols)]
    for idx, value in enumerate(values):
        if str(value).strip() == label:
            for next_idx in range(idx + 1, len(values)):
                candidate = values[next_idx]
                if candidate not in ("", None):
                    return candidate
            return ""
    return ""


def find_row_label_value_with_offsets(sheet, row_index, label, offsets):
    if row_index >= sheet.nrows:
        return ""
    values = [sheet.cell_value(row_index, c) for c in range(sheet.ncols)]
    for idx, value in enumerate(values):
        if str(value).strip() == label:
            for offset in offsets:
                candidate_index = idx + offset
                if 0 <= candidate_index < len(values):
                    candidate = values[candidate_index]
                    if candidate not in ("", None):
                        return candidate
            return ""
    return ""


def extract_payable_hours(summary_sheet):
    totals = {}
    for row_index in range(summary_sheet.nrows):
        label = str(get_sheet_cell(summary_sheet, row_index, 9, "")).strip()
        if label in PAYABLE_HOUR_LABELS:
            totals[label] = parse_hours_value(get_sheet_cell(summary_sheet, row_index, 13, ""))
    available = [value for value in totals.values() if value is not None]
    if not available:
        return None, totals
    return sum(available), totals


def extract_flamingo_worker_pair(detail_sheet, summary_sheet):
    worker_name = str(find_row_label_value_with_offsets(detail_sheet, 5, "׳©׳ ׳׳×׳¦׳•׳’׳”", [2, 1])).strip() or detail_sheet.name
    department = str(find_row_label_value(detail_sheet, 5, "׳׳—׳׳§׳”")).strip()
    rate_raw = find_row_label_value(detail_sheet, 5, "׳”׳¢׳¨׳•׳×")
    worker_number = find_row_label_value(detail_sheet, 5, "׳׳¡׳₪׳¨ ׳‘׳©׳›׳¨")
    id_number = find_row_label_value(detail_sheet, 5, "׳×׳¢׳•׳“׳× ׳–׳”׳•׳×")
    start_date = find_row_label_value(detail_sheet, 5, "׳×׳—׳™׳׳× ׳¢׳‘׳•׳“׳”")
    department = str(find_row_label_value_with_offsets(detail_sheet, 5, "׳׳—׳׳§׳”", [3, 2, 1])).strip()
    rate_raw = find_row_label_value_with_offsets(detail_sheet, 5, "׳”׳¢׳¨׳•׳×", [4, 3, 2, 1])
    worker_number = find_row_label_value_with_offsets(detail_sheet, 5, "׳׳¡׳₪׳¨ ׳‘׳©׳›׳¨", [5, 4, 3, 2, 1])
    id_number = find_row_label_value_with_offsets(detail_sheet, 5, "׳×׳¢׳•׳“׳× ׳–׳”׳•׳×", [2, 1])
    start_date = find_row_label_value_with_offsets(detail_sheet, 5, "׳×׳—׳™׳׳× ׳¢׳‘׳•׳“׳”", [4, 3, 2, 1])
    notes = []
    status = "OK"

    try:
        hourly_rate = parse_numeric_rate(rate_raw)
    except ValueError:
        hourly_rate = None
        status = "Invalid hourly rate"
        notes.append(f"Hourly rate value is invalid: {rate_raw}")

    if hourly_rate is None and status == "OK":
        status = "Missing hourly rate"
        notes.append("Set hourly rate in the ׳”׳¢׳¨׳•׳× field and export the report again.")

    payable_hours = None
    payable_breakdown = {}
    summary_name = ""
    if summary_sheet is None:
        status = "Could not match summary sheet"
        notes.append("Expected the summary sheet immediately after the worker detail sheet.")
    else:
        summary_name = summary_sheet.name
        payable_hours, payable_breakdown = extract_payable_hours(summary_sheet)
        if payable_hours is None and status == "OK":
            status = "Missing payable hours"
            notes.append("No payable hour values were found in the summary sheet.")

    calculated_salary = None
    if status == "OK":
        calculated_salary = round(payable_hours * hourly_rate, 2)

    return {
        "worker_name": worker_name,
        "department": department,
        "worker_number": worker_number,
        "id_number": id_number,
        "start_date": start_date,
        "detail_sheet": detail_sheet.name,
        "summary_sheet": summary_name,
        "hourly_rate": hourly_rate,
        "hourly_rate_raw": rate_raw,
        "payable_hours": payable_hours,
        "payable_breakdown": payable_breakdown,
        "salary": calculated_salary,
        "status": status,
        "notes": " | ".join(notes),
    }


def write_flamingo_summary_sheet(ws, worker_rows):
    ws.sheet_view.rightToLeft = True
    ws.sheet_view.showGridLines = False
    ws.title = safe_sheet_title("Payroll Summary", "Payroll Summary")
    ws.freeze_panes = "A12"

    successful_rows = [row for row in worker_rows if row["status"] == "OK"]
    total_workers = len(worker_rows)
    unresolved_workers = len([row for row in worker_rows if row["status"] != "OK"])
    total_hours = sum(row["payable_hours"] or 0 for row in successful_rows)
    total_salary = sum(row["salary"] or 0 for row in successful_rows)
    average_rate = (sum(row["hourly_rate"] or 0 for row in successful_rows) / len(successful_rows)) if successful_rows else 0

    metrics = [
        ("Total workers", total_workers, "DBEAFE"),
        ("Calculated successfully", len(successful_rows), "DCFCE7"),
        ("Requires attention", unresolved_workers, "FEE2E2"),
        ("Total payable hours", format_hours(total_hours), "FEF3C7"),
        ("Total payroll", round(total_salary, 2), "E9D5FF"),
        ("Average hourly rate", round(average_rate, 2), "FCE7F3"),
    ]

    ws.merge_cells("A1:D1")
    ws["A1"] = "Flamingo Payroll Summary"
    ws["A1"].font = Font(bold=True, size=18, color="0F172A")
    ws["A1"].fill = PatternFill(fill_type="solid", fgColor="BFDBFE")
    ws["A1"].alignment = Alignment(horizontal="center")
    ws["A2"] = "Automatic payroll calculation from monthly attendance export"
    ws["A2"].font = Font(italic=True, size=11, color="475569")

    for index, (label, value, fill_color) in enumerate(metrics):
        start_col = 1 + (index % 3) * 4
        row = 4 + (index // 3) * 3
        ws.merge_cells(start_row=row, start_column=start_col, end_row=row, end_column=start_col + 1)
        ws.merge_cells(start_row=row + 1, start_column=start_col, end_row=row + 1, end_column=start_col + 1)
        label_cell = ws.cell(row=row, column=start_col, value=label)
        value_cell = ws.cell(row=row + 1, column=start_col, value=value)
        label_cell.font = Font(bold=True, color="334155")
        value_cell.font = Font(bold=True, size=14, color="0F172A")
        label_cell.fill = PatternFill(fill_type="solid", fgColor=fill_color)
        value_cell.fill = PatternFill(fill_type="solid", fgColor=fill_color)
        label_cell.alignment = Alignment(horizontal="center")
        value_cell.alignment = Alignment(horizontal="center")

    header_row = 11
    headers = [
        "Worker Name",
        "Worker Number",
        "ID Number",
        "Department",
        "Start Date",
        "Detail Sheet",
        "Summary Sheet",
        "Hourly Rate",
        "Payable Hours",
        "Calculated Salary",
        "Status",
        "Notes",
    ]

    for col_index, header in enumerate(headers, start=1):
        cell = ws.cell(row=header_row, column=col_index, value=header)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill(fill_type="solid", fgColor="1E3A8A")
        cell.alignment = Alignment(horizontal="center")
    ws.auto_filter.ref = f"A{header_row}:{get_column_letter(len(headers))}{header_row}"

    for row_index, worker in enumerate(worker_rows, start=header_row + 1):
        values = [
            worker["worker_name"],
            worker["worker_number"],
            worker["id_number"],
            worker["department"],
            worker["start_date"],
            worker["detail_sheet"],
            worker["summary_sheet"],
            worker["hourly_rate"],
            format_hours(worker["payable_hours"]),
            worker["salary"],
            worker["status"],
            worker["notes"],
        ]
        for col_index, value in enumerate(values, start=1):
            ws.cell(row=row_index, column=col_index, value=value)
        fill_color = "ECFDF5" if worker["status"] == "OK" else "FEE2E2"
        for col_index in range(1, len(headers) + 1):
            ws.cell(row=row_index, column=col_index).fill = PatternFill(fill_type="solid", fgColor=fill_color)
        ws.cell(row=row_index, column=8).number_format = '0.00'
        ws.cell(row=row_index, column=10).number_format = '#,##0.00'

    widths = [22, 16, 18, 18, 14, 18, 18, 14, 14, 16, 22, 42]
    for col_index, width in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(col_index)].width = width


def write_flamingo_attention_sheet(ws, worker_rows):
    ws.sheet_view.rightToLeft = True
    ws.sheet_view.showGridLines = False
    ws.title = safe_sheet_title("Requires Attention", "Requires Attention")
    ws.freeze_panes = "A2"

    headers = ["Worker Name", "Worker Number", "ID Number", "Issue", "Hourly Rate", "Payable Hours", "Recommended Action"]
    for col_index, header in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=col_index, value=header)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill(fill_type="solid", fgColor="B91C1C")
        cell.alignment = Alignment(horizontal="center")
    ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}1"

    issues = [row for row in worker_rows if row["status"] != "OK"]
    for row_index, worker in enumerate(issues, start=2):
        if worker["status"] in {"Missing hourly rate", "Invalid hourly rate"}:
            action = "Update the hourly rate in ׳”׳¢׳¨׳•׳× and export the report again."
        elif worker["status"] == "Could not match summary sheet":
            action = "Verify the report structure and confirm that each detail sheet has a following summary sheet."
        else:
            action = "Verify payable hour values in the summary sheet."

        values = [
            worker["worker_name"],
            worker["worker_number"],
            worker["id_number"],
            worker["status"],
            worker["hourly_rate_raw"],
            format_hours(worker["payable_hours"]),
            action,
        ]
        for col_index, value in enumerate(values, start=1):
            ws.cell(row=row_index, column=col_index, value=value)
            ws.cell(row=row_index, column=col_index).fill = PatternFill(fill_type="solid", fgColor="FEF2F2")

    widths = [22, 16, 18, 24, 14, 14, 60]
    for col_index, width in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(col_index)].width = width


def write_flamingo_department_sheet(ws, worker_rows):
    ws.sheet_view.rightToLeft = True
    ws.sheet_view.showGridLines = False
    ws.title = safe_sheet_title("Department Summary", "Department Summary")
    ws.freeze_panes = "A2"

    headers = ["Department", "Workers", "Calculated Workers", "Payable Hours", "Payroll"]
    for col_index, header in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=col_index, value=header)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill(fill_type="solid", fgColor="0F766E")
        cell.alignment = Alignment(horizontal="center")
    ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}1"

    department_totals = defaultdict(lambda: {"workers": 0, "calculated": 0, "hours": 0.0, "salary": 0.0})
    for worker in worker_rows:
        department = worker["department"] or "Unassigned"
        bucket = department_totals[department]
        bucket["workers"] += 1
        if worker["status"] == "OK":
            bucket["calculated"] += 1
            bucket["hours"] += worker["payable_hours"] or 0
            bucket["salary"] += worker["salary"] or 0

    for row_index, (department, totals) in enumerate(sorted(department_totals.items()), start=2):
        values = [
            department,
            totals["workers"],
            totals["calculated"],
            format_hours(totals["hours"]),
            round(totals["salary"], 2),
        ]
        for col_index, value in enumerate(values, start=1):
            ws.cell(row=row_index, column=col_index, value=value)
            if row_index % 2 == 0:
                ws.cell(row=row_index, column=col_index).fill = PatternFill(fill_type="solid", fgColor="F0FDFA")

    widths = [24, 12, 18, 16, 16]
    for col_index, width in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(col_index)].width = width


def write_flamingo_top_earners_sheet(ws, worker_rows):
    ws.sheet_view.rightToLeft = True
    ws.sheet_view.showGridLines = False
    ws.title = safe_sheet_title("Top Earners", "Top Earners")
    ws.freeze_panes = "A2"

    headers = ["Rank", "Worker Name", "ID Number", "Department", "Hourly Rate", "Payable Hours", "Calculated Salary"]
    for col_index, header in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=col_index, value=header)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill(fill_type="solid", fgColor="7C3AED")
        cell.alignment = Alignment(horizontal="center")
    ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}1"

    ranked_workers = sorted(
        [row for row in worker_rows if row["status"] == "OK" and row["salary"] is not None],
        key=lambda row: row["salary"],
        reverse=True,
    )[:10]

    for row_index, worker in enumerate(ranked_workers, start=2):
        values = [
            row_index - 1,
            worker["worker_name"],
            worker["id_number"],
            worker["department"],
            worker["hourly_rate"],
            format_hours(worker["payable_hours"]),
            worker["salary"],
        ]
        for col_index, value in enumerate(values, start=1):
            ws.cell(row=row_index, column=col_index, value=value)
            if row_index % 2 == 0:
                ws.cell(row=row_index, column=col_index).fill = PatternFill(fill_type="solid", fgColor="F5F3FF")
        ws.cell(row=row_index, column=5).number_format = '0.00'
        ws.cell(row=row_index, column=7).number_format = '#,##0.00'

    widths = [10, 22, 18, 20, 14, 14, 16]
    for col_index, width in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(col_index)].width = width


def load_org_structure_csv(csv_path):
    records = {}
    unmatched = []
    with open(csv_path, "r", encoding="utf-8-sig", newline="") as handle:
        reader = csv.DictReader(handle)
        for row in reader:
            employee_number = (row.get("׳©׳›׳¨") or "").strip()
            employee_id = (row.get("׳×.׳–") or "").strip()
            entry = {
                "employee_number": employee_number,
                "id_number": employee_id,
                "employee_name": (row.get("׳©׳ ׳¢׳•׳‘׳“") or "").strip(),
                "direct_manager": (row.get("׳׳ ׳”׳ ׳™׳©׳™׳¨") or "").strip(),
                "department": (row.get("׳׳—׳׳§׳”") or "").strip(),
                "agreement_name": (row.get("׳©׳ ׳”׳¡׳›׳") or "").strip(),
                "agreement_number": (row.get("׳׳¡' ׳”׳¡׳›׳") or "").strip(),
            }
            if employee_number:
                records[("number", employee_number)] = entry
            if employee_id:
                records[("id", employee_id)] = entry
            if not employee_number and not employee_id:
                unmatched.append(entry)
    return records, unmatched


def parse_matan_missing_report(input_path):
    wb = xlrd.open_workbook(input_path)
    ws = wb.sheet_by_index(0)
    header_row = 4
    headers = [str(ws.cell_value(header_row, c)).strip() for c in range(ws.ncols)]
    header_index = {header: idx for idx, header in enumerate(headers) if header}
    employee_number_col = header_index.get("מספר עובד", -1)
    month_col = header_index.get("חודש", -1)
    employee_name_col = header_index.get("שם עובד", -1)
    standard_hours_col = header_index.get("ש.תקן", -1)
    missing_hours_col = header_index.get("חוסר", -1)
    attendance_hours_col = header_index.get("ש.נוכחות", -1)
    vacation_hours_col = header_index.get("חופשה", -1)
    sick_hours_col = header_index.get("מחלה", -1)
    reserve_hours_col = header_index.get("מילואים", -1)
    pregnancy_hours_col = header_index.get("שעות הריון", -1)
    special_child_hours_col = header_index.get("שעות ילד מיחד", -1)
    absence_hours_col = header_index.get("היעדרות", -1)
    rows = []
    for row_index in range(header_row + 1, ws.nrows):
        employee_number = str(get_sheet_cell(ws, row_index, employee_number_col, "")).strip()
        employee_name = str(get_sheet_cell(ws, row_index, employee_name_col, "")).strip()
        if not employee_number and not employee_name:
            continue
        row = {
            "employee_number": employee_number,
            "month": str(get_sheet_cell(ws, row_index, month_col, "")).strip(),
            "employee_name": employee_name,
            "standard_hours": parse_hours_or_zero(get_sheet_cell(ws, row_index, standard_hours_col, "")),
            "missing_hours": parse_hours_or_zero(get_sheet_cell(ws, row_index, missing_hours_col, "")),
            "attendance_hours": parse_hours_or_zero(get_sheet_cell(ws, row_index, attendance_hours_col, "")),
            "vacation_hours": parse_hours_or_zero(get_sheet_cell(ws, row_index, vacation_hours_col, "")),
            "sick_hours": parse_hours_or_zero(get_sheet_cell(ws, row_index, sick_hours_col, "")),
            "reserve_hours": parse_hours_or_zero(get_sheet_cell(ws, row_index, reserve_hours_col, "")),
            "pregnancy_hours": parse_hours_or_zero(get_sheet_cell(ws, row_index, pregnancy_hours_col, "")),
            "special_child_hours": parse_hours_or_zero(get_sheet_cell(ws, row_index, special_child_hours_col, "")),
            "absence_hours": parse_hours_or_zero(get_sheet_cell(ws, row_index, absence_hours_col, "")),
        }
        rows.append(row)
    return rows


def apply_matan_missing_filters(rows, options):
    min_missing = parse_float_or_none(options.get("min_missing_hours", ""))
    max_missing = parse_float_or_none(options.get("max_missing_hours", ""))

    filtered = []
    for row in rows:
        missing_hours = row["missing_hours"] or 0.0
        if min_missing is not None and missing_hours < min_missing:
            continue
        if max_missing is not None and missing_hours > max_missing:
            continue
        filtered.append(row)
    return filtered


def write_matan_missing_summary(ws, filtered_rows, filters_used):
    ws.title = safe_sheet_title("Missing Hours Summary", "Missing Summary")
    ws.sheet_view.rightToLeft = True
    ws.sheet_view.showGridLines = False
    ws.freeze_panes = "A10"

    total_missing = sum(row["missing_hours"] for row in filtered_rows)
    total_attendance = sum(row["attendance_hours"] for row in filtered_rows)
    total_standard = sum(row["standard_hours"] for row in filtered_rows)
    avg_missing = (total_missing / len(filtered_rows)) if filtered_rows else 0.0
    over_4 = sum(1 for row in filtered_rows if (row["missing_hours"] or 0.0) > 4.0)
    over_8 = sum(1 for row in filtered_rows if (row["missing_hours"] or 0.0) > 8.0)

    ws["A1"] = "Matan Missing Hours Summary"
    ws["A1"].font = Font(bold=True, size=18)
    ws["A1"].fill = PatternFill(fill_type="solid", fgColor="BFDBFE")

    metrics = [
        ("Employees in result", len(filtered_rows)),
        ("Total missing hours", format_hours(total_missing)),
        ("Average missing hours", format_hours(avg_missing)),
        ("Total attendance hours", format_hours(total_attendance)),
        ("Total standard hours", format_hours(total_standard)),
        ("Employees above 4 missing hours", over_4),
        ("Employees above 8 missing hours", over_8),
    ]
    for idx, (label, value) in enumerate(metrics, start=3):
        ws.cell(row=idx, column=1, value=label).font = Font(bold=True)
        ws.cell(row=idx, column=2, value=value)

    ws["D3"] = "Filters used"
    ws["D3"].font = Font(bold=True)
    for idx, (label, value) in enumerate(filters_used.items(), start=4):
        ws.cell(row=idx, column=4, value=label).font = Font(bold=True)
        ws.cell(row=idx, column=5, value=value or "All")

    header_row = 10
    headers = ["Employee Number", "Employee Name", "Month", "Missing Hours", "Attendance Hours", "Standard Hours", "Vacation", "Sick", "Reserve", "Absence"]
    for col, header in enumerate(headers, start=1):
        cell = ws.cell(row=header_row, column=col, value=header)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill(fill_type="solid", fgColor="1E3A8A")

    for row_idx, row in enumerate(filtered_rows, start=header_row + 1):
        values = [
            row["employee_number"],
            row["employee_name"],
            row["month"],
            format_hours(row["missing_hours"]),
            format_hours(row["attendance_hours"]),
            format_hours(row["standard_hours"]),
            format_hours(row["vacation_hours"]),
            format_hours(row["sick_hours"]),
            format_hours(row["reserve_hours"]),
            format_hours(row["absence_hours"]),
        ]
        for col, value in enumerate(values, start=1):
            ws.cell(row=row_idx, column=col, value=value)
            if row_idx % 2 == 0:
                ws.cell(row=row_idx, column=col).fill = PatternFill(fill_type="solid", fgColor="F8FAFC")
    widths = [16, 24, 14, 14, 14, 14, 12, 12, 12, 12]
    for col, width in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(col)].width = width


def write_matan_missing_filtered(ws, filtered_rows):
    ws.title = safe_sheet_title("Filtered Employees", "Filtered Employees")
    ws.sheet_view.rightToLeft = True
    ws.sheet_view.showGridLines = False
    ws.freeze_panes = "A2"
    headers = [
        "Employee Number", "Employee Name", "Month", "Standard Hours", "Missing Hours", "Attendance Hours",
        "Vacation", "Sick", "Reserve", "Pregnancy", "Special Child", "Absence"
    ]
    for col, header in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=col, value=header)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill(fill_type="solid", fgColor="0F766E")
    for row_idx, row in enumerate(filtered_rows, start=2):
        values = [
            row["employee_number"], row["employee_name"], row["month"], format_hours(row["standard_hours"]), format_hours(row["missing_hours"]),
            format_hours(row["attendance_hours"]), format_hours(row["vacation_hours"]), format_hours(row["sick_hours"]),
            format_hours(row["reserve_hours"]), format_hours(row["pregnancy_hours"]), format_hours(row["special_child_hours"]),
            format_hours(row["absence_hours"]),
        ]
        for col, value in enumerate(values, start=1):
            ws.cell(row=row_idx, column=col, value=value)
    widths = [16, 24, 14, 14, 14, 14, 12, 12, 12, 12, 14, 12]
    for col, width in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(col)].width = width


def write_matan_missing_unmatched(ws, unmatched_rows):
    ws.title = safe_sheet_title("Unmatched Employees", "Unmatched Employees")
    ws.sheet_view.rightToLeft = True
    ws.sheet_view.showGridLines = False
    headers = ["Employee Number", "Employee Name", "Note"]
    for col, header in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=col, value=header)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill(fill_type="solid", fgColor="B91C1C")
    for row_idx, row in enumerate(unmatched_rows, start=2):
        ws.cell(row=row_idx, column=1, value=row["employee_number"])
        ws.cell(row=row_idx, column=2, value=row["employee_name"])
        ws.cell(row=row_idx, column=3, value="No matching record was found in the organization CSV")
    widths = [16, 24, 42]
    for col, width in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(col)].width = width


def run_matan_missing_filter(input_path, output_path, extension, options=None):
    if extension != "xls":
        raise ValueError("Matan missing-hours tool currently supports XLS export only")
    options = options or {}
    rows = parse_matan_missing_report(input_path)
    filtered_rows = apply_matan_missing_filters(rows, options)
    wb = Workbook()
    write_matan_missing_summary(
        wb.active,
        filtered_rows,
        {
            "Min missing hours": options.get("min_missing_hours", ""),
            "Max missing hours": options.get("max_missing_hours", ""),
        },
    )
    write_matan_missing_filtered(wb.create_sheet(), filtered_rows)
    wb.save(output_path)


def run_flamingo_payroll(input_path, output_path, extension, options=None):
    if extension != "xls":
        raise ValueError("Flamingo payroll currently supports original XLS exports only")

    workbook = xlrd.open_workbook(input_path)
    worker_rows = []
    for sheet_index in range(0, workbook.nsheets, 2):
        detail_sheet = workbook.sheet_by_index(sheet_index)
        summary_sheet = workbook.sheet_by_index(sheet_index + 1) if sheet_index + 1 < workbook.nsheets else None
        worker_rows.append(extract_flamingo_worker_pair(detail_sheet, summary_sheet))

    output_wb = Workbook()
    summary_ws = output_wb.active
    write_flamingo_summary_sheet(summary_ws, worker_rows)
    write_flamingo_attention_sheet(output_wb.create_sheet(), worker_rows)
    write_flamingo_department_sheet(output_wb.create_sheet(), worker_rows)
    write_flamingo_top_earners_sheet(output_wb.create_sheet(), worker_rows)
    output_wb.save(output_path)


app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "zman_emet_secret_2024")
app.config["MAX_CONTENT_LENGTH"] = MAX_UPLOAD_SIZE

BASE_DIR = Path(__file__).resolve().parent
DB = str(BASE_DIR / "platform.db")
UPLOAD_FOLDER = BASE_DIR / "uploads"
OUTPUT_FOLDER = BASE_DIR / "outputs"
UPLOAD_FOLDER.mkdir(exist_ok=True)
OUTPUT_FOLDER.mkdir(exist_ok=True)


class DatabaseConnection:
    def __init__(self):
        self.conn = None
        self.is_postgres = bool(DATABASE_URL)

    def __enter__(self):
        if self.is_postgres:
            if psycopg is None:
                raise RuntimeError("PostgreSQL driver is not installed")
            self.conn = psycopg.connect(DATABASE_URL, row_factory=dict_row)
        else:
            self.conn = sqlite3.connect(DB)
            self.conn.row_factory = sqlite3.Row
        return self

    def __exit__(self, exc_type, exc, tb):
        if self.conn is None:
            return False
        try:
            if exc_type is not None:
                self.conn.rollback()
        finally:
            self.conn.close()
        return False

    def execute(self, sql, params=()):
        if self.conn is None:
            raise RuntimeError("Database connection is not open")
        if self.is_postgres:
            sql = sql.replace("?", "%s").replace("AUTOINCREMENT", "").replace("INTEGER PRIMARY KEY", "SERIAL PRIMARY KEY")
            sql = sql.replace("INSERT OR IGNORE INTO permissions(user_id,script_id) VALUES (%s,%s)", "INSERT INTO permissions(user_id,script_id) VALUES (%s,%s) ON CONFLICT DO NOTHING")
        return self.conn.execute(sql, params)

    def commit(self):
        if self.conn is not None:
            self.conn.commit()


def get_db():
    return DatabaseConnection()


def is_integrity_error(exc):
    if isinstance(exc, sqlite3.IntegrityError):
        return True
    if psycopg is not None and isinstance(exc, psycopg.IntegrityError):
        return True
    return False

SCRIPTS = {
    "nikuy": {
        "id": "nikuy",
        "name": "Attendance Cleanup",
        "desc": "Remove * and ? from monthly attendance report",
        "accept": ".xls,.xlsx",
        "icon": "🧹",
    }
}

SCRIPT_REGISTRY = {
    "nikuy": {
        **SCRIPTS["nikuy"],
        "processor": run_attendance_cleanup,
        "output_suffix": "cleaned",
        "success_title": "File is ready",
        "success_action": "Download cleaned file",
        "retry_action": "Process another file",
        "submit_label": "Run",
        "back_label": "Back to tools",
        "empty_error": "No file selected",
        "unsupported_error": "Unsupported file type",
        "invalid_error": "The uploaded file is not a valid Excel file",
        "empty_file_error": "The uploaded file is empty",
        "too_large_error": "The uploaded file is too large",
        "processing_error": "The uploaded file could not be processed",
        "processing_title": "File is being processed",
        "processing_note": "Preparing the cleaned report may take a few minutes. Please keep this page open.",
        "file_picker_label": "Choose file",
    }
}

SCRIPTS["flamingo_payroll"] = {
    "id": "flamingo_payroll",
    "name": "Flamingo Payroll",
    "desc": "Payroll summary based on payable hours and hourly rate",
    "accept": ".xls",
    "icon": "$",
}

SCRIPT_REGISTRY["flamingo_payroll"] = {
    **SCRIPTS["flamingo_payroll"],
    "processor": run_flamingo_payroll,
    "output_suffix": "flamingo_payroll",
    "success_title": "Payroll file is ready",
    "success_action": "Download payroll summary",
    "retry_action": "Process another payroll file",
    "submit_label": "Create payroll summary",
    "back_label": "Back to tools",
    "empty_error": "No file selected",
    "unsupported_error": "Please upload the original Flamingo XLS export",
    "invalid_error": "The uploaded file is not a valid Excel file",
    "empty_file_error": "The uploaded file is empty",
    "too_large_error": "The uploaded file is too large",
    "processing_error": "Could not generate a payroll summary from this file",
    "processing_title": "Payroll summary is being prepared",
    "processing_note": "The system is calculating payable hours and salary for all employees. This may take a few minutes.",
    "file_picker_label": "Choose Flamingo file",
}

SCRIPTS["matan_missing"] = {
    "id": "matan_missing",
    "name": "Matan Missing Hours",
    "desc": "Filter employees by missing-hours range",
    "accept": ".xls",
    "icon": "ג±",
}

SCRIPT_REGISTRY["matan_missing"] = {
    **SCRIPTS["matan_missing"],
    "processor": run_matan_missing_filter,
    "output_suffix": "matan_missing",
    "success_title": "Missing-hours report is ready",
    "success_action": "Download report",
    "retry_action": "Process another file",
    "submit_label": "Create report",
    "back_label": "Back to tools",
    "empty_error": "No file selected",
    "unsupported_error": "Please upload the original XLS missing-hours report",
    "invalid_error": "The uploaded file is not a valid Excel file",
    "empty_file_error": "The uploaded file is empty",
    "too_large_error": "The uploaded file is too large",
    "processing_error": "Could not generate the missing-hours report from this file",
    "processing_title": "Report is being prepared",
    "processing_note": "The system is filtering the missing-hours report. This may take a few minutes.",
    "file_picker_label": "Choose missing-hours report",
    "filter_fields": [
        {"name": "min_missing_hours", "label": "Minimum missing hours", "placeholder": "For example 4"},
        {"name": "max_missing_hours", "label": "Maximum missing hours", "placeholder": "For example 8"}
    ],
}

SCRIPTS = SCRIPT_REGISTRY


def get_script(script_id):
    return SCRIPT_REGISTRY.get(script_id)


def build_output_filename(script, uid):
    suffix = script.get("output_suffix", "output")
    return f"{uid}_{suffix}.xlsx"


def execute_script(script, input_path, output_path, extension, options=None):
    processor = script.get("processor")
    if processor is None:
        raise ValueError("Script processor is not configured")
    processor(input_path, output_path, extension, options)


CSS = """
* { box-sizing: border-box; margin: 0; padding: 0; }
body { font-family: 'Segoe UI', Arial, sans-serif; background: #f0f4ff; min-height: 100vh; direction: rtl; }
.topbar { background: #1e3a8a; color: white; padding: 0 2rem; height: 58px; display: flex; align-items: center; justify-content: space-between; }
.topbar h1 { font-size: 17px; font-weight: 700; }
.topbar a { color: #93c5fd; font-size: 13px; text-decoration: none; }
.wrap { max-width: 900px; margin: 2rem auto; padding: 0 1rem; }
.login-wrap { max-width: 400px; margin: 5rem auto; padding: 0 1rem; }
.card { background: white; border-radius: 16px; box-shadow: 0 4px 24px rgba(37,99,235,.1); padding: 2rem; margin-bottom: 1.5rem; }
.card h2 { font-size: 16px; font-weight: 700; color: #1e3a8a; margin-bottom: 1rem; padding-bottom: .75rem; border-bottom: 1.5px solid #e0e7ff; }
label.field-label { font-size: 13px; font-weight: 600; color: #374151; margin-bottom: 5px; display: block; }
input[type=text], input[type=password] { padding: 9px 12px; border: 1.5px solid #e2e8f0; border-radius: 8px; font-size: 13px; font-family: inherit; outline: none; width: 100%; margin-bottom: .75rem; }
input:focus { border-color: #2563eb; }
.btn { padding: 10px 20px; border: none; border-radius: 8px; font-size: 13px; font-weight: 600; cursor: pointer; font-family: inherit; }
.btn-blue { background: #2563eb; color: white; }
.btn-blue:hover { background: #1d4ed8; }
.btn-red { background: #fee2e2; color: #dc2626; }
.btn-gray { background: #f1f5f9; color: #475569; }
.flash { background: #f0fdf4; border: 1px solid #86efac; color: #15803d; border-radius: 8px; padding: 10px 14px; font-size: 13px; margin-bottom: 1rem; }
.flash-err { background: #fef2f2; border: 1px solid #fecaca; color: #dc2626; border-radius: 8px; padding: 10px 14px; font-size: 13px; margin-bottom: 1rem; }
table { width: 100%; border-collapse: collapse; font-size: 13px; }
th { text-align: right; padding: 10px 12px; background: #f8fafc; color: #64748b; font-weight: 600; border-bottom: 1.5px solid #e2e8f0; }
td { padding: 12px; border-bottom: 1px solid #f1f5f9; vertical-align: middle; }
.badge { display: inline-block; padding: 3px 10px; border-radius: 99px; font-size: 11px; font-weight: 600; background: #f1f5f9; color: #64748b; }
.form-row { display: flex; gap: 10px; flex-wrap: wrap; align-items: flex-end; }
.form-group { flex: 1; min-width: 130px; }
.drop-zone { border: 2px dashed #c7d7f5; border-radius: 14px; padding: 2rem; text-align: center; cursor: pointer; background: #fafcff; margin-bottom: 1rem; }
.drop-zone:hover { border-color: #2563eb; background: #eff6ff; }
.success-box { padding: 1.25rem; background: #f0fdf4; border: 1.5px solid #86efac; border-radius: 13px; text-align: center; margin-top: 1rem; }
.dl-btn { display: inline-block; padding: 11px 28px; background: #16a34a; color: white; border-radius: 9px; font-size: 14px; font-weight: 700; text-decoration: none; }
.processing-box { display: none; margin-top: 1rem; padding: 1.1rem 1.25rem; background: #eff6ff; border: 1.5px solid #bfdbfe; border-radius: 13px; }
.processing-box.show { display: block; }
.progress-track { width: 100%; height: 12px; border-radius: 999px; background: #dbeafe; overflow: hidden; margin: .9rem 0 .7rem; }
.progress-bar { width: 45%; height: 100%; border-radius: 999px; background: linear-gradient(90deg, #2563eb 0%, #60a5fa 100%); animation: loadingSlide 1.6s ease-in-out infinite; }
.processing-note { font-size: 13px; color: #1d4ed8; font-weight: 600; }
.processing-subnote { font-size: 12px; color: #64748b; line-height: 1.6; }
@keyframes loadingSlide {
  0% { transform: translateX(0); }
  50% { transform: translateX(120%); }
  100% { transform: translateX(0); }
}
.modal-bg { display: none; position: fixed; inset: 0; background: rgba(0,0,0,.4); z-index: 100; align-items: center; justify-content: center; }
.modal-box { background: white; border-radius: 16px; padding: 1.75rem; width: 320px; }
"""


def init_db():
    with get_db() as db:
        db.execute(
            """CREATE TABLE IF NOT EXISTS users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            username TEXT UNIQUE NOT NULL,
            password TEXT NOT NULL,
            full_name TEXT,
            is_admin INTEGER DEFAULT 0,
            active INTEGER DEFAULT 1)"""
        )
        db.execute(
            """CREATE TABLE IF NOT EXISTS permissions (
            user_id INTEGER, script_id TEXT,
            PRIMARY KEY (user_id, script_id))"""
        )
        if not db.execute("SELECT id FROM users WHERE username='admin'").fetchone():
            db.execute(
                "INSERT INTO users(username,password,full_name,is_admin) VALUES (?,?,?,1)",
                ("admin", generate_password_hash("admin123"), "׳׳ ׳”׳ ׳׳¢׳¨׳›׳×"),
            )
        db.commit()


init_db()


def add_flash(msg):
    session.setdefault("msgs", []).append(msg)


def pop_flashes():
    msgs = session.pop("msgs", [])
    return "".join('<div class="flash">' + m + "</div>" for m in msgs)


def generate_temp_password(length=10):
    alphabet = string.ascii_letters + string.digits
    return "".join(secrets.choice(alphabet) for _ in range(length))


def render(title, body, nav=True):
    topbar = ""
    if nav:
        name = session.get("name", "")
        topbar = (
            '<div class="topbar">'
            "<h1>&#9201; Scriptly</h1>"
            '<div style="display:flex;gap:16px;align-items:center">'
            '<span style="font-size:13px;color:#93c5fd">Hello, ' + name + "</span>"
            '<a href="/logout">Logout</a>'
            "</div></div>"
        )
    wrap_cls = "wrap" if nav else "login-wrap"
    return (
        '<!DOCTYPE html><html dir="rtl" lang="he">'
        "<head><meta charset=\"UTF-8\">"
        '<meta name="viewport" content="width=device-width,initial-scale=1">'
        "<title>" + title + " | Scriptly</title>"
        "<style>" + CSS + "</style></head>"
        "<body>" + topbar + '<div class="' + wrap_cls + '">' + pop_flashes() + body + "</div></body></html>"
    )


def login_required(f):
    from functools import wraps

    @wraps(f)
    def decorated(*args, **kwargs):
        if "user_id" not in session:
            return redirect("/")
        return f(*args, **kwargs)

    return decorated


def admin_required(f):
    from functools import wraps

    @wraps(f)
    def decorated(*args, **kwargs):
        if not session.get("is_admin"):
            return redirect("/dashboard")
        return f(*args, **kwargs)

    return decorated


@app.route("/", methods=["GET", "POST"])
def login():
    if "user_id" in session:
        return redirect("/admin" if session.get("is_admin") else "/dashboard")

    error = ""
    if request.method == "POST":
        username = request.form["username"].strip()
        password = request.form["password"]
        with get_db() as db:
            user = db.execute("SELECT * FROM users WHERE username=? AND active=1", (username,)).fetchone()
        if user and check_password_hash(user["password"], password):
            session.update(
                {
                    "user_id": user["id"],
                    "username": user["username"],
                    "name": user["full_name"],
                    "is_admin": bool(user["is_admin"]),
                }
            )
            return redirect("/admin" if user["is_admin"] else "/dashboard")
        error = '<div class="flash-err">Wrong username or password</div>'

    body = (
        '<div class="card" style="padding:2rem">'
        '<div style="text-align:center;margin-bottom:1.5rem">'
        '<div style="font-size:40px">&#9201;</div>'
        '<h1 style="font-size:20px;font-weight:700;color:#1e3a8a;margin-top:8px">Scriptly</h1>'
        '<p style="font-size:12px;color:#888;margin-top:3px">Attendance and payroll tools</p>'
        "</div>"
        + error
        + '<form method="POST">'
        '<label class="field-label">Username</label>'
        '<input type="text" name="username" required autofocus>'
        '<label class="field-label">Password</label>'
        '<input type="password" name="password" required>'
        '<button type="submit" class="btn btn-blue" style="width:100%;padding:12px;font-size:15px;margin-top:.5rem">Login</button>'
        "</form>"
        '<p style="text-align:center;margin-top:1.5rem;font-size:11px;color:#bbb">&#169; Scriptly</p>'
        "</div>"
    )
    return render("Login", body, nav=False)


@app.route("/logout")
def logout():
    session.clear()
    return redirect("/")


@app.route("/dashboard")
@login_required
def dashboard():
    if session.get("is_admin"):
        return redirect("/admin")

    with get_db() as db:
        perms = db.execute("SELECT script_id FROM permissions WHERE user_id=?", (session["user_id"],)).fetchall()

    allowed = [SCRIPTS[p["script_id"]] for p in perms if p["script_id"] in SCRIPTS]
    cards = ""
    for script in allowed:
        cards += (
            '<a href="/run/' + script["id"] + '" style="background:white;border-radius:16px;box-shadow:0 2px 16px rgba(0,0,0,.06);padding:1.5rem;text-decoration:none;display:block">'
            '<div style="font-size:36px;margin-bottom:.75rem">' + script["icon"] + "</div>"
            '<div style="font-size:15px;font-weight:700;color:#1e3a8a;margin-bottom:4px">' + script["name"] + "</div>"
            '<div style="font-size:12px;color:#64748b">' + script["desc"] + "</div>"
            "</a>"
        )

    if not allowed:
        cards = (
            '<div style="text-align:center;padding:3rem;color:#94a3b8">'
            '<div style="font-size:48px;margin-bottom:1rem">&#128274;</div>'
            "<div>No tools are available yet</div>"
            "</div>"
        )

    body = (
        '<h2 style="font-size:22px;font-weight:700;color:#1e3a8a;margin-bottom:.4rem">Hello, '
        + session["name"]
        + ' &#128075;</h2><p style="font-size:14px;color:#64748b;margin-bottom:2rem">Your available tools:</p>'
        '<div style="display:grid;grid-template-columns:repeat(auto-fill,minmax(200px,1fr));gap:1rem">'
        + cards
        + "</div>"
    )
    return render("My Tools", body)


@app.route("/run/<script_id>", methods=["GET", "POST"])
@login_required
def run_script(script_id):
    if session.get("is_admin"):
        return redirect("/admin")

    with get_db() as db:
        perm = db.execute(
            "SELECT 1 FROM permissions WHERE user_id=? AND script_id=?",
            (session["user_id"], script_id),
        ).fetchone()

    scr = get_script(script_id)

    if not perm or scr is None:
        add_flash("You do not have access to this tool")
        return redirect("/dashboard")
    result = None
    error = ""

    if request.method == "POST":
        file_obj = request.files.get("file")
        validation_error, ext = validate_upload(file_obj)
        if validation_error == "missing":
            error = '<div class="flash-err">' + scr["empty_error"] + '</div>'
        elif validation_error == "unsupported":
            error = '<div class="flash-err">' + scr["unsupported_error"] + '</div>'
        elif validation_error == "invalid_excel":
            error = '<div class="flash-err">' + scr["invalid_error"] + '</div>'
        elif validation_error == "empty":
            error = '<div class="flash-err">' + scr["empty_file_error"] + '</div>'
        elif validation_error == "too_large":
            error = '<div class="flash-err">' + scr["too_large_error"] + '</div>'
        else:
            uid = str(uuid.uuid4())[:8]
            inp = str(UPLOAD_FOLDER / f"{uid}.{ext}")
            result_name = build_output_filename(scr, uid)
            out = str(OUTPUT_FOLDER / result_name)
            options = {}
            extra_paths = []
            for field in scr.get("filter_fields", []):
                options[field["name"]] = request.form.get(field["name"], "").strip()
            for upload in scr.get("extra_uploads", []):
                extra_file = request.files.get(upload["name"])
                if extra_file and extra_file.filename:
                    extra_ext = get_extension(extra_file.filename)
                    expected = upload.get("accept", "").lstrip(".").lower()
                    if expected and extra_ext != expected:
                        error = '<div class="flash-err">׳¡׳•׳’ ׳”׳§׳•׳‘׳¥ ׳”׳ ׳•׳¡׳£ ׳׳™׳ ׳• ׳ ׳×׳׳</div>'
                        break
                    extra_path = str(UPLOAD_FOLDER / f"{uid}_{upload['name']}.{extra_ext or 'dat'}")
                    extra_file.save(extra_path)
                    options[f"{upload['name']}_path"] = extra_path
                    extra_paths.append(extra_path)
                elif upload.get("required"):
                    error = '<div class="flash-err">׳—׳¡׳¨ ׳§׳•׳‘׳¥ ׳ ׳•׳¡׳£ ׳ ׳“׳¨׳©</div>'
                    break
            if error:
                for path in extra_paths:
                    try:
                        os.remove(path)
                    except OSError:
                        pass
            else:
                file_obj.save(inp)
                try:
                    execute_script(scr, inp, out, ext, options)
                    result = result_name
                except (xlrd.biffh.XLRDError, BadZipFile, OSError, ValueError):
                    error = '<div class="flash-err">' + scr["processing_error"] + '</div>'
                except Exception as e:
                    error = '<div class="flash-err">Processing error: ' + str(e) + "</div>"
                finally:
                    try:
                        os.remove(inp)
                    except OSError:
                        pass
                    for path in extra_paths:
                        try:
                            os.remove(path)
                        except OSError:
                            pass

    if result:
        content = (
            '<div class="success-box">'
            '<div style="font-size:32px;margin-bottom:6px">&#9989;</div>'
            '<div style="font-size:16px;font-weight:700;color:#15803d;margin-bottom:10px">' + scr["success_title"] + '</div>'
            '<a href="/download/' + result + '" class="dl-btn">&#8681; ' + scr["success_action"] + '</a>'
            '<br><br><a href="/run/' + script_id + '" style="font-size:13px;color:#2563eb">' + scr["retry_action"] + '</a>'
            '</div>'
        )
    else:
        extra_uploads_html = ""
        for upload in scr.get("extra_uploads", []):
            extra_uploads_html += (
                '<div style="margin-top:1rem;padding-top:1rem;border-top:1px solid #e2e8f0">'
                + '<div style="font-size:14px;font-weight:600;color:#334155;margin-bottom:8px">' + upload["label"] + '</div>'
                + '<input type="file" name="' + upload["name"] + '" accept="' + upload.get("accept", "") + '" style="width:100%;max-width:420px;margin:0 auto 8px;display:block;font-family:inherit">'
                + '<div style="font-size:12px;color:#94a3b8">' + upload.get("help", "") + '</div>'
                + '</div>'
            )
        filter_fields_html = ""
        if scr.get("filter_fields"):
            filter_fields_html += '<div style="display:grid;grid-template-columns:repeat(auto-fit,minmax(180px,1fr));gap:10px;margin-bottom:1rem">'
            for field in scr.get("filter_fields", []):
                filter_fields_html += (
                    '<div><label class="field-label">' + field["label"] + '</label>'
                    + '<input type="text" name="' + field["name"] + '" placeholder="' + field.get("placeholder", "") + '" style="margin-bottom:0"></div>'
                )
            filter_fields_html += '</div>'
        content = (
            error
            + '<form method="POST" enctype="multipart/form-data" id="uploadForm">'
            + filter_fields_html
            + '<div style="background:#fafcff;border:2px dashed #c7d7f5;border-radius:14px;padding:1.5rem;margin-bottom:1rem;text-align:center">'
            + '<div style="font-size:32px;margin-bottom:8px">&#128194;</div>'
            + '<div style="font-size:15px;font-weight:600;color:#1e40af;margin-bottom:12px">' + scr["file_picker_label"] + '</div>'
            + '<input type="file" name="file" id="fi" accept="' + scr["accept"] + '" style="width:100%;max-width:420px;margin:0 auto 10px;display:block;font-family:inherit">'
            + '<div style="font-size:12px;color:#94a3b8" id="lbl">' + scr["accept"] + '</div>'
            + extra_uploads_html
            + '</div>'
            + '<button type="submit" class="btn btn-blue" id="gb" style="width:100%;padding:13px;font-size:15px;font-weight:700">' + scr["icon"] + ' ' + scr["submit_label"] + '</button>'
            + '<div class="processing-box" id="processingBox">'
            + '<div class="processing-note">' + scr["processing_title"] + '</div>'
            + '<div class="progress-track"><div class="progress-bar"></div></div>'
            + '<div class="processing-subnote">' + scr["processing_note"] + '</div>'
            + '</div>'
            + '</form>'
        )

    body = (
        '<a href="/dashboard" style="color:#2563eb;font-size:13px;text-decoration:none;display:block;margin-bottom:1rem">&#8592; ' + scr["back_label"] + '</a>'
        + '<div class="card">'
        + '<div style="font-size:40px;margin-bottom:.5rem">' + scr["icon"] + '</div>'
        + '<div style="font-size:20px;font-weight:700;color:#1e3a8a;margin-bottom:4px">' + scr["name"] + '</div>'
        + '<div style="font-size:13px;color:#64748b;margin-bottom:1.75rem">' + scr["desc"] + '</div>'
        + content
        + '</div>'
        + '<script>'
        + 'var fileInput=document.getElementById("fi");'
        + 'var label=document.getElementById("lbl");'
        + 'var button=document.getElementById("gb");'
        + 'var form=document.getElementById("uploadForm");'
        + 'if(fileInput && label){fileInput.addEventListener("change", function(){if(this.files && this.files.length){label.textContent=this.files[0].name;}});}'
        + 'if(form){form.addEventListener("submit", function(event){if(!fileInput || !fileInput.files || !fileInput.files.length){event.preventDefault();return false;}button.disabled=true;button.textContent="' + scr["processing_title"] + '";var box=document.getElementById("processingBox");if(box){box.classList.add("show");}return true;});}'
        + '</script>'
    )
    return render(scr["name"], body)

@app.route("/download/<filename>")
@login_required
def download(filename):
    path = OUTPUT_FOLDER / filename
    if not path.exists():
        add_flash("File not found")
        return redirect("/dashboard")
    download_name = filename.split("_", 1)[-1] if "_" in filename else filename
    return send_file(path, as_attachment=True, download_name=download_name)


@app.route("/admin")
@login_required
@admin_required
def admin():
    with get_db() as db:
        users = db.execute("SELECT * FROM users WHERE is_admin=0").fetchall()
        perms = db.execute("SELECT * FROM permissions").fetchall()

    user_perms = {}
    for perm in perms:
        user_perms.setdefault(perm["user_id"], set()).add(perm["script_id"])

    rows = ""
    for user in users:
        uid = user["id"]
        checks = ""
        for sid, script in SCRIPTS.items():
            checked = "checked" if (uid in user_perms and sid in user_perms[uid]) else ""
            checks += (
                '<label style="display:flex;align-items:center;gap:5px;font-size:13px;margin-left:10px">'
                '<input type="checkbox" name="scripts" value="' + sid + '" ' + checked + ">"
                + script["icon"]
                + " "
                + script["name"]
                + "</label>"
            )

        rows += (
            "<tr>"
            "<td><strong>" + user["full_name"] + "</strong></td>"
            '<td><span class="badge">' + user["username"] + "</span></td>"
            '<td><form method="POST" action="/admin/permissions/' + str(uid) + '" style="display:inline"><div style="display:flex;flex-wrap:wrap">'
            + checks
            + '</div><button type="submit" class="btn btn-gray" style="margin-top:6px;font-size:12px;padding:5px 12px">Save</button></form></td>'
            + '<td><div style="display:flex;gap:6px;flex-wrap:wrap">'
            + '<button type="button" class="btn btn-gray" style="font-size:12px;padding:5px 12px" onclick="openPass('
            + str(uid)
            + ')">Change password</button>'
            + '<form method="POST" action="/admin/resetpass/'
            + str(uid)
            + '" style="display:inline"><button type="submit" class="btn btn-gray" style="font-size:12px;padding:5px 12px">Temporary password</button></form>'
            '</div></td>'
            '<td><a href="/admin/delete/' + str(uid) + '" onclick="return confirm(\'Delete?\');" class="btn btn-red" style="text-decoration:none;font-size:12px;padding:5px 12px">Delete</a></td>'
            "</tr>"
        )

    table = (
        "<table><thead><tr><th>Name</th><th>Username</th><th>Permissions</th><th>Password</th><th>Delete</th></tr></thead><tbody>"
        + rows
        + "</tbody></table>"
    ) if users else '<p style="color:#94a3b8;text-align:center;padding:2rem">No users yet</p>'

    body = (
        '<div class="card"><h2>&#10133; Add New User</h2><form method="POST" action="/admin/add_user"><div class="form-row">'
        '<div class="form-group"><label class="field-label">Full Name</label><input type="text" name="full_name" placeholder="Customer name" required style="margin-bottom:0"></div>'
        '<div class="form-group"><label class="field-label">Username</label><input type="text" name="username" placeholder="Login username" required style="margin-bottom:0"></div>'
        '<div class="form-group"><label class="field-label">Password</label><input type="password" name="password" placeholder="Initial password" required style="margin-bottom:0"></div>'
        '<button type="submit" class="btn btn-blue" style="height:40px;align-self:flex-end">Add</button></div></form></div>'
        '<div class="card"><h2>&#128101; Users In System</h2>'
        + table
        + '</div><div class="modal-bg" id="passModal"><div class="modal-box"><h3 style="font-size:15px;font-weight:700;margin-bottom:1rem;color:#1e3a8a">Change Password &#8212; <span id="pname"></span></h3>'
        '<form method="POST" id="pform"><input type="password" name="new_password" placeholder="New password" required>'
        '<div style="display:flex;gap:8px;margin-top:.5rem;justify-content:flex-end"><button type="button" class="btn btn-gray" onclick="closePass()">Cancel</button>'
        '<button type="submit" class="btn btn-blue">Update</button></div></form></div></div>'
        '<script>function openPass(id,name){document.getElementById("pname").textContent=name||"";document.getElementById("pform").action="/admin/setpass/"+id;document.getElementById("passModal").style.display="flex";}function closePass(){document.getElementById("passModal").style.display="none";}</script>'
    )
    return render("Admin", body)


@app.route("/admin/add_user", methods=["POST"])
@login_required
@admin_required
def add_user():
    username = request.form["username"].strip()
    password = request.form["password"]
    full_name = request.form["full_name"].strip()
    try:
        with get_db() as db:
            db.execute(
                "INSERT INTO users(username,password,full_name) VALUES (?,?,?)",
                (username, generate_password_hash(password), full_name),
            )
            db.commit()
        add_flash("User " + full_name + " was created successfully")
    except Exception as exc:
        if not is_integrity_error(exc):
            raise
        add_flash("Username already exists")
    return redirect("/admin")


@app.route("/admin/delete/<int:uid>")
@login_required
@admin_required
def delete_user(uid):
    with get_db() as db:
        db.execute("DELETE FROM users WHERE id=?", (uid,))
        db.execute("DELETE FROM permissions WHERE user_id=?", (uid,))
        db.commit()
    add_flash("User deleted")
    return redirect("/admin")


@app.route("/admin/setpass/<int:uid>", methods=["POST"])
@login_required
@admin_required
def set_password(uid):
    with get_db() as db:
        db.execute(
            "UPDATE users SET password=? WHERE id=?",
            (generate_password_hash(request.form["new_password"]), uid),
        )
        db.commit()
    add_flash("Password updated")
    return redirect("/admin")


@app.route("/admin/resetpass/<int:uid>", methods=["POST"])
@login_required
@admin_required
def reset_password(uid):
    temp_password = generate_temp_password()
    with get_db() as db:
        user = db.execute("SELECT full_name FROM users WHERE id=?", (uid,)).fetchone()
        db.execute(
            "UPDATE users SET password=? WHERE id=?",
            (generate_password_hash(temp_password), uid),
        )
        db.commit()
    name = user["full_name"] if user else str(uid)
    add_flash("Temporary password for " + name + ": " + temp_password)
    return redirect("/admin")


@app.route("/admin/permissions/<int:uid>", methods=["POST"])
@login_required
@admin_required
def set_permissions(uid):
    selected = request.form.getlist("scripts")
    with get_db() as db:
        db.execute("DELETE FROM permissions WHERE user_id=?", (uid,))
        for script_id in selected:
            if script_id in SCRIPTS:
                db.execute(
                    "INSERT OR IGNORE INTO permissions(user_id,script_id) VALUES (?,?)",
                    (uid, script_id),
                )
        db.commit()
    add_flash("Permissions updated")
    return redirect("/admin")


if __name__ == "__main__":
    app.run(debug=False, use_reloader=False)

