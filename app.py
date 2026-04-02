from pathlib import Path
from zipfile import BadZipFile, ZIP_DEFLATED, ZipFile
from collections import defaultdict
import calendar
import csv
from datetime import date, datetime, timedelta
import html
import json
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
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_AUTO_SHAPE_TYPE, MSO_CONNECTOR
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt

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
ALLOWED_EXTENSIONS = {"xls", "xlsx", "csv"}
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

    file_storage.stream.seek(0, os.SEEK_END)
    size = file_storage.stream.tell()
    file_storage.stream.seek(0)
    if size <= 0:
        return "empty", None
    if size > MAX_UPLOAD_SIZE:
        return "too_large", None

    if ext == "csv":
        return None, ext

    detected = detect_excel_signature(file_storage)
    if detected != ext:
        return "invalid_excel", None

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


PAYABLE_HOUR_LABELS = {"רגילות", "׳¨׳’׳™׳׳•׳×", "100%", "125%", "150%", "175%", "200%"}
REGULAR_PAYABLE_HOUR_LABELS = {"רגילות", "׳¨׳’׳™׳׳•׳×"}


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
    regular_hours = [totals[label] for label in REGULAR_PAYABLE_HOUR_LABELS if totals.get(label) is not None]
    if regular_hours:
        return regular_hours[0], totals
    available = [value for value in totals.values() if value is not None]
    if not available:
        return None, totals
    return sum(available), totals


def extract_flamingo_worker_pair(detail_sheet, summary_sheet):
    worker_name = ""
    for label in ("שם לתצוגה", "׳©׳ ׳׳×׳¦׳•׳’׳”"):
        worker_name = str(find_row_label_value_with_offsets(detail_sheet, 5, label, [2, 1])).strip()
        if worker_name:
            break
    worker_name = worker_name or detail_sheet.name
    department = ""
    rate_raw = ""
    worker_number = ""
    id_number = ""
    start_date = ""
    for label in ("מחלקה", "׳׳—׳׳§׳”"):
        department = str(find_row_label_value_with_offsets(detail_sheet, 5, label, [3, 2, 1])).strip()
        if department:
            break
    for rate_label in ("הערות", "׳”׳¢׳¨׳•׳×"):
        rate_raw = find_row_label_value_with_offsets(detail_sheet, 5, rate_label, [4])
        if rate_raw not in ("", None):
            break
    for label in ("מספר בשכר", "מס' מפעל בשכר", "׳׳¡׳₪׳¨ ׳‘׳©׳›׳¨"):
        worker_number = find_row_label_value_with_offsets(detail_sheet, 5, label, [5, 4, 3, 2, 1])
        if worker_number not in ("", None):
            break
    for label in ("תעודת זהות", "׳×׳¢׳•׳“׳× ׳–׳”׳•׳×"):
        id_number = find_row_label_value_with_offsets(detail_sheet, 5, label, [2, 1])
        if id_number not in ("", None):
            break
    for label in ("תחילת עבודה", "׳×׳—׳™׳׳× ׳¢׳‘׳•׳“׳”"):
        start_date = find_row_label_value_with_offsets(detail_sheet, 5, label, [4, 3, 2, 1])
        if start_date not in ("", None):
            break
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
        notes.append("יש לעדכן את התעריף בשדה הערות ולייצא את הדוח מחדש.")

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
            action = "יש לעדכן את התעריף בשדה הערות ולייצא את הדוח מחדש."
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
            employee_number = (row.get("שכר") or row.get("׳©׳›׳¨") or "").strip()
            employee_id = (row.get("ת.ז") or row.get("׳×.׳–") or "").strip()
            entry = {
                "employee_number": employee_number,
                "id_number": employee_id,
                "employee_name": (row.get("שם עובד") or row.get("׳©׳ ׳¢׳•׳‘׳“") or "").strip(),
                "direct_manager": (row.get("מנהל ישיר") or row.get("׳׳ ׳”׳ ׳™׳©׳™׳¨") or "").strip(),
                "department": (row.get("מחלקה") or row.get("׳׳—׳׳§׳”") or "").strip(),
                "agreement_name": (row.get("שם הסכם") or row.get("׳©׳ ׳”׳¡׳›׳") or "").strip(),
                "agreement_number": (row.get("מס' הסכם") or row.get("׳׳¡' ׳”׳¡׳›׳") or "").strip(),
            }
            if employee_number:
                records[("number", employee_number)] = entry
            if employee_id:
                records[("id", employee_id)] = entry
            if not employee_number and not employee_id:
                unmatched.append(entry)
    return records, unmatched


def parse_org_hierarchy_csv(csv_path):
    rows = []
    with open(csv_path, "r", encoding="utf-8-sig", newline="") as handle:
        reader = csv.DictReader(handle)
        for raw in reader:
            employee_name = (raw.get("שם עובד") or raw.get("׳©׳ ׳¢׳•׳‘׳“") or "").strip()
            employee_number = (raw.get("שכר") or raw.get("׳©׳›׳¨") or "").strip()
            id_number = (raw.get("ת.ז") or raw.get("׳×.׳–") or "").strip()
            direct_manager = (raw.get("מנהל ישיר") or raw.get("׳׳ ׳”׳ ׳™׳©׳™׳¨") or "").strip()
            manager_flag = (raw.get("מנהל") or "").strip()
            department = (raw.get("מחלקה") or raw.get("׳׳—׳׳§׳”") or "").strip()
            email = (raw.get("אימייל") or "").strip()
            if not any([employee_name, employee_number, id_number, direct_manager, department, email]):
                continue
            rows.append(
                {
                    "employee_name": employee_name,
                    "employee_number": employee_number,
                    "id_number": id_number,
                    "direct_manager": direct_manager,
                    "is_manager": manager_flag == "[+]",
                    "department": department,
                    "email": email,
                }
            )

    name_counts = defaultdict(int)
    id_counts = defaultdict(int)
    for row in rows:
        if row["employee_name"]:
            name_counts[row["employee_name"]] += 1
        if row["id_number"]:
            id_counts[row["id_number"]] += 1

    unique_name_map = {row["employee_name"]: row for row in rows if row["employee_name"] and name_counts[row["employee_name"]] == 1}
    children_map = defaultdict(list)
    exception_rows = []
    roots = []

    for row in rows:
        if not row["direct_manager"]:
            roots.append(row)
            exception_rows.append(
                {
                    "category": "צומת שורש",
                    "employee_name": row["employee_name"],
                    "employee_number": row["employee_number"],
                    "detail": "מנהל ישיר ריק",
                }
            )
            continue
        if row["direct_manager"] not in unique_name_map:
            roots.append(row)
            exception_rows.append(
                {
                    "category": "מנהל לא נמצא",
                    "employee_name": row["employee_name"],
                    "employee_number": row["employee_number"],
                    "detail": row["direct_manager"],
                }
            )
            continue
        children_map[row["direct_manager"]].append(row)

    for name, count in sorted(name_counts.items()):
        if count > 1:
            exception_rows.append(
                {
                    "category": "שם כפול",
                    "employee_name": name,
                    "employee_number": "",
                    "detail": f"נמצאו {count} שורות עם אותו שם עובד",
                }
            )
    for id_number, count in sorted(id_counts.items()):
        if id_number and count > 1:
            exception_rows.append(
                {
                    "category": "תעודת זהות כפולה",
                    "employee_name": "",
                    "employee_number": "",
                    "detail": f"תעודת הזהות {id_number} מופיעה {count} פעמים",
                }
            )

    visited = set()
    summary_rows = []
    tree_rows = []

    def walk(node, depth, root_name, root_key, stack):
        node_key = (node["employee_number"], node["id_number"], node["employee_name"])
        if node_key in stack:
            exception_rows.append(
                {
                    "category": "זוהתה לולאה",
                    "employee_name": node["employee_name"],
                    "employee_number": node["employee_number"],
                    "detail": "זוהתה לולאה בהיררכיה במהלך מעבר על מנהלים",
                }
            )
            return

        visited.add(node_key)
        summary_rows.append(
            {
                "employee_name": node["employee_name"],
                "employee_number": node["employee_number"],
                "id_number": node["id_number"],
                "direct_manager": node["direct_manager"],
                "is_manager": "כן" if node["is_manager"] else "לא",
                "department": node["department"],
                "email": node["email"],
                "depth": depth,
                "root_name": root_name,
                "root_employee_number": root_key[0],
                "root_id_number": root_key[1],
            }
        )
        tree_rows.append(
            {
                "root_name": root_name,
                "depth": depth,
                "display_name": ("    " * depth) + node["employee_name"],
                "employee_number": node["employee_number"],
                "department": node["department"],
                "direct_manager": node["direct_manager"],
            }
        )
        next_stack = stack | {node_key}
        for child in sorted(children_map.get(node["employee_name"], []), key=lambda item: (item["employee_name"], item["employee_number"])):
            walk(child, depth + 1, root_name, root_key, next_stack)

    unique_roots = []
    seen_root_keys = set()
    for row in roots:
        root_key = (row["employee_number"], row["id_number"], row["employee_name"])
        if root_key not in seen_root_keys:
            unique_roots.append(row)
            seen_root_keys.add(root_key)

    for root in sorted(unique_roots, key=lambda item: (item["employee_name"], item["employee_number"])):
        walk(root, 0, root["employee_name"], (root["employee_number"], root["id_number"]), set())

    for row in rows:
        node_key = (row["employee_number"], row["id_number"], row["employee_name"])
        if node_key not in visited:
            exception_rows.append(
                {
                    "category": "צומת לא מקושר",
                    "employee_name": row["employee_name"],
                    "employee_number": row["employee_number"],
                    "detail": "השורה לא קושרה לאף צומת שורש",
                }
            )
            walk(row, 0, row["employee_name"], (row["employee_number"], row["id_number"]), set())

    stats = {
        "employee_count": len(rows),
        "root_count": len(unique_roots),
        "exception_count": len(exception_rows),
    }
    return summary_rows, tree_rows, exception_rows, stats


def write_org_hierarchy_summary(ws, summary_rows, stats, root_rows):
    ws.title = safe_sheet_title("סיכום מבנה ארגוני", "Org Summary")
    ws.sheet_view.rightToLeft = True
    ws.sheet_view.showGridLines = False
    ws.freeze_panes = "A11"

    ws["A1"] = "דוח מבנה ארגוני"
    ws["A1"].font = Font(bold=True, size=18, color="0F172A")
    ws["A1"].fill = PatternFill(fill_type="solid", fgColor="BFDBFE")

    direct_reports = defaultdict(int)
    department_counts = defaultdict(int)
    manager_rows = []
    for row in summary_rows:
        if row["direct_manager"]:
            direct_reports[row["direct_manager"]] += 1
        if row["department"]:
            department_counts[row["department"]] += 1
        if row["is_manager"] == "כן":
            manager_rows.append(row)

    metrics = [
        ("סה\"כ עובדים", stats["employee_count"], "DBEAFE"),
        ("סה\"כ שורשים (ללא מנהל ישיר או עם מנהל לא מזוהה)", stats["root_count"], "DCFCE7"),
        ("סה\"כ מנהלים", len(manager_rows), "EDE9FE"),
        ("סה\"כ חריגים", stats["exception_count"], "FEE2E2"),
    ]
    for idx, (label, value, fill_color) in enumerate(metrics, start=3):
        ws.cell(row=idx, column=1, value=label).font = Font(bold=True, color="334155")
        ws.cell(row=idx, column=2, value=value).font = Font(bold=True, color="0F172A")
        ws.cell(row=idx, column=1).fill = PatternFill(fill_type="solid", fgColor=fill_color)
        ws.cell(row=idx, column=2).fill = PatternFill(fill_type="solid", fgColor=fill_color)

    ws["D3"] = "שורשים שזוהו"
    ws["D3"].font = Font(bold=True, color="1E3A8A")
    ws["D3"].fill = PatternFill(fill_type="solid", fgColor="DBEAFE")
    ws["D4"] = ", ".join(root["employee_name"] for root in sorted(root_rows, key=lambda item: (item["employee_name"], item["employee_number"]))) or "אין"

    ws["D6"] = "רשימת מנהלים ומספר כפיפים ישירים"
    ws["D6"].font = Font(bold=True, color="1E3A8A")
    ws["D6"].fill = PatternFill(fill_type="solid", fgColor="EDE9FE")
    ws["D7"] = "מנהל"
    ws["E7"] = "כפיפים ישירים"
    for cell_ref in ("D7", "E7"):
        ws[cell_ref].font = Font(bold=True, color="FFFFFF")
        ws[cell_ref].fill = PatternFill(fill_type="solid", fgColor="7C3AED")
    for row_idx, row in enumerate(sorted(manager_rows, key=lambda item: (-direct_reports.get(item["employee_name"], 0), item["employee_name"])), start=8):
        ws.cell(row=row_idx, column=4, value=row["employee_name"])
        ws.cell(row=row_idx, column=5, value=direct_reports.get(row["employee_name"], 0))
        if row_idx % 2 == 0:
            ws.cell(row=row_idx, column=4).fill = PatternFill(fill_type="solid", fgColor="F5F3FF")
            ws.cell(row=row_idx, column=5).fill = PatternFill(fill_type="solid", fgColor="F5F3FF")

    dept_start_row = 8
    ws["G6"] = "רשימת מחלקות וכמות עובדים"
    ws["G6"].font = Font(bold=True, color="0F766E")
    ws["G6"].fill = PatternFill(fill_type="solid", fgColor="CCFBF1")
    ws["G7"] = "מחלקה"
    ws["H7"] = "כמות עובדים"
    for cell_ref in ("G7", "H7"):
        ws[cell_ref].font = Font(bold=True, color="FFFFFF")
        ws[cell_ref].fill = PatternFill(fill_type="solid", fgColor="0F766E")
    for row_idx, (department, count) in enumerate(sorted(department_counts.items(), key=lambda item: (-item[1], item[0])), start=dept_start_row):
        ws.cell(row=row_idx, column=7, value=department)
        ws.cell(row=row_idx, column=8, value=count)
        if row_idx % 2 == 0:
            ws.cell(row=row_idx, column=7).fill = PatternFill(fill_type="solid", fgColor="ECFDF5")
            ws.cell(row=row_idx, column=8).fill = PatternFill(fill_type="solid", fgColor="ECFDF5")

    headers = [
        "שם עובד",
        "מספר שכר",
        "תעודת זהות",
        "מנהל ישיר",
        "האם מנהל",
        "מחלקה",
        "אימייל",
        "עומק בעץ",
        "שם שורש",
    ]
    header_row = 10
    for col, header in enumerate(headers, start=1):
        cell = ws.cell(row=header_row, column=col, value=header)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill(fill_type="solid", fgColor="1E3A8A")

    for row_idx, row in enumerate(sorted(summary_rows, key=lambda item: (item["root_name"], item["depth"], item["employee_name"])), start=header_row + 1):
        values = [
            row["employee_name"],
            row["employee_number"],
            row["id_number"],
            row["direct_manager"],
            row["is_manager"],
            row["department"],
            row["email"],
            row["depth"],
            row["root_name"],
        ]
        for col, value in enumerate(values, start=1):
            ws.cell(row=row_idx, column=col, value=value)
            if row_idx % 2 == 0:
                ws.cell(row=row_idx, column=col).fill = PatternFill(fill_type="solid", fgColor="F8FAFC")

    widths = [24, 14, 16, 22, 12, 24, 26, 10, 20]
    for col, width in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(col)].width = width
    ws.column_dimensions["D"].width = 28
    ws.column_dimensions["E"].width = 14
    ws.column_dimensions["G"].width = 28
    ws.column_dimensions["H"].width = 14


def write_org_hierarchy_tree(ws, tree_rows):
    ws.title = safe_sheet_title("טבלת היררכיה", "Hierarchy Table")
    ws.sheet_view.rightToLeft = True
    ws.sheet_view.showGridLines = False
    ws.freeze_panes = "A5"

    ws["A1"] = "טבלת היררכיה ארגונית"
    ws["A1"].font = Font(bold=True, size=18, color="0F172A")
    ws["A1"].fill = PatternFill(fill_type="solid", fgColor="CCFBF1")
    ws["A2"] = "הטבלה מציגה נתיב דיווח ורמת עומק. ההזחה מסייעת לקריאה, אך אינה תרשים חזותי מלא."
    ws["A2"].font = Font(bold=True, color="0F766E")

    headers = ["שם עובד בהזחה לפי רמה", "רמה", "מספר שכר", "מחלקה", "מנהל ישיר"]
    for col, header in enumerate(headers, start=1):
        cell = ws.cell(row=4, column=col, value=header)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill(fill_type="solid", fgColor="0F766E")

    row_idx = 5
    current_root = None
    for row in sorted(tree_rows, key=lambda item: (item["root_name"], item["depth"], item["display_name"])):
        if row["root_name"] != current_root:
            if current_root is not None:
                row_idx += 1
            ws.cell(row=row_idx, column=1, value=f"מקטע שורש: {row['root_name']}")
            ws.cell(row=row_idx, column=1).font = Font(bold=True, color="1E3A8A")
            ws.cell(row=row_idx, column=1).fill = PatternFill(fill_type="solid", fgColor="DBEAFE")
            current_root = row["root_name"]
            row_idx += 1
        values = [row["display_name"], row["depth"], row["employee_number"], row["department"], row["direct_manager"]]
        for col, value in enumerate(values, start=1):
            ws.cell(row=row_idx, column=col, value=value)
            if row_idx % 2 == 0:
                ws.cell(row=row_idx, column=col).fill = PatternFill(fill_type="solid", fgColor="ECFDF5")
        row_idx += 1

    widths = [40, 10, 14, 24, 22]
    for col, width in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(col)].width = width


def write_org_hierarchy_exceptions(ws, exception_rows, root_rows):
    ws.title = safe_sheet_title("חריגים", "Exceptions")
    ws.sheet_view.rightToLeft = True
    ws.sheet_view.showGridLines = False
    ws.freeze_panes = "A2"

    headers = ["סוג חריג", "שם עובד", "מספר שכר", "פירוט"]
    for col, header in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=col, value=header)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill(fill_type="solid", fgColor="B91C1C")

    row_idx = 2
    for row in exception_rows:
        values = [row["category"], row["employee_name"], row["employee_number"], row["detail"]]
        for col, value in enumerate(values, start=1):
            ws.cell(row=row_idx, column=col, value=value)
            if row_idx % 2 == 0:
                ws.cell(row=row_idx, column=col).fill = PatternFill(fill_type="solid", fgColor="FEF2F2")
        row_idx += 1

    if row_idx == 2:
        ws.cell(row=row_idx, column=1, value="אין חריגים")
        row_idx += 1

    row_idx += 1
    ws.cell(row=row_idx, column=1, value="רשימת שורשים").font = Font(bold=True, color="1E3A8A")
    ws.cell(row=row_idx, column=1).fill = PatternFill(fill_type="solid", fgColor="DBEAFE")
    row_idx += 1
    for root in sorted(root_rows, key=lambda item: (item["employee_name"], item["employee_number"])):
        ws.cell(row=row_idx, column=1, value=root["employee_name"])
        ws.cell(row=row_idx, column=2, value=root["employee_number"])
        ws.cell(row=row_idx, column=3, value=root["department"])
        row_idx += 1

    widths = [20, 24, 14, 36]
    for col, width in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(col)].width = width


def build_org_summary_slide_lines(summary_rows, stats):
    manager_count = sum(1 for row in summary_rows if row["is_manager"] == "כן")
    return [
        f"סה\"כ עובדים: {stats['employee_count']}",
        f"סה\"כ שורשים: {stats['root_count']}",
        f"סה\"כ מנהלים: {manager_count}",
        f"סה\"כ חריגים: {stats['exception_count']}",
    ]


def build_org_root_slide_lines(tree_rows, root_name):
    lines = []
    for row in tree_rows:
        if row["root_name"] != root_name:
            continue
        meta = []
        if row["department"]:
            meta.append(row["department"])
        if row["employee_number"]:
            meta.append(f"מס' שכר {row['employee_number']}")
        line = row["display_name"].strip()
        if meta:
            line += " | " + " | ".join(meta)
        lines.append((row["depth"], line))
    return lines


def build_org_exceptions_slide_lines(exception_rows):
    if not exception_rows:
        return ["אין חריגים"]
    return [f"{row['category']}: {row['employee_name']} | {row['detail']}".strip(" |") for row in exception_rows]


def format_display_name(name):
    parts = str(name or "").split()
    if len(parts) <= 1:
        return str(name or "").strip()
    return " ".join(reversed(parts))


def build_org_root_nodes(summary_rows, root_name):
    nodes = []
    for row in summary_rows:
        row_root_key = (row.get("root_employee_number", ""), row.get("root_id_number", ""), row["root_name"])
        if row_root_key != root_name:
            continue
        node_key = (row["employee_number"], row["id_number"], row["employee_name"])
        nodes.append(
            {
                "node_key": node_key,
                "employee_name": row["employee_name"],
                "employee_number": row["employee_number"],
                "id_number": row["id_number"],
                "department": row["department"],
                "direct_manager": row["direct_manager"],
                "depth": row["depth"],
            }
        )
    keys_by_name = defaultdict(list)
    for node in nodes:
        keys_by_name[node["employee_name"]].append(node["node_key"])
    children_map = defaultdict(list)
    for node in nodes:
        manager_keys = keys_by_name.get(node["direct_manager"], [])
        node["parent_key"] = manager_keys[0] if len(manager_keys) == 1 else None
        node["manager_resolution"] = "unique" if len(manager_keys) == 1 else ("ambiguous" if len(manager_keys) > 1 else "missing")
        children_map[node["parent_key"]].append(node)
    for parent_key, children in children_map.items():
        children.sort(key=lambda item: (item["employee_name"], item["department"]))
    top_nodes = [node for node in nodes if node.get("parent_key") is None]
    metrics = {}
    for top_node in top_nodes:
        compute_org_subtree_metrics(top_node, children_map, metrics)
    for node in nodes:
        node["direct_reports_count"] = metrics.get(node["node_key"], {}).get("direct_reports_count", 0)
        node["subtree_employee_count"] = metrics.get(node["node_key"], {}).get("subtree_employee_count", 1)
    return nodes, children_map


def compute_org_subtree_metrics(node, children_map, metrics):
    children = children_map.get(node["node_key"], [])
    subtree_total = 1
    for child in children:
        subtree_total += compute_org_subtree_metrics(child, children_map, metrics)
    metrics[node["node_key"]] = {
        "direct_reports_count": len(children),
        "subtree_employee_count": subtree_total,
    }
    return subtree_total


def compute_org_subtree_widths(node, children_map, widths):
    children = children_map.get(node["node_key"], [])
    if not children:
        widths[node["node_key"]] = 1.0
        return 1.0
    total_width = 0.0
    for child in children:
        total_width += compute_org_subtree_widths(child, children_map, widths)
    widths[node["node_key"]] = max(total_width, 1.0)
    return widths[node["node_key"]]


def layout_org_chart(node, children_map, widths, positions, left_units=0.0):
    node_width = widths[node["node_key"]]
    children = children_map.get(node["node_key"], [])
    if not children:
        positions[node["node_key"]] = left_units + (node_width / 2.0)
        return
    current_left = left_units
    for child in children:
        child_width = widths[child["node_key"]]
        layout_org_chart(child, children_map, widths, positions, current_left)
        current_left += child_width
    first_child = children[0]["node_key"]
    last_child = children[-1]["node_key"]
    positions[node["node_key"]] = (positions[first_child] + positions[last_child]) / 2.0


def build_org_chart_subset(top_nodes, children_map, level_limit=None):
    subset_nodes = []
    subset_keys = set()

    def visit(node, chart_depth):
        node_copy = dict(node)
        node_copy["chart_depth"] = chart_depth
        subset_nodes.append(node_copy)
        subset_keys.add(node["node_key"])
        if level_limit is not None and chart_depth >= level_limit:
            return
        for child in children_map.get(node["node_key"], []):
            visit(child, chart_depth + 1)

    for top_node in top_nodes:
        visit(top_node, 0)

    subset_children_map = defaultdict(list)
    for node in subset_nodes:
        parent_key = node.get("parent_key")
        if parent_key not in subset_keys:
            parent_key = None
        subset_children_map[parent_key].append(node)
    for parent_key, children in subset_children_map.items():
        children.sort(key=lambda item: (item["employee_name"], item["department"]))
    return subset_nodes, subset_children_map


def build_org_detail_slide_nodes(root_node, children_map):
    detail_nodes = []

    def visit(node):
        if node["direct_reports_count"] > 0 and (
            node["direct_reports_count"] >= 5 or node["subtree_employee_count"] >= 12
        ):
            detail_nodes.append(node)
        for child in children_map.get(node["node_key"], []):
            visit(child)

    for child in children_map.get(root_node["node_key"], []):
        visit(child)
    return detail_nodes


def build_department_summary_rows(summary_rows):
    department_rows = []
    grouped = defaultdict(list)
    for row in summary_rows:
        grouped[row["department"] or "ללא מחלקה"].append(row)
    for department, rows in sorted(grouped.items(), key=lambda item: (-len(item[1]), item[0])):
        managers = [row for row in rows if row["is_manager"] == "כן"]
        manager_names = [format_display_name(row["employee_name"]) for row in sorted(managers, key=lambda item: item["employee_name"])[:4]]
        department_rows.append(
            {
                "department": department,
                "employee_count": len(rows),
                "manager_count": len(managers),
                "manager_names": manager_names,
            }
        )
    return department_rows


def build_manager_summary_rows(summary_rows, root_nodes):
    manager_rows = []
    for root_node in root_nodes:
        nodes, children_map = build_org_root_nodes(summary_rows, root_node["root_key"])
        node_map = {node["node_key"]: node for node in nodes}
        for node in sorted(nodes, key=lambda item: (-item["subtree_employee_count"], item["employee_name"])):
            if node["direct_reports_count"] <= 0:
                continue
            direct_reports = children_map.get(node["node_key"], [])
            manager_rows.append(
                {
                    "manager_name": format_display_name(node["employee_name"]),
                    "department": node["department"] or "ללא מחלקה",
                    "direct_reports_count": node["direct_reports_count"],
                    "subtree_employee_count": node["subtree_employee_count"],
                    "employee_names": [format_display_name(child["employee_name"]) for child in direct_reports],
                }
            )
    return manager_rows


def add_pptx_bullets(text_frame, lines, font_size=18):
    text_frame.clear()
    if not lines:
        lines = [""]
    for index, line in enumerate(lines):
        paragraph = text_frame.paragraphs[0] if index == 0 else text_frame.add_paragraph()
        paragraph.text = str(line.get("text") if isinstance(line, dict) else line)
        paragraph.level = int(line.get("level", 0)) if isinstance(line, dict) else 0
        paragraph.alignment = PP_ALIGN.RIGHT
        if paragraph.runs:
            run = paragraph.runs[0]
            run.font.size = Pt(line.get("size", font_size) if isinstance(line, dict) else font_size)
            run.font.bold = bool(line.get("bold", False)) if isinstance(line, dict) else False


def add_pptx_slide(prs, title, body_lines):
    slide = prs.slides.add_slide(prs.slide_layouts[1])
    slide.shapes.title.text = title
    title_frame = slide.shapes.title.text_frame
    for paragraph in title_frame.paragraphs:
        paragraph.alignment = PP_ALIGN.RIGHT
        if paragraph.runs:
            paragraph.runs[0].font.size = Pt(26)
            paragraph.runs[0].font.bold = True
    body = slide.placeholders[1]
    add_pptx_bullets(body.text_frame, body_lines, font_size=18)
    return slide


def add_org_chart_slide(prs, title, top_nodes, children_map, note_text, level_limit=2, compact=False):
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    title_box = slide.shapes.add_textbox(Inches(0.45), Inches(0.2), Inches(12.2), Inches(0.55))
    title_frame = title_box.text_frame
    title_frame.clear()
    title_paragraph = title_frame.paragraphs[0]
    title_paragraph.text = title
    title_paragraph.alignment = PP_ALIGN.RIGHT
    if title_paragraph.runs:
        title_run = title_paragraph.runs[0]
        title_run.font.size = Pt(24)
        title_run.font.bold = True

    note_box = slide.shapes.add_textbox(Inches(0.45), Inches(0.78), Inches(12.2), Inches(0.3))
    note_frame = note_box.text_frame
    note_frame.clear()
    note_paragraph = note_frame.paragraphs[0]
    note_paragraph.text = note_text
    note_paragraph.alignment = PP_ALIGN.RIGHT
    if note_paragraph.runs:
        note_run = note_paragraph.runs[0]
        note_run.font.size = Pt(11)
        note_run.font.color.rgb = RGBColor(71, 85, 105)

    nodes, children_map = build_org_chart_subset(top_nodes, children_map, level_limit=level_limit)
    if not nodes:
        return slide, {}

    box_nodes = [node for node in nodes if node["direct_reports_count"] > 0]
    if not box_nodes:
        return slide, {}
    box_node_keys = {node["node_key"] for node in box_nodes}
    top_node_keys = {node["node_key"] for node in box_nodes if node.get("chart_depth", 0) == 0}
    box_children_map = defaultdict(list)
    for parent_key, children in children_map.items():
        visible_box_children = [child for child in children if child["node_key"] in box_node_keys]
        if visible_box_children or parent_key in box_node_keys or parent_key is None:
            box_children_map[parent_key].extend(visible_box_children)
            box_children_map[parent_key].sort(key=lambda item: (item["employee_name"], item["department"]))

    widths = {}
    total_units = 0.0
    left_units = 0.0
    positions = {}
    for top_node in sorted([node for node in box_nodes if node.get("chart_depth", 0) == 0], key=lambda item: (item["employee_name"], item["department"])):
        subtree_width = compute_org_subtree_widths(top_node, box_children_map, widths)
        layout_org_chart(top_node, box_children_map, widths, positions, left_units)
        left_units += subtree_width
        total_units += subtree_width

    slide_width = prs.slide_width
    slide_height = prs.slide_height
    left_margin = Inches(0.3)
    right_margin = Inches(0.3)
    top_margin = Inches(1.2)
    bottom_margin = Inches(0.35)
    chart_width = slide_width - left_margin - right_margin
    chart_height = slide_height - top_margin - bottom_margin
    max_depth = max((row.get("chart_depth", 0) for row in box_nodes), default=0)
    level_count = max_depth + 1
    vertical_gap = chart_height / max(level_count, 1)
    total_units = max(total_units, 1.0)

    if compact:
        base_box_height = min(int(Inches(0.62)), max(int(Inches(0.34)), int(vertical_gap * 0.52)))
        manager_box_height = min(int(Inches(0.84)), max(int(Inches(0.48)), int(vertical_gap * 0.72)))
        min_box_width = int(Inches(1.05))
        max_box_width = int(Inches(1.6))
        name_size_regular = 9.5
        name_size_manager = 9
        dept_size_regular = 7.5
        dept_size_manager = 7
        metrics_size = 6.5
    else:
        base_box_height = min(int(Inches(0.72)), max(int(Inches(0.42)), int(vertical_gap * 0.58)))
        manager_box_height = min(int(Inches(0.96)), max(int(Inches(0.56)), int(vertical_gap * 0.82)))
        min_box_width = int(Inches(1.35))
        max_box_width = int(Inches(1.95))
        name_size_regular = 11
        name_size_manager = 10.5
        dept_size_regular = 9
        dept_size_manager = 8.5
        metrics_size = 7.5
    shape_bounds = {}
    node_shapes = {}

    sorted_nodes = sorted(box_nodes, key=lambda item: (item.get("chart_depth", 0), positions.get(item["node_key"], 0.0), item["employee_name"]))
    for node in sorted_nodes:
        node_width_units = widths.get(node["node_key"], 1.0)
        center_unit = positions.get(node["node_key"], 0.5)
        depth = node.get("chart_depth", 0)
        children = box_children_map.get(node["node_key"], [])

        center_x = int(left_margin + (center_unit / total_units) * chart_width)
        center_y = int(top_margin + depth * vertical_gap + vertical_gap * 0.45)
        width_by_subtree = int((node_width_units / total_units) * chart_width * 0.72)
        box_width = max(min_box_width, min(max_box_width, width_by_subtree))
        box_height = manager_box_height if node["direct_reports_count"] > 0 else base_box_height
        x = max(int(left_margin), min(int(slide_width - right_margin - box_width), center_x - box_width // 2))
        y = max(int(top_margin), min(int(slide_height - bottom_margin - box_height), center_y - box_height // 2))

        shape = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE, x, y, box_width, box_height)
        shape.line.color.rgb = RGBColor(100, 116, 139)
        shape.line.width = Pt(1)
        shape.fill.solid()
        if depth == 0:
            shape.fill.fore_color.rgb = RGBColor(191, 219, 254)
        elif node["node_key"] in top_node_keys:
            shape.fill.fore_color.rgb = RGBColor(254, 240, 138)
        elif children:
            shape.fill.fore_color.rgb = RGBColor(220, 252, 231)
        else:
            shape.fill.fore_color.rgb = RGBColor(241, 245, 249)

        text_frame = shape.text_frame
        text_frame.clear()
        text_frame.word_wrap = True
        text_frame.margin_top = int(Inches(0.05))
        text_frame.margin_bottom = int(Inches(0.04))
        text_frame.margin_left = int(Inches(0.05))
        text_frame.margin_right = int(Inches(0.05))
        name_paragraph = text_frame.paragraphs[0]
        name_paragraph.text = format_display_name(node["employee_name"])
        name_paragraph.alignment = PP_ALIGN.CENTER
        if name_paragraph.runs:
            name_run = name_paragraph.runs[0]
            name_run.font.bold = True
            name_run.font.size = Pt(12 if depth == 0 else (name_size_manager if node["direct_reports_count"] > 0 else name_size_regular))
            name_run.font.color.rgb = RGBColor(30, 41, 59)

        dept_paragraph = text_frame.add_paragraph()
        dept_paragraph.text = node["department"] or ""
        dept_paragraph.alignment = PP_ALIGN.CENTER
        if dept_paragraph.runs:
            dept_run = dept_paragraph.runs[0]
            dept_run.font.size = Pt(dept_size_manager if node["direct_reports_count"] > 0 else dept_size_regular)
            dept_run.font.color.rgb = RGBColor(71, 85, 105)

        if node["direct_reports_count"] > 0:
            metrics_paragraph = text_frame.add_paragraph()
            metrics_paragraph.text = f"כפיפים ישירים: {node['direct_reports_count']} | סה\"כ בענף: {node['subtree_employee_count']}"
            metrics_paragraph.alignment = PP_ALIGN.CENTER
            if metrics_paragraph.runs:
                metrics_run = metrics_paragraph.runs[0]
                metrics_run.font.size = Pt(metrics_size)
                metrics_run.font.bold = True
                metrics_run.font.color.rgb = RGBColor(51, 65, 85)

        shape_bounds[node["node_key"]] = {"x": x, "y": y, "w": box_width, "h": box_height}
        node_shapes[node["node_key"]] = shape

    for node in sorted_nodes:
        parent_bounds = shape_bounds.get(node["node_key"])
        if not parent_bounds:
            continue
        for child in box_children_map.get(node["node_key"], []):
            child_bounds = shape_bounds.get(child["node_key"])
            if not child_bounds:
                continue
            connector = slide.shapes.add_connector(
                MSO_CONNECTOR.STRAIGHT,
                parent_bounds["x"] + parent_bounds["w"] // 2,
                parent_bounds["y"] + parent_bounds["h"],
                child_bounds["x"] + child_bounds["w"] // 2,
                child_bounds["y"],
            )
            connector.line.color.rgb = RGBColor(148, 163, 184)
            connector.line.width = Pt(1.2)
    return slide, node_shapes


def add_department_summary_slide(prs, title, department_rows, page_index=1, page_count=1):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    title_box = slide.shapes.add_textbox(Inches(0.45), Inches(0.2), Inches(12.2), Inches(0.55))
    title_frame = title_box.text_frame
    title_frame.clear()
    title_paragraph = title_frame.paragraphs[0]
    title_paragraph.text = title if page_count == 1 else f"{title} ({page_index}/{page_count})"
    title_paragraph.alignment = PP_ALIGN.RIGHT
    if title_paragraph.runs:
        title_run = title_paragraph.runs[0]
        title_run.font.size = Pt(24)
        title_run.font.bold = True

    cols = 2
    rows = 3
    card_width = Inches(6.15)
    card_height = Inches(1.55)
    start_x = Inches(0.45)
    start_y = Inches(1.1)
    gap_x = Inches(0.18)
    gap_y = Inches(0.16)
    for index, department in enumerate(department_rows):
        row = index // cols
        col = index % cols
        x = start_x + col * (card_width + gap_x)
        y = start_y + row * (card_height + gap_y)
        card = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE, x, y, card_width, card_height)
        card.line.color.rgb = RGBColor(15, 118, 110)
        card.line.width = Pt(1)
        card.fill.solid()
        card.fill.fore_color.rgb = RGBColor(236, 253, 245)
        frame = card.text_frame
        frame.clear()
        frame.margin_top = int(Inches(0.05))
        frame.margin_bottom = int(Inches(0.04))
        frame.margin_left = int(Inches(0.06))
        frame.margin_right = int(Inches(0.06))

        p1 = frame.paragraphs[0]
        p1.text = department["department"]
        p1.alignment = PP_ALIGN.CENTER
        if p1.runs:
            p1.runs[0].font.bold = True
            p1.runs[0].font.size = Pt(13)
            p1.runs[0].font.color.rgb = RGBColor(15, 23, 42)
        p2 = frame.add_paragraph()
        p2.text = f"עובדים: {department['employee_count']} | מנהלים: {department['manager_count']}"
        p2.alignment = PP_ALIGN.CENTER
        if p2.runs:
            p2.runs[0].font.size = Pt(10)
            p2.runs[0].font.color.rgb = RGBColor(51, 65, 85)
        p3 = frame.add_paragraph()
        p3.text = "מנהלים מובילים: " + (", ".join(department["manager_names"]) if department["manager_names"] else "ללא")
        p3.alignment = PP_ALIGN.CENTER
        if p3.runs:
            p3.runs[0].font.size = Pt(9)
            p3.runs[0].font.color.rgb = RGBColor(71, 85, 105)
    return slide


def add_manager_summary_slide(prs, title, manager_rows, page_index=1, page_count=1):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    title_box = slide.shapes.add_textbox(Inches(0.45), Inches(0.2), Inches(12.2), Inches(0.55))
    title_frame = title_box.text_frame
    title_frame.clear()
    title_paragraph = title_frame.paragraphs[0]
    title_paragraph.text = title if page_count == 1 else f"{title} ({page_index}/{page_count})"
    title_paragraph.alignment = PP_ALIGN.RIGHT
    if title_paragraph.runs:
        title_run = title_paragraph.runs[0]
        title_run.font.size = Pt(24)
        title_run.font.bold = True

    cols = 2
    card_width = Inches(6.1)
    card_height = Inches(2.0)
    start_x = Inches(0.45)
    start_y = Inches(1.05)
    gap_x = Inches(0.25)
    gap_y = Inches(0.18)
    for index, manager in enumerate(manager_rows):
        row = index // cols
        col = index % cols
        x = start_x + col * (card_width + gap_x)
        y = start_y + row * (card_height + gap_y)
        card = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE, x, y, card_width, card_height)
        card.line.color.rgb = RGBColor(100, 116, 139)
        card.line.width = Pt(1)
        card.fill.solid()
        if manager["subtree_employee_count"] >= 12:
            card.fill.fore_color.rgb = RGBColor(254, 240, 138)
        elif manager["direct_reports_count"] >= 5:
            card.fill.fore_color.rgb = RGBColor(220, 252, 231)
        else:
            card.fill.fore_color.rgb = RGBColor(239, 246, 255)
        frame = card.text_frame
        frame.clear()
        frame.margin_top = int(Inches(0.05))
        frame.margin_bottom = int(Inches(0.04))
        frame.margin_left = int(Inches(0.06))
        frame.margin_right = int(Inches(0.06))

        p1 = frame.paragraphs[0]
        p1.text = manager["manager_name"]
        p1.alignment = PP_ALIGN.CENTER
        if p1.runs:
            p1.runs[0].font.bold = True
            p1.runs[0].font.size = Pt(13)
            p1.runs[0].font.color.rgb = RGBColor(30, 41, 59)
        p2 = frame.add_paragraph()
        p2.text = f"{manager['department']} | כפיפים ישירים: {manager['direct_reports_count']} | סה\"כ בענף: {manager['subtree_employee_count']}"
        p2.alignment = PP_ALIGN.CENTER
        if p2.runs:
            p2.runs[0].font.size = Pt(9.5)
            p2.runs[0].font.color.rgb = RGBColor(51, 65, 85)

        employee_names = manager["employee_names"]
        if employee_names:
            p3 = frame.add_paragraph()
            p3.text = "עובדים:"
            p3.alignment = PP_ALIGN.CENTER
            if p3.runs:
                p3.runs[0].font.size = Pt(9)
                p3.runs[0].font.bold = True
                p3.runs[0].font.color.rgb = RGBColor(30, 41, 59)

            names_per_line = 1 if len(employee_names) <= 4 else (2 if len(employee_names) <= 8 else 3)
            name_lines = []
            for start in range(0, len(employee_names), names_per_line):
                group = employee_names[start:start + names_per_line]
                name_lines.append(" • ".join(group))

            if len(employee_names) <= 4:
                name_font_size = 10.5
            elif len(employee_names) <= 8:
                name_font_size = 9.5
            else:
                name_font_size = 8.5

            for line in name_lines:
                p_name = frame.add_paragraph()
                p_name.text = line
                p_name.alignment = PP_ALIGN.CENTER
                if p_name.runs:
                    p_name.runs[0].font.size = Pt(name_font_size)
                    p_name.runs[0].font.color.rgb = RGBColor(51, 65, 85)
    return slide


def write_org_hierarchy_pptx(output_path, summary_rows, tree_rows, exception_rows, stats):
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    manager_count = sum(1 for row in summary_rows if row["is_manager"] == "כן")
    department_rows = build_department_summary_rows(summary_rows)
    root_nodes = [
        {
            "root_key": (row["employee_number"], row["id_number"], row["employee_name"]),
            "title": format_display_name(row["employee_name"]),
        }
        for row in summary_rows
        if row["depth"] == 0
    ]
    manager_rows = build_manager_summary_rows(summary_rows, root_nodes)
    overview_top_nodes = []
    overview_children_map = defaultdict(list)
    for root_node in root_nodes:
        nodes, children_map = build_org_root_nodes(summary_rows, root_node["root_key"])
        actual_root = next((row for row in nodes if row["node_key"] == root_node["root_key"] and row["depth"] == 0), nodes[0] if nodes else None)
        if actual_root is None:
            continue
        top_nodes = [actual_root] + [
            node for node in nodes if node["node_key"] != actual_root["node_key"] and node.get("parent_key") is None
        ]
        overview_top_nodes.extend(top_nodes)
        for parent_key, children in children_map.items():
            overview_children_map[parent_key].extend(children)
    add_pptx_slide(
        prs,
        "סיכום מבנה ארגוני",
        [
            {"text": line, "size": 20}
            for line in (
                build_org_summary_slide_lines(summary_rows, stats)
                + [f"סה\"כ מחלקות: {len(department_rows)}"]
            )
        ],
    )
    add_org_chart_slide(
        prs,
        "מפת מנהלים ומחלקות",
        overview_top_nodes,
        overview_children_map,
        "תרשים ניהולי של מבנה הדיווח בין מנהלים. עובדים ללא צוות אינם מוצגים בתרשים זה.",
        level_limit=3,
        compact=True,
    )
    add_org_chart_slide(
        prs,
        "מבט-על ארגוני",
        overview_top_nodes,
        overview_children_map,
        (
            f"סקירת כלל הארגון לפי מחלקות ומנהלים | עובדים: {stats['employee_count']} | "
            f"מנהלים: {manager_count} | מחלקות: {len(department_rows)}"
        ),
        level_limit=None,
        compact=True,
    )
    department_page_size = 6
    department_pages = max(1, (len(department_rows) + department_page_size - 1) // department_page_size)
    for page_index in range(department_pages):
        start = page_index * department_page_size
        end = start + department_page_size
        add_department_summary_slide(
            prs,
            "סיכום מחלקות",
            department_rows[start:end],
            page_index=page_index + 1,
            page_count=department_pages,
        )
    manager_page_size = 4
    manager_pages = max(1, (len(manager_rows) + manager_page_size - 1) // manager_page_size)
    for page_index in range(manager_pages):
        start = page_index * manager_page_size
        end = start + manager_page_size
        add_manager_summary_slide(
            prs,
            "סיכום מנהלים וצוותים",
            manager_rows[start:end],
            page_index=page_index + 1,
            page_count=manager_pages,
        )
    add_pptx_slide(
        prs,
        "חריגים",
        [{"text": line, "size": 18} for line in build_org_exceptions_slide_lines(exception_rows)],
    )
    prs.save(output_path)


def run_org_hierarchy_report(input_path, output_path, extension, options=None):
    if extension != "csv":
        raise ValueError("Organizational hierarchy report currently supports CSV input only")
    options = options or {}
    output_type = options.get("output_type", "both").strip() or "both"
    summary_rows, tree_rows, exception_rows, stats = parse_org_hierarchy_csv(input_path)
    root_rows = [row for row in summary_rows if row["depth"] == 0]
    output_file = Path(output_path)
    excel_output_path = output_file.with_name(output_file.stem + ".xlsx")
    pptx_output_path = output_file.with_name(output_file.stem + ".pptx")
    wb = Workbook()
    write_org_hierarchy_summary(wb.active, summary_rows, stats, root_rows)
    write_org_hierarchy_tree(wb.create_sheet(), tree_rows)
    write_org_hierarchy_exceptions(wb.create_sheet(), exception_rows, root_rows)
    if output_type == "excel":
        wb.save(output_file)
        return
    if output_type == "powerpoint":
        write_org_hierarchy_pptx(str(output_file), summary_rows, tree_rows, exception_rows, stats)
        return

    wb.save(excel_output_path)
    write_org_hierarchy_pptx(str(pptx_output_path), summary_rows, tree_rows, exception_rows, stats)
    with ZipFile(output_file, "w", compression=ZIP_DEFLATED) as bundle:
        bundle.write(excel_output_path, arcname=excel_output_path.name)
        bundle.write(pptx_output_path, arcname=pptx_output_path.name)
    for temp_path in (excel_output_path, pptx_output_path):
        try:
            temp_path.unlink()
        except OSError:
            pass


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


def yes_no(value):
    return "כן" if value else "לא"


def parse_excel_date(workbook, value):
    if value in ("", None):
        return None
    if isinstance(value, (int, float)):
        try:
            return xlrd.xldate_as_datetime(float(value), workbook.datemode).date()
        except (ValueError, OverflowError):
            return None
    return None


def parse_matan_manual_corrections(input_path):
    workbook = xlrd.open_workbook(input_path)
    employee_rows = []
    daily_rows = []

    for sheet in workbook.sheets():
        employee_name = str(get_sheet_cell(sheet, 5, 5, "")).strip() or sheet.name
        department = str(get_sheet_cell(sheet, 5, 23, "")).strip()
        payroll_number = str(get_sheet_cell(sheet, 5, 50, "")).strip()
        id_number = str(get_sheet_cell(sheet, 7, 5, "")).strip()

        raw_corrections = 0
        entry_corrections = 0
        exit_corrections = 0
        days_with_corrections = 0
        capped_corrections = 0
        work_days = 0
        month_days = 0

        for row_index in range(12, sheet.nrows):
            entry_value = str(get_sheet_cell(sheet, row_index, 13, "")).strip()
            exit_value = str(get_sheet_cell(sheet, row_index, 18, "")).strip()
            event_value = str(get_sheet_cell(sheet, row_index, 22, "")).strip()
            total_hours = str(get_sheet_cell(sheet, row_index, 30, "")).strip()
            day_date = parse_excel_date(workbook, get_sheet_cell(sheet, row_index, 0, ""))

            if not any([day_date, entry_value, exit_value, event_value, total_hours]):
                continue

            if day_date and not month_days:
                month_days = calendar.monthrange(day_date.year, day_date.month)[1]

            if any([entry_value, exit_value, event_value, total_hours]):
                work_days += 1

            entry_corrected = "*" in entry_value
            exit_corrected = "*" in exit_value
            raw_daily = int(entry_corrected) + int(exit_corrected)
            capped_daily = min(raw_daily, 2)

            if raw_daily:
                days_with_corrections += 1
                raw_corrections += raw_daily
                entry_corrections += int(entry_corrected)
                exit_corrections += int(exit_corrected)
                capped_corrections += capped_daily

            daily_rows.append(
                {
                    "employee_name": employee_name,
                    "payroll_number": payroll_number,
                    "id_number": id_number,
                    "department": department,
                    "date": day_date.isoformat() if day_date else "",
                    "entry_corrected": entry_corrected,
                    "exit_corrected": exit_corrected,
                    "raw_daily_corrections": raw_daily,
                    "capped_daily_corrections": capped_daily,
                }
            )

        employee_rows.append(
            {
                "employee_name": employee_name,
                "payroll_number": payroll_number,
                "id_number": id_number,
                "department": department,
                "raw_correction_count": raw_corrections,
                "entry_correction_count": entry_corrections,
                "exit_correction_count": exit_corrections,
                "days_with_corrections": days_with_corrections,
                "capped_correction_count": capped_corrections,
                "average_per_calendar_day": (capped_corrections / month_days) if month_days else 0.0,
                "average_per_work_day": (capped_corrections / work_days) if work_days else 0.0,
            }
        )

    return employee_rows, daily_rows


def apply_matan_manual_corrections_filters(rows, options):
    min_corrections = parse_float_or_none(options.get("min_corrections", ""))
    max_corrections = parse_float_or_none(options.get("max_corrections", ""))
    filtered = []
    for row in rows:
        correction_count = row["raw_correction_count"]
        if min_corrections is not None and correction_count < min_corrections:
            continue
        if max_corrections is not None and correction_count > max_corrections:
            continue
        filtered.append(row)
    return filtered


def write_matan_corrections_summary(ws, employee_rows, filters_used):
    ws.title = safe_sheet_title("סיכום תיקונים", "Corrections Summary")
    ws.sheet_view.rightToLeft = True
    ws.sheet_view.showGridLines = False
    ws.freeze_panes = "A8"

    ws["A1"] = "דוח תיקונים ידניים - מתן"
    ws["A1"].font = Font(bold=True, size=18, color="0F172A")
    ws["A1"].fill = PatternFill(fill_type="solid", fgColor="BFDBFE")

    metrics = [
        ("עובדים בתוצאה", len(employee_rows), "DBEAFE"),
        ("סה\"כ תיקונים גולמיים", sum(row["raw_correction_count"] for row in employee_rows), "FEE2E2"),
        ("סה\"כ תיקונים לאחר תקרה", sum(row["capped_correction_count"] for row in employee_rows), "DCFCE7"),
        ("סה\"כ ימים עם תיקונים", sum(row["days_with_corrections"] for row in employee_rows), "FEF3C7"),
    ]
    for idx, (label, value, fill_color) in enumerate(metrics, start=3):
        label_cell = ws.cell(row=idx, column=1, value=label)
        value_cell = ws.cell(row=idx, column=2, value=value)
        label_cell.font = Font(bold=True, color="334155")
        value_cell.font = Font(bold=True, color="0F172A")
        label_cell.fill = PatternFill(fill_type="solid", fgColor=fill_color)
        value_cell.fill = PatternFill(fill_type="solid", fgColor=fill_color)

    ws["D3"] = "סינונים"
    ws["D3"].font = Font(bold=True)
    for idx, (label, value) in enumerate(filters_used.items(), start=4):
        ws.cell(row=idx, column=4, value=label).font = Font(bold=True)
        ws.cell(row=idx, column=5, value=value or "ללא")

    header_row = 7
    headers = [
        "שם עובד",
        "מספר שכר",
        "תעודת זהות",
        "מחלקה",
        "כמות תיקונים גולמית",
        "תיקוני כניסה",
        "תיקוני יציאה",
        "ימים עם תיקונים",
        "כמות תיקונים לאחר תקרה",
        "ממוצע ליום קלנדרי",
        "ממוצע ליום עבודה",
    ]
    for col, header in enumerate(headers, start=1):
        cell = ws.cell(row=header_row, column=col, value=header)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill(fill_type="solid", fgColor="1E3A8A")

    sorted_rows = sorted(employee_rows, key=lambda row: (-row["raw_correction_count"], row["employee_name"]))
    for row_idx, row in enumerate(sorted_rows, start=header_row + 1):
        values = [
            row["employee_name"],
            row["payroll_number"],
            row["id_number"],
            row["department"],
            row["raw_correction_count"],
            row["entry_correction_count"],
            row["exit_correction_count"],
            row["days_with_corrections"],
            row["capped_correction_count"],
            row["average_per_calendar_day"],
            row["average_per_work_day"],
        ]
        for col, value in enumerate(values, start=1):
            ws.cell(row=row_idx, column=col, value=value)
            if row_idx % 2 == 0:
                ws.cell(row=row_idx, column=col).fill = PatternFill(fill_type="solid", fgColor="F8FAFC")
        ws.cell(row=row_idx, column=5).fill = PatternFill(fill_type="solid", fgColor="FEE2E2")
        ws.cell(row=row_idx, column=9).fill = PatternFill(fill_type="solid", fgColor="DCFCE7")
        ws.cell(row=row_idx, column=5).font = Font(bold=True, color="991B1B")
        ws.cell(row=row_idx, column=9).font = Font(bold=True, color="166534")
        ws.cell(row=row_idx, column=10).number_format = "0.00"
        ws.cell(row=row_idx, column=11).number_format = "0.00"

    widths = [24, 16, 16, 24, 18, 18, 18, 18, 20, 20, 18]
    for col, width in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(col)].width = width


def write_matan_corrections_daily(ws, daily_rows, allowed_names):
    ws.title = safe_sheet_title("פירוט יומי", "Daily Corrections")
    ws.sheet_view.rightToLeft = True
    ws.sheet_view.showGridLines = False
    ws.freeze_panes = "A2"

    headers = [
        "עובד",
        "תאריך",
        "כניסה תוקנה",
        "יציאה תוקנה",
        "תיקונים יומיים גולמיים",
        "תיקונים יומיים לאחר תקרה",
    ]
    for col, header in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=col, value=header)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill(fill_type="solid", fgColor="0F766E")

    row_idx = 2
    for row in daily_rows:
        if row["employee_name"] not in allowed_names or row["raw_daily_corrections"] <= 0:
            continue
        values = [
            row["employee_name"],
            row["date"],
            yes_no(row["entry_corrected"]),
            yes_no(row["exit_corrected"]),
            row["raw_daily_corrections"],
            row["capped_daily_corrections"],
        ]
        for col, value in enumerate(values, start=1):
            ws.cell(row=row_idx, column=col, value=value)
            if row_idx % 2 == 0:
                ws.cell(row=row_idx, column=col).fill = PatternFill(fill_type="solid", fgColor="ECFDF5")
        ws.cell(row=row_idx, column=5).fill = PatternFill(fill_type="solid", fgColor="FEF2F2")
        ws.cell(row=row_idx, column=6).fill = PatternFill(fill_type="solid", fgColor="ECFDF5")
        row_idx += 1

    widths = [24, 14, 16, 16, 20, 22]
    for col, width in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(col)].width = width


def run_matan_manual_corrections(input_path, output_path, extension, options=None):
    if extension != "xls":
        raise ValueError("Matan manual-corrections tool currently supports XLS export only")
    options = options or {}
    employee_rows, daily_rows = parse_matan_manual_corrections(input_path)
    filtered_rows = apply_matan_manual_corrections_filters(employee_rows, options)
    allowed_names = {row["employee_name"] for row in filtered_rows}
    wb = Workbook()
    write_matan_corrections_summary(
        wb.active,
        filtered_rows,
        {
            "מינימום תיקונים": options.get("min_corrections", ""),
            "מקסימום תיקונים": options.get("max_corrections", ""),
        },
    )
    write_matan_corrections_daily(wb.create_sheet(), daily_rows, allowed_names)
    wb.save(output_path)


def has_rimon_work_activity(entry_value, exit_value, total_hours):
    return any([str(entry_value).strip(), str(exit_value).strip(), parse_hours_value(total_hours)])


def is_rimon_error_text(error_text):
    text = str(error_text).strip()
    if not text:
        return False
    return text not in {"יום חסר"}


def parse_rimon_home_office_report(input_path):
    workbook = xlrd.open_workbook(input_path)
    employee_rows = []
    daily_rows = []

    for sheet in workbook.sheets():
        employee_name = str(get_sheet_cell(sheet, 5, 5, "")).strip() or sheet.name
        department = str(get_sheet_cell(sheet, 5, 21, "")).strip()
        payroll_number = str(get_sheet_cell(sheet, 5, 48, "")).strip()
        id_number = str(get_sheet_cell(sheet, 7, 5, "")).strip()

        grouped_dates = {}
        current_date = None

        for row_index in range(11, sheet.nrows):
            row_date = parse_excel_date(workbook, get_sheet_cell(sheet, row_index, 0, ""))
            if row_date:
                current_date = row_date
            if current_date is None:
                continue

            entry_value = str(get_sheet_cell(sheet, row_index, 8, "")).strip()
            exit_value = str(get_sheet_cell(sheet, row_index, 12, "")).strip()
            event_value = str(get_sheet_cell(sheet, row_index, 17, "")).strip()
            total_hours = str(get_sheet_cell(sheet, row_index, 20, "")).strip()
            standard_hours = str(get_sheet_cell(sheet, row_index, 25, "")).strip()
            missing_hours = str(get_sheet_cell(sheet, row_index, 30, "")).strip()
            error_text = str(get_sheet_cell(sheet, row_index, 51, "")).strip()

            if not any([row_date, entry_value, exit_value, event_value, total_hours, standard_hours, missing_hours, error_text]):
                continue

            day_key = current_date.isoformat()
            if day_key not in grouped_dates:
                grouped_dates[day_key] = {
                    "employee_name": employee_name,
                    "payroll_number": payroll_number,
                    "id_number": id_number,
                    "department": department,
                    "date": day_key,
                    "home_office": False,
                    "work_activity": False,
                    "missing_absence": False,
                    "error": False,
                    "events": [],
                    "total_hours_value": None,
                    "standard_hours_value": None,
                    "missing_hours_value": None,
                    "errors": [],
                }

            grouped = grouped_dates[day_key]
            grouped["home_office"] = grouped["home_office"] or event_value == "עבודה מהבית"
            grouped["work_activity"] = grouped["work_activity"] or has_rimon_work_activity(entry_value, exit_value, total_hours)
            grouped["error"] = grouped["error"] or is_rimon_error_text(error_text)

            parsed_total = parse_hours_value(total_hours)
            if parsed_total is not None:
                grouped["total_hours_value"] = max(grouped["total_hours_value"] or 0.0, parsed_total)

            parsed_standard = parse_hours_value(standard_hours)
            if parsed_standard is not None:
                grouped["standard_hours_value"] = max(grouped["standard_hours_value"] or 0.0, parsed_standard)

            parsed_missing = parse_hours_value(missing_hours)
            if parsed_missing is not None:
                grouped["missing_hours_value"] = max(grouped["missing_hours_value"] or 0.0, parsed_missing)

            if event_value and event_value not in grouped["events"]:
                grouped["events"].append(event_value)
            if error_text and error_text not in grouped["errors"]:
                grouped["errors"].append(error_text)

        office_days = 0
        home_office_days = 0
        missing_absence_days = 0
        error_days = 0

        for day_key in sorted(grouped_dates):
            grouped = grouped_dates[day_key]
            absence_signal = bool(grouped["events"] or grouped["errors"] or (grouped["missing_hours_value"] or 0.0) > 0)
            grouped["missing_absence"] = (not grouped["work_activity"]) and absence_signal
            office_work = grouped["work_activity"] and not grouped["home_office"]

            if grouped["home_office"]:
                home_office_days += 1
            if office_work:
                office_days += 1
            if grouped["missing_absence"]:
                missing_absence_days += 1
            if grouped["error"]:
                error_days += 1

            daily_rows.append(
                {
                    "employee_name": grouped["employee_name"],
                    "date": grouped["date"],
                    "home_office": grouped["home_office"],
                    "office_work": office_work,
                    "missing_absence": grouped["missing_absence"],
                    "error": grouped["error"],
                    "event": " | ".join(grouped["events"]),
                    "total_hours": format_hours(grouped["total_hours_value"]),
                    "standard_hours": format_hours(grouped["standard_hours_value"]),
                    "missing_hours": format_hours(grouped["missing_hours_value"]),
                    "error_text": " | ".join(grouped["errors"]),
                }
            )

        employee_rows.append(
            {
                "employee_name": employee_name,
                "payroll_number": payroll_number,
                "id_number": id_number,
                "department": department,
                "office_work_days": office_days,
                "home_office_days": home_office_days,
                "missing_absence_days": missing_absence_days,
                "error_days": error_days,
                "total_grouped_dates": len(grouped_dates),
            }
        )

    return employee_rows, daily_rows


def write_rimon_home_office_summary(ws, employee_rows):
    ws.title = safe_sheet_title("סיכום רימון", "Rimon Summary")
    ws.sheet_view.rightToLeft = True
    ws.sheet_view.showGridLines = False
    ws.freeze_panes = "A7"

    ws["A1"] = "דוח סיכום עבודה מהבית - רימון"
    ws["A1"].font = Font(bold=True, size=18, color="0F172A")
    ws["A1"].fill = PatternFill(fill_type="solid", fgColor="BFDBFE")

    metrics = [
        ("סה\"כ ימי עבודה", sum(row["office_work_days"] + row["home_office_days"] for row in employee_rows), "DBEAFE"),
        ("סה\"כ ימי עבודה מהבית", sum(row["home_office_days"] for row in employee_rows), "DDD6FE"),
        ("סה\"כ ימי עבודה מהמשרד", sum(row["office_work_days"] for row in employee_rows), "DCFCE7"),
    ]
    for idx, (label, value, fill_color) in enumerate(metrics, start=3):
        label_cell = ws.cell(row=idx, column=1, value=label)
        value_cell = ws.cell(row=idx, column=2, value=value)
        label_cell.font = Font(bold=True, color="334155")
        value_cell.font = Font(bold=True, color="0F172A")
        label_cell.fill = PatternFill(fill_type="solid", fgColor=fill_color)
        value_cell.fill = PatternFill(fill_type="solid", fgColor=fill_color)

    headers = [
        "שם עובד",
        "מספר שכר",
        "תעודת זהות",
        "מחלקה",
        "ימי עבודה מהמשרד",
        "ימי עבודה מהבית",
        "ימי חסר/היעדרות",
        "ימי שגיאה",
        "סה\"כ ימי עבודה שזוהו",
    ]
    header_row = 6
    for col, header in enumerate(headers, start=1):
        cell = ws.cell(row=header_row, column=col, value=header)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill(fill_type="solid", fgColor="1E3A8A")

    sorted_rows = sorted(employee_rows, key=lambda row: (row["employee_name"], row["payroll_number"]))
    for row_idx, row in enumerate(sorted_rows, start=header_row + 1):
        values = [
            row["employee_name"],
            row["payroll_number"],
            row["id_number"],
            row["department"],
            row["office_work_days"],
            row["home_office_days"],
            row["missing_absence_days"],
            row["error_days"],
            row["office_work_days"] + row["home_office_days"],
        ]
        for col, value in enumerate(values, start=1):
            ws.cell(row=row_idx, column=col, value=value)
            if row_idx % 2 == 0:
                ws.cell(row=row_idx, column=col).fill = PatternFill(fill_type="solid", fgColor="F8FAFC")

    widths = [24, 16, 16, 24, 18, 18, 18, 14, 28]
    for col, width in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(col)].width = width


def write_rimon_home_office_daily(ws, daily_rows):
    ws.title = safe_sheet_title("פירוט יומי", "Daily Breakdown")
    ws.sheet_view.rightToLeft = True
    ws.sheet_view.showGridLines = False
    ws.freeze_panes = "A2"

    headers = [
        "עובד",
        "תאריך",
        "עבודה מהבית",
        "עבודה מהמשרד",
        "חסר/היעדרות",
        "שגיאה",
        "אירוע",
        "סה\"כ שעות",
        "שעות תקן",
        "שעות חסר",
        "פירוט שגיאה",
    ]
    for col, header in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=col, value=header)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill(fill_type="solid", fgColor="0F766E")

    sorted_rows = sorted(daily_rows, key=lambda row: (row["employee_name"], row["date"]))
    for row_idx, row in enumerate(sorted_rows, start=2):
        values = [
            row["employee_name"],
            row["date"],
            yes_no(row["home_office"]),
            yes_no(row["office_work"]),
            yes_no(row["missing_absence"]),
            yes_no(row["error"]),
            row["event"],
            row["total_hours"],
            row["standard_hours"],
            row["missing_hours"],
            row["error_text"],
        ]
        for col, value in enumerate(values, start=1):
            ws.cell(row=row_idx, column=col, value=value)
            if row_idx % 2 == 0:
                ws.cell(row=row_idx, column=col).fill = PatternFill(fill_type="solid", fgColor="ECFDF5")

    widths = [24, 14, 14, 14, 18, 10, 20, 14, 14, 14, 22]
    for col, width in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(col)].width = width


def run_rimon_home_office_summary(input_path, output_path, extension, options=None):
    if extension != "xls":
        raise ValueError("Rimon home-office summary currently supports original XLS exports only")
    employee_rows, daily_rows = parse_rimon_home_office_report(input_path)
    wb = Workbook()
    write_rimon_home_office_summary(wb.active, employee_rows)
    write_rimon_home_office_daily(wb.create_sheet(), daily_rows)
    wb.save(output_path)


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


FLOW_LANGUAGES = {"en", "he"}

FLOW_TEXTS = {
    "en": {
        "topbar_greeting": "Hello, ",
        "logout": "Logout",
        "login_page_title": "Login",
        "login_error": "Incorrect username or password",
        "login_subtitle": "Attendance and payroll tools portal",
        "login_username": "Username",
        "login_password": "Password",
        "login_submit": "Log In",
        "dashboard_page_title": "Your Tools",
        "dashboard_empty": "No tools are available for your account yet",
        "dashboard_greeting": "Welcome, ",
        "dashboard_intro": "Your available tools:",
        "run_access_denied": "You do not have access to this tool",
        "run_extra_file_type_error": "The extra file type is not supported",
        "run_missing_extra_file_error": "A required extra file is missing",
        "run_unexpected_error_prefix": "Unexpected processing error: ",
        "back_arrow": "&#8592;",
        "scripts": {
            "nikuy": {
                "name": "Attendance Report Cleanup",
                "desc": "Clean a detailed monthly attendance report by removing asterisks and question marks",
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
            },
            "flamingo_payroll": {
                "name": "Fixed Hourly-Rate Payroll Summary",
                "desc": "Generate a clear monthly payroll summary based on a detailed report and a fixed hourly rate",
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
            },
            "matan_missing": {
                "name": "Missing Hours Report",
                "desc": "Filter employees by their monthly missing-hours totals with a clear summary ready for follow-up",
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
                    {"label": "Minimum missing hours", "placeholder": "For example 4"},
                    {"label": "Maximum missing hours", "placeholder": "For example 8"},
                ],
            },
        },
    },
    "he": {
        "topbar_greeting": "שלום, ",
        "logout": "התנתקות",
        "login_page_title": "כניסה",
        "login_error": "שם המשתמש או הסיסמה שגויים",
        "login_subtitle": "פורטל כלי נוכחות ושכר",
        "login_username": "שם משתמש",
        "login_password": "סיסמה",
        "login_submit": "כניסה",
        "dashboard_page_title": "הכלים שלך",
        "dashboard_empty": "עדיין אין כלים זמינים לחשבון שלך",
        "dashboard_greeting": "ברוך/ה הבא/ה, ",
        "dashboard_intro": "",
        "run_access_denied": "אין לך גישה לכלי הזה",
        "run_extra_file_type_error": "סוג הקובץ הנוסף אינו נתמך",
        "run_missing_extra_file_error": "חסר קובץ נוסף נדרש",
        "run_unexpected_error_prefix": "שגיאת עיבוד לא צפויה: ",
        "back_arrow": "&#8594;",
        "scripts": {
            "nikuy": {
                "name": "ניקוי דוח נוכחות",
                "desc": "ניקוי סימני כוכביות וסימני שאלה מדוח נוכחות מפורט חודשי",
                "help_label": "דרישות לקובץ",
                "help_title": "מה צריך להעלות?",
                "help_intro": "יש להעלות דוח נוכחות מפורט חודשי הכולל סימני * או ?.",
                "help_items": ["המערכת תחזיר את אותו הדוח לאחר ניקוי הסימנים"],
                "help_note": "הפלט מתקבל מוכן להמשך עבודה",
                "success_title": "הקובץ מוכן",
                "success_action": "הורדת הקובץ הנקי",
                "retry_action": "עיבוד קובץ נוסף",
                "submit_label": "הפעל",
                "back_label": "חזרה לכלים",
                "empty_error": "לא נבחר קובץ",
                "unsupported_error": "סוג הקובץ אינו נתמך",
                "invalid_error": "הקובץ שהועלה אינו קובץ אקסל תקין",
                "empty_file_error": "הקובץ שהועלה ריק",
                "too_large_error": "הקובץ שהועלה גדול מדי",
                "processing_error": "לא ניתן היה לעבד את הקובץ שהועלה",
                "processing_title": "הקובץ נמצא בעיבוד",
                "processing_note": "הכנת הדוח הנקי עשויה להימשך כמה דקות. נא להשאיר את הדף פתוח.",
                "file_picker_label": "בחירת קובץ",
            },
            "flamingo_payroll": {
                "name": "סיכום שכר לפי תעריף שעתי קבוע",
                "desc": "הפקת סיכום שכר לפי דוח מפורט חודשי בצורה ברורה ומוכנה לבדיקה",
                "help_label": "דרישות לקובץ",
                "help_title": "מה צריך להעלות?",
                "help_intro": "יש להעלות דוח מפורט חודשי הכולל שדה של תעריף שעתי קבוע.",
                "help_items": ["המערכת תפיק דוח שכר לעובדים", "כולל חישוב שכר וסיכומים מתקדמים"],
                "help_note": "מיועד בעיקר לחברות כוח אדם",
                "success_title": "קובץ השכר מוכן",
                "success_action": "הורדת סיכום השכר",
                "retry_action": "עיבוד קובץ שכר נוסף",
                "submit_label": "יצירת סיכום שכר",
                "back_label": "חזרה לכלים",
                "empty_error": "לא נבחר קובץ",
                "unsupported_error": "יש להעלות את ייצוא ה-XLS המקורי של פלמינגו",
                "invalid_error": "הקובץ שהועלה אינו קובץ אקסל תקין",
                "empty_file_error": "הקובץ שהועלה ריק",
                "too_large_error": "הקובץ שהועלה גדול מדי",
                "processing_error": "לא ניתן היה ליצור סיכום שכר מהקובץ הזה",
                "processing_title": "סיכום השכר בהכנה",
                "processing_note": "המערכת מחשבת את שעות התשלום והשכר לכל העובדים. זה עשוי להימשך כמה דקות.",
                "file_picker_label": "בחירת קובץ דוח מפורט חודשי",
            },
            "matan_missing": {
                "name": "דוח שעות חסר",
                "desc": "אפשרות לסינון עובדים לפי כמות שעות החוסר החודשיות שלהם כולל סיכום ברור ונוח לטיפול",
                "help_label": "דרישות לקובץ",
                "help_title": "מה צריך להעלות?",
                "help_intro": "יש להעלות דוח חודשי מרוכז של שעות תקן מול שעות החוסר.",
                "help_items": ["המערכת מנתחת את הנתונים ומציגה רק את העובדים שעומדים בתנאי הסינון שהוגדרו", "יותר מ־X שעות חוסר", "פחות מ־Y שעות חוסר", "או שילוב של שני התנאים ביחס לתקן החודשי שלהם"],
                "help_note": "הפלט מאפשר מיקוד מהיר בעובדים הרלוונטיים לטיפול",
                "success_title": "דוח שעות החסר מוכן",
                "success_action": "הורדת הדוח",
                "retry_action": "עיבוד קובץ נוסף",
                "submit_label": "יצירת דוח",
                "back_label": "חזרה לכלים",
                "empty_error": "לא נבחר קובץ",
                "unsupported_error": "יש להעלות את דוח שעות החסר המקורי מסוג XLS",
                "invalid_error": "הקובץ שהועלה אינו קובץ אקסל תקין",
                "empty_file_error": "הקובץ שהועלה ריק",
                "too_large_error": "הקובץ שהועלה גדול מדי",
                "processing_error": "לא ניתן היה ליצור את דוח שעות החסר מהקובץ הזה",
                "processing_title": "הדוח בהכנה",
                "processing_note": "המערכת מסננת את דוח שעות החסר. זה עשוי להימשך כמה דקות.",
                "file_picker_label": "בחירת דוח שעות חסר",
                "filter_fields": [
                    {"label": "מינימום שעות חסר", "placeholder": "לדוגמה 4"},
                    {"label": "מקסימום שעות חסר", "placeholder": "לדוגמה 8"},
                ],
            },
        },
    },
}


def get_flow_lang():
    requested = request.args.get("lang")
    if requested in FLOW_LANGUAGES:
        session["flow_lang"] = requested
    saved = session.get("flow_lang", "he")
    return saved if saved in FLOW_LANGUAGES else "he"


def get_flow_dir(lang):
    return "rtl" if lang == "he" else "ltr"


def get_flow_text(lang):
    return FLOW_TEXTS.get(lang, FLOW_TEXTS["en"])


def get_localized_script(script, lang):
    localized = dict(script)
    script_texts = get_flow_text(lang)["scripts"].get(script["id"], {})
    localized.update({k: v for k, v in script_texts.items() if k not in {"filter_fields", "extra_uploads"}})
    if script.get("filter_fields"):
        localized_fields = []
        field_overrides = script_texts.get("filter_fields", [])
        for index, field in enumerate(script.get("filter_fields", [])):
            merged = dict(field)
            if index < len(field_overrides):
                merged.update(field_overrides[index])
            localized_fields.append(merged)
        localized["filter_fields"] = localized_fields
    if script.get("extra_uploads"):
        localized_uploads = []
        upload_overrides = script_texts.get("extra_uploads", [])
        for index, upload in enumerate(script.get("extra_uploads", [])):
            merged = dict(upload)
            if index < len(upload_overrides):
                merged.update(upload_overrides[index])
            localized_uploads.append(merged)
        localized["extra_uploads"] = localized_uploads
    return localized


def build_lang_switch(lang):
    current_path = request.path
    buttons = []
    for code, label in (("en", "EN"), ("he", "HE")):
        cls = "lang-btn active" if code == lang else "lang-btn"
        buttons.append('<a href="' + current_path + '?lang=' + code + '" class="' + cls + '">' + label + "</a>")
    return '<div class="lang-switch">' + "".join(buttons) + "</div>"

SCRIPTS = {
    "nikuy": {
        "id": "nikuy",
        "name": "ניקוי דוח נוכחות",
        "desc": "ניקוי סימני כוכביות וסימני שאלה מדוח נוכחות מפורט חודשי",
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
    "name": "סיכום שכר לפי תעריף שעתי קבוע",
    "desc": "הפקת סיכום שכר לפי דוח מפורט חודשי בצורה ברורה ומוכנה לבדיקה",
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
    "name": "דוח שעות חסר",
    "desc": "אפשרות לסינון עובדים לפי כמות שעות החוסר החודשיות שלהם כולל סיכום ברור ונוח לטיפול",
    "accept": ".xls",
    "icon": "📊",
}

SCRIPTS["matan_manual_corrections"] = {
    "id": "matan_manual_corrections",
    "name": "דוח תיקונים ידניים",
    "desc": "איתור וסיכום של תיקוני נוכחות ידניים מתוך הדוח, כולל ספירה ותצוגה נוחה לבדיקה",
    "help_label": "דרישות לקובץ",
    "help_title": "מה צריך להעלות?",
    "help_intro": "יש להעלות דוח מפורט חודשי הכולל תיקוני כניסה ויציאה ידניים.",
    "help_items": ["המערכת מזהה תיקונים ידניים מתוך הדוח", "סופרת את התיקונים לכל עובד", "ומציגה סיכום ברור ונוח לבדיקה"],
    "help_note": "מיועד למצבים שבהם רוצים לעקוב אחר תיקונים ידניים בדיווחי הנוכחות",
    "accept": ".xls",
    "icon": "📝",
}

SCRIPTS["rimon_home_office_summary"] = {
    "id": "rimon_home_office_summary",
    "name": "סיכום עבודה מהבית והמשרד",
    "desc": "סיכום ימי עבודה מהבית, ימי עבודה מהמשרד, היעדרויות ושגיאות מתוך דוח הנוכחות",
    "help_label": "דרישות לקובץ",
    "help_title": "מה צריך להעלות?",
    "help_intro": "יש להעלות דוח מפורט חודשי.",
    "help_items": ["המערכת מזהה ימי עבודה מהבית", "ימי עבודה מהמשרד", "היעדרויות", "ושגיאות בדיווח"],
    "help_note": "הפלט מחזיר סיכום ברור לפי עובד",
    "accept": ".xls",
    "icon": "🏠",
}

SCRIPTS["org_hierarchy_report"] = {
    "id": "org_hierarchy_report",
    "name": "תרשים מבנה ארגוני",
    "desc": "הפקת תרשים מבנה ארגוני ודוחות סיכום לפי מנהלים, מחלקות ומבנה הדיווח בארגון, כולל פלט אקסל ו-PowerPoint",
    "help_label": "דרישות לקובץ",
    "help_title": "מה צריך לכלול בקובץ?",
    "help_intro": "יש להעלות קובץ מבנה ארגוני בפורמט CSV או Excel. כדי שהמערכת תפיק את הדוח, הקובץ חייב לכלול לפחות:",
    "help_items": ["שם עובד", "מנהל ישיר", "מחלקה", "ואחד מאמצעי הזיהוי הבאים: מספר עובד, תעודת זהות או דרכון"],
    "help_note": "שדות נוספים כמו סימון מנהל, אימייל או תפקיד יכולים לשפר את הפלט.",
    "accept": ".csv",
    "icon": "🌳",
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

SCRIPT_REGISTRY["matan_manual_corrections"] = {
    **SCRIPTS["matan_manual_corrections"],
    "processor": run_matan_manual_corrections,
    "output_suffix": "matan_manual_corrections",
    "success_title": "דוח התיקונים מוכן",
    "success_action": "הורדת הדוח",
    "retry_action": "עיבוד קובץ נוסף",
    "submit_label": "יצירת דוח תיקונים",
    "back_label": "חזרה לכלים",
    "empty_error": "לא נבחר קובץ",
    "unsupported_error": "יש להעלות את דוח הנוכחות החודשי המפורט המקורי מסוג XLS",
    "invalid_error": "הקובץ שהועלה אינו קובץ אקסל תקין",
    "empty_file_error": "הקובץ שהועלה ריק",
    "too_large_error": "הקובץ שהועלה גדול מדי",
    "processing_error": "לא ניתן היה להפיק את דוח התיקונים מהקובץ הזה",
    "processing_title": "דוח התיקונים בהכנה",
    "processing_note": "המערכת סופרת תיקוני נוכחות ידניים לכל עובד. הפעולה עשויה להימשך כמה דקות.",
    "file_picker_label": "בחירת דוח נוכחות חודשי",
    "filter_fields": [
        {"name": "min_corrections", "label": "מינימום תיקונים", "placeholder": "לדוגמה 4"},
        {"name": "max_corrections", "label": "מקסימום תיקונים", "placeholder": "לדוגמה 12"},
    ],
}

SCRIPT_REGISTRY["rimon_home_office_summary"] = {
    **SCRIPTS["rimon_home_office_summary"],
    "processor": run_rimon_home_office_summary,
    "output_suffix": "rimon_home_office_summary",
    "success_title": "דוח הסיכום מוכן",
    "success_action": "הורדת הדוח",
    "retry_action": "עיבוד קובץ נוסף",
    "submit_label": "יצירת דוח סיכום",
    "back_label": "חזרה לכלים",
    "empty_error": "לא נבחר קובץ",
    "unsupported_error": "יש להעלות את דוח רימון החודשי המפורט המקורי מסוג XLS",
    "invalid_error": "הקובץ שהועלה אינו קובץ אקסל תקין",
    "empty_file_error": "הקובץ שהועלה ריק",
    "too_large_error": "הקובץ שהועלה גדול מדי",
    "processing_error": "לא ניתן היה להפיק את דוח הסיכום מהקובץ הזה",
    "processing_title": "דוח הסיכום בהכנה",
    "processing_note": "המערכת מקבצת תאריכים וסופרת ימי משרד, עבודה מהבית, היעדרות ושגיאות. הפעולה עשויה להימשך כמה דקות.",
    "file_picker_label": "בחירת דוח מפורט חודשי",
}

SCRIPT_REGISTRY["org_hierarchy_report"] = {
    **SCRIPTS["org_hierarchy_report"],
    "processor": run_org_hierarchy_report,
    "output_suffix": "org_hierarchy_report",
    "output_extension": "zip",
    "output_option_name": "output_type",
    "output_extension_map": {"excel": "xlsx", "powerpoint": "pptx", "both": "zip"},
    "success_title": "דוח המבנה הארגוני מוכן",
    "success_action": "הורדת הקובץ",
    "retry_action": "עיבוד קובץ נוסף",
    "submit_label": "יצירת דוח מבנה ארגוני",
    "back_label": "חזרה לכלים",
    "empty_error": "לא נבחר קובץ",
    "unsupported_error": "יש להעלות את קובץ המבנה הארגוני המקורי מסוג CSV",
    "invalid_error": "הקובץ שהועלה אינו קובץ תקין",
    "empty_file_error": "הקובץ שהועלה ריק",
    "too_large_error": "הקובץ שהועלה גדול מדי",
    "processing_error": "לא ניתן היה להפיק את דוח המבנה הארגוני מהקובץ הזה",
    "processing_title": "דוח המבנה הארגוני בהכנה",
    "processing_note": "המערכת בונה דוח אקסל, מצגת או קובץ ZIP לפי הבחירה שלך. הפעולה עשויה להימשך כמה דקות.",
    "file_picker_label": "בחירת קובץ מבנה ארגוני",
    "filter_fields": [
        {
            "name": "output_type",
            "label": "סוג פלט",
            "type": "select",
            "default": "powerpoint",
            "options": [
                {"value": "excel", "label": "אקסל בלבד"},
                {"value": "powerpoint", "label": "PowerPoint בלבד"},
                {"value": "both", "label": "XL+PowerPoint"},
            ],
        },
    ],
}

SCRIPTS = SCRIPT_REGISTRY


def get_script(script_id):
    return SCRIPT_REGISTRY.get(script_id)


def build_output_filename(script, uid, options=None):
    suffix = script.get("output_suffix", "output")
    extension = script.get("output_extension", "xlsx")
    option_name = script.get("output_option_name")
    extension_map = script.get("output_extension_map", {})
    if option_name and options:
        selected = str(options.get(option_name, "")).strip()
        if selected in extension_map:
            extension = extension_map[selected]
    return f"{uid}_{suffix}.{extension}"


def execute_script(script, input_path, output_path, extension, options=None):
    processor = script.get("processor")
    if processor is None:
        raise ValueError("Script processor is not configured")
    processor(input_path, output_path, extension, options)


CSS = """
* { box-sizing: border-box; margin: 0; padding: 0; }
body { font-family: 'Segoe UI', Arial, sans-serif; background: #f0f4ff; min-height: 100vh; direction: inherit; }
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
th { text-align: start; padding: 10px 12px; background: #f8fafc; color: #64748b; font-weight: 600; border-bottom: 1.5px solid #e2e8f0; }
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
.lang-switch { display: inline-flex; align-items: center; gap: 6px; background: rgba(255,255,255,.14); border: 1px solid rgba(255,255,255,.2); border-radius: 999px; padding: 4px; }
.lang-switch.standalone { background: #ffffff; border-color: #dbeafe; box-shadow: 0 4px 16px rgba(37,99,235,.08); margin-bottom: 1rem; }
.lang-btn { display: inline-block; min-width: 38px; padding: 6px 10px; border-radius: 999px; font-size: 12px; font-weight: 700; text-align: center; text-decoration: none; color: #cbd5e1; }
.lang-switch.standalone .lang-btn { color: #64748b; }
.lang-btn.active { background: #ffffff; color: #1e3a8a; }
.lang-switch.standalone .lang-btn.active { background: #2563eb; color: #ffffff; }
@keyframes loadingSlide {
  0% { transform: translateX(0); }
  50% { transform: translateX(120%); }
  100% { transform: translateX(0); }
}
.modal-bg { display: none; position: fixed; inset: 0; background: rgba(0,0,0,.4); z-index: 100; align-items: center; justify-content: center; }
.modal-box { background: white; border-radius: 16px; padding: 1.75rem; width: 320px; }
"""


def get_table_columns(db, table_name):
    if db.is_postgres:
        rows = db.execute(
            "SELECT column_name FROM information_schema.columns WHERE table_schema = current_schema() AND table_name = %s",
            (table_name,),
        ).fetchall()
        return {row["column_name"] for row in rows}
    rows = db.execute(f"PRAGMA table_info({table_name})").fetchall()
    return {row["name"] for row in rows}


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
        db.execute(
            """CREATE TABLE IF NOT EXISTS activity_logs (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id INTEGER,
            username TEXT,
            full_name TEXT,
            event_type TEXT NOT NULL,
            action_label TEXT NOT NULL,
            script_id TEXT,
            script_name TEXT,
            details TEXT,
            created_at TEXT NOT NULL)"""
        )
        db.execute("CREATE INDEX IF NOT EXISTS idx_activity_logs_created_at ON activity_logs(created_at)")
        existing_columns = get_table_columns(db, "users")
        desired_columns = {
            "company_name": "TEXT",
            "company_id": "TEXT",
            "email": "TEXT",
            "phone": "TEXT",
            "join_date": "TEXT",
            "trial_start_date": "TEXT",
            "service_valid_until": "TEXT",
            "billing_mode": "TEXT DEFAULT 'monthly'",
        }
        for column_name, column_sql in desired_columns.items():
            if column_name not in existing_columns:
                db.execute(f"ALTER TABLE users ADD COLUMN {column_name} {column_sql}")
        if not db.execute("SELECT id FROM users WHERE username='admin'").fetchone():
            db.execute(
                "INSERT INTO users(username,password,full_name,is_admin) VALUES (?,?,?,1)",
                ("admin", generate_password_hash("admin123"), "מנהל מערכת"),
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


def esc(value):
    return html.escape(str(value or ""))


def format_ui_datetime(value):
    text = str(value or "").strip()
    if not text:
        return ""
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%dT%H:%M:%S"):
        try:
            return datetime.strptime(text, fmt).strftime("%d/%m/%Y %H:%M")
        except ValueError:
            continue
    return text


def parse_datetime_value(value):
    text = str(value or "").strip()
    if not text:
        return None
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%dT%H:%M:%S"):
        try:
            return datetime.strptime(text, fmt)
        except ValueError:
            continue
    return None


def parse_iso_date(value):
    text = str(value or "").strip()
    if not text:
        return None
    for fmt in ("%Y-%m-%d", "%d/%m/%Y"):
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            continue
    return None


def format_ui_date(value, lang="he"):
    parsed = parse_iso_date(value)
    if not parsed:
        return "לא הוגדר" if lang == "he" else "Not set"
    return parsed.strftime("%d/%m/%Y")


def billing_mode_label(value, lang="he"):
    normalized = str(value or "monthly").strip().lower()
    labels = {
        "monthly": "חודשי" if lang == "he" else "Monthly",
        "yearly_prepaid": "שנתי מראש" if lang == "he" else "Yearly prepaid",
    }
    return labels.get(normalized, normalized or ("לא הוגדר" if lang == "he" else "Not set"))


def log_user_activity(event_type, action_label, script_id="", script_name="", details=""):
    user_id = session.get("user_id")
    if not user_id or session.get("is_admin"):
        return
    try:
        with get_db() as db:
            db.execute(
                """INSERT INTO activity_logs(
                user_id, username, full_name, event_type, action_label, script_id, script_name, details, created_at
                ) VALUES (?,?,?,?,?,?,?,?,?)""",
                (
                    user_id,
                    session.get("username", ""),
                    session.get("name", ""),
                    str(event_type or "")[:80],
                    str(action_label or "")[:120],
                    str(script_id or "")[:80],
                    str(script_name or "")[:120],
                    str(details or "")[:240],
                    datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                ),
            )
            db.commit()
    except Exception:
        pass


def resolve_script_from_output_name(filename):
    logical_name = filename.split("_", 1)[-1] if "_" in filename else filename
    for script in SCRIPT_REGISTRY.values():
        suffix = script.get("output_suffix", "")
        if suffix and logical_name.startswith(suffix + "."):
            return script
    return None


def get_account_status(user_row):
    today = date.today()
    trial_start = parse_iso_date(user_row["trial_start_date"])
    valid_until = parse_iso_date(user_row["service_valid_until"])

    if valid_until:
        if valid_until >= today:
            return {
                "status_key": "active",
                "status_label_he": "לקוח פעיל",
                "status_label_en": "Active customer",
                "renewal_date": valid_until,
                "days_remaining": None,
            }
        return {
            "status_key": "expired",
            "status_label_he": "שירות שפג תוקפו",
            "status_label_en": "Expired service",
            "renewal_date": valid_until,
            "days_remaining": None,
        }

    if trial_start:
        days_remaining = max(0, 30 - (today - trial_start).days)
        return {
            "status_key": "trial",
            "status_label_he": "ניסיון ל-30 יום",
            "status_label_en": "30-day trial",
            "renewal_date": None,
            "days_remaining": days_remaining,
        }

    return {
        "status_key": "unknown",
        "status_label_he": "סטטוס לא הוגדר",
        "status_label_en": "Status not set",
        "renewal_date": None,
        "days_remaining": None,
    }


def render(title, body, nav=True, lang="en", topbar_greeting="Hello, ", logout_label="Logout", show_lang_switch=False):
    direction = get_flow_dir(lang)
    topbar = ""
    if nav:
        name = session.get("name", "")
        lang_switch = build_lang_switch(lang) if show_lang_switch else ""
        topbar = (
            '<div class="topbar">'
            "<h1>&#9201; Scriptly</h1>"
            '<div style="display:flex;gap:16px;align-items:center">'
            + lang_switch
            + '<span style="font-size:13px;color:#93c5fd">' + topbar_greeting + name + "</span>"
            + '<a href="/logout">' + logout_label + '</a>'
            "</div></div>"
        )
    wrap_cls = "wrap" if nav else "login-wrap"
    return (
        '<!DOCTYPE html><html dir="' + direction + '" lang="' + lang + '">'
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

    lang = get_flow_lang()
    text = get_flow_text(lang)
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
        error = '<div class="flash-err">' + text["login_error"] + "</div>"

    body = (
        build_lang_switch(lang).replace('class="lang-switch"', 'class="lang-switch standalone"')
        + '<div class="card" style="padding:2rem">'
        '<div style="text-align:center;margin-bottom:1.5rem">'
        '<div style="font-size:40px">&#9201;</div>'
        '<h1 style="font-size:20px;font-weight:700;color:#1e3a8a;margin-top:8px">Scriptly</h1>'
        + '<p style="font-size:12px;color:#888;margin-top:3px">' + text["login_subtitle"] + '</p>'
        "</div>"
        + error
        + '<form method="POST">'
        + '<label class="field-label">' + text["login_username"] + '</label>'
        '<input type="text" name="username" required autofocus>'
        + '<label class="field-label">' + text["login_password"] + '</label>'
        '<input type="password" name="password" required>'
        + '<button type="submit" class="btn btn-blue" style="width:100%;padding:12px;font-size:15px;margin-top:.5rem">' + text["login_submit"] + '</button>'
        "</form>"
        '<p style="text-align:center;margin-top:1.5rem;font-size:11px;color:#bbb">&#169; Scriptly</p>'
        "</div>"
    )
    return render(text["login_page_title"], body, nav=False, lang=lang)


@app.route("/logout")
def logout():
    session.clear()
    return redirect("/")


@app.route("/dashboard")
@login_required
def dashboard():
    if session.get("is_admin"):
        return redirect("/admin")

    lang = get_flow_lang()
    text = get_flow_text(lang)
    with get_db() as db:
        user = db.execute("SELECT * FROM users WHERE id=?", (session["user_id"],)).fetchone()
        perms = db.execute("SELECT script_id FROM permissions WHERE user_id=?", (session["user_id"],)).fetchall()
    if user is None:
        session.clear()
        return redirect("/")

    allowed = [get_localized_script(SCRIPTS[p["script_id"]], lang) for p in perms if p["script_id"] in SCRIPTS]
    status = get_account_status(user)
    status_label = status["status_label_he"] if lang == "he" else status["status_label_en"]
    not_set = "לא הוגדר" if lang == "he" else "Not set"
    status_colors = {
        "trial": ("#fff7ed", "#c2410c"),
        "active": ("#ecfdf5", "#047857"),
        "expired": ("#fef2f2", "#b91c1c"),
        "unknown": ("#f8fafc", "#475569"),
    }
    status_bg, status_fg = status_colors.get(status["status_key"], ("#f8fafc", "#475569"))
    if status["status_key"] == "trial":
        service_note = f"נותרו {status['days_remaining']} ימי ניסיון" if lang == "he" else f"{status['days_remaining']} trial days remaining"
    elif status["renewal_date"]:
        service_note = ("בתוקף עד " if lang == "he" else "Valid until ") + format_ui_date(status["renewal_date"], lang)
    else:
        service_note = "ללא פרטי שירות נוספים" if lang == "he" else "No additional service details"

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
            "<div>" + text["dashboard_empty"] + "</div>"
            "</div>"
        )

    info_items = [
        ("שם חברה" if lang == "he" else "Company", user["company_name"] or user["full_name"] or user["username"]),
        ("שם משתמש" if lang == "he" else "Username", user["username"]),
        ("איש קשר" if lang == "he" else "Contact name", user["full_name"]),
        ("ח.פ / מזהה חברה" if lang == "he" else "Company ID", user["company_id"]),
        ("אימייל" if lang == "he" else "Email", user["email"]),
        ("טלפון" if lang == "he" else "Phone", user["phone"]),
        ("תאריך הצטרפות" if lang == "he" else "Join date", format_ui_date(user["join_date"], lang)),
    ]
    info_grid = "".join(
        '<div style="background:#f8fafc;border:1px solid #e2e8f0;border-radius:14px;padding:14px">'
        '<div style="font-size:12px;color:#64748b;margin-bottom:6px">' + esc(label) + '</div>'
        '<div style="font-size:15px;font-weight:700;color:#0f172a">' + esc(value or not_set) + "</div>"
        "</div>"
        for label, value in info_items
    )
    body = (
        '<h2 style="font-size:24px;font-weight:800;color:#1e3a8a;margin-bottom:.4rem">' + text["dashboard_greeting"]
        + esc(session["name"])
        + ' &#128075;</h2>'
        + ('<p style="font-size:14px;color:#64748b;margin-bottom:1.5rem">' + text["dashboard_intro"] + "</p>" if text["dashboard_intro"] else "")
        +
        '<div style="display:grid;grid-template-columns:1.15fr .85fr;gap:1rem;margin-bottom:1rem">'
        '<div class="card" style="margin:0"><div style="font-size:18px;font-weight:800;color:#0f172a;margin-bottom:14px">'
        + ("פרטי חשבון ולקוח" if lang == "he" else "Account and company details")
        + '</div><div style="display:grid;grid-template-columns:repeat(auto-fit,minmax(180px,1fr));gap:12px">'
        + info_grid
        + '</div></div>'
        '<div class="card" style="margin:0;background:linear-gradient(180deg,#ffffff 0%,#f8fbff 100%)">'
        '<div style="font-size:18px;font-weight:800;color:#0f172a;margin-bottom:12px">' + ("סטטוס שירות" if lang == "he" else "Service status") + '</div>'
        '<div style="display:inline-flex;align-items:center;padding:8px 12px;border-radius:999px;background:' + status_bg + ';color:' + status_fg + ';font-size:14px;font-weight:800;margin-bottom:10px">' + esc(status_label) + '</div>'
        '<div style="font-size:14px;color:#334155;margin-bottom:8px">' + esc(service_note) + '</div>'
        '<div style="font-size:12px;color:#64748b;margin-bottom:4px">' + ("מסלול חיוב" if lang == "he" else "Billing mode") + '</div>'
        '<div style="font-size:15px;font-weight:700;color:#0f172a">' + esc(billing_mode_label(user["billing_mode"], lang)) + '</div>'
        '</div></div>'
        '<div class="card" style="margin:0"><div style="font-size:18px;font-weight:800;color:#0f172a;margin-bottom:14px">'
        + ("הכלים הזמינים לך" if lang == "he" else "Your available tools")
        + '</div><div style="display:grid;grid-template-columns:repeat(auto-fill,minmax(200px,1fr));gap:1rem">'
        + cards
        + '</div></div>'
        '<div class="card" style="margin-top:1rem;background:linear-gradient(135deg,#eff6ff 0%,#f8fafc 100%);border:1px solid #bfdbfe">'
        '<div style="font-size:18px;font-weight:800;color:#1e3a8a;margin-bottom:8px">' + ("אבטחת מידע ופרטיות" if lang == "he" else "Security and privacy") + '</div>'
        '<div style="font-size:14px;line-height:1.8;color:#334155">'
        + (
            "הדוחות שאתם מעלים משמשים לעיבוד בלבד ואינם נשמרים כחלק ממאגר קבוע. קבצי העבודה ותוצרי העיבוד נשמרים זמנית לצורך השלמת התהליך וההורדה, ולאחר מכן מנוקים באופן אוטומטי."
            if lang == "he"
            else "Uploaded reports are used only for processing and are not kept as part of a permanent data store. Working files and generated outputs are kept temporarily for processing and download, and are cleaned up automatically afterward."
        )
        + '</div></div>'
        '<details id="serviceTerms" style="margin-top:1rem;background:#ffffff;border:1px solid #e2e8f0;border-radius:16px;box-shadow:0 2px 16px rgba(0,0,0,.04);overflow:hidden">'
        '<summary style="list-style:none;cursor:pointer;padding:16px 18px;font-size:15px;font-weight:800;color:#0f172a;display:flex;align-items:center;justify-content:space-between;background:linear-gradient(180deg,#ffffff 0%,#f8fafc 100%);direction:rtl;text-align:right">'
        '<span>פירוט השירות והמנוי</span><span style="font-size:18px;color:#64748b">+</span></summary>'
        '<div style="padding:0 18px 18px;font-size:14px;line-height:1.9;color:#334155;direction:rtl;text-align:right">'
        'הפלטפורמה מרכזת כלים ודוחות שפותחו מתוך צרכים אמיתיים שעלו מהשטח.<br>'
        'לקוחות מנויים נהנים מגישה לכלל הכלים הזמינים במערכת.<br>'
        'פיתוח של כלי חדש מתבצע בתמחור חד־פעמי נפרד, בהתאם להיקף העבודה ובהצעת מחיר מראש.<br>'
        'לקוחות בתקופת ניסיון אינם כלולים בשירות המלא.<br>'
        'עלות השירות: 250 ש&quot;ח לחודש, בחיוב שנתי מראש, בתוספת מע&quot;מ.'
        '<br>ליצירת קשר ומענה: בוואטסאפ 0525776994 או במייל shaiproductionsltd@gmail.com'
        '</div></details>'
        '<script>'
        'function trackUserActivity(eventType, actionLabel, scriptId, scriptName, details){try{var data=new FormData();data.append("event_type",eventType||"");data.append("action_label",actionLabel||"");data.append("script_id",scriptId||"");data.append("script_name",scriptName||"");data.append("details",details||"");if(navigator.sendBeacon){navigator.sendBeacon("/activity",data);}else{fetch("/activity",{method:"POST",body:data,credentials:"same-origin",keepalive:true});}}catch(e){}}'
        'var serviceTerms=document.getElementById("serviceTerms");if(serviceTerms){serviceTerms.addEventListener("toggle",function(){if(this.open){trackUserActivity("open_service_terms","פתח פירוט שירות ומחיר","","","פירוט השירות והמנוי");}});}'
        '</script>'
    )
    return render(
        text["dashboard_page_title"],
        body,
        lang=lang,
        topbar_greeting=text["topbar_greeting"],
        logout_label=text["logout"],
        show_lang_switch=True,
    )


@app.route("/run/<script_id>", methods=["GET", "POST"])
@login_required
def run_script(script_id):
    if session.get("is_admin"):
        return redirect("/admin")

    lang = get_flow_lang()
    text = get_flow_text(lang)
    with get_db() as db:
        perm = db.execute(
            "SELECT 1 FROM permissions WHERE user_id=? AND script_id=?",
            (session["user_id"], script_id),
        ).fetchone()

    base_script = get_script(script_id)
    scr = get_localized_script(base_script, lang) if base_script else None

    if not perm or scr is None:
        add_flash(text["run_access_denied"])
        return redirect("/dashboard")
    if request.method == "GET":
        log_user_activity("open_script", "פתח כלי", script_id, scr["name"], "")
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
            options = {}
            extra_paths = []
            for field in scr.get("filter_fields", []):
                options[field["name"]] = request.form.get(field["name"], "").strip()
            result_name = build_output_filename(scr, uid, options)
            out = str(OUTPUT_FOLDER / result_name)
            for upload in scr.get("extra_uploads", []):
                extra_file = request.files.get(upload["name"])
                if extra_file and extra_file.filename:
                    extra_ext = get_extension(extra_file.filename)
                    expected = upload.get("accept", "").lstrip(".").lower()
                    if expected and extra_ext != expected:
                        error = '<div class="flash-err">' + text["run_extra_file_type_error"] + "</div>"
                        break
                    extra_path = str(UPLOAD_FOLDER / f"{uid}_{upload['name']}.{extra_ext or 'dat'}")
                    extra_file.save(extra_path)
                    options[f"{upload['name']}_path"] = extra_path
                    extra_paths.append(extra_path)
                elif upload.get("required"):
                    error = '<div class="flash-err">' + text["run_missing_extra_file_error"] + "</div>"
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
                    log_user_activity("generate_report", "הפיק דוח", script_id, scr["name"], result_name)
                except (xlrd.biffh.XLRDError, BadZipFile, OSError, ValueError):
                    error = '<div class="flash-err">' + scr["processing_error"] + '</div>'
                except Exception as e:
                    error = '<div class="flash-err">' + text["run_unexpected_error_prefix"] + str(e) + "</div>"
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
                if field.get("type") == "select":
                    select_options = ""
                    default_value = field.get("default", "")
                    for option in field.get("options", []):
                        selected = ' selected' if option.get("value") == default_value else ""
                        select_options += '<option value="' + option.get("value", "") + '"' + selected + '>' + option.get("label", "") + '</option>'
                    filter_fields_html += (
                        '<div><label class="field-label">' + field["label"] + '</label>'
                        + '<select name="' + field["name"] + '" style="padding:9px 12px;border:1.5px solid #e2e8f0;border-radius:8px;font-size:13px;font-family:inherit;outline:none;width:100%;margin-bottom:0;background:white">'
                        + select_options
                        + '</select></div>'
                    )
                else:
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

    help_trigger_html = ""
    help_modal_html = ""
    activity_script_id = json.dumps(script_id)
    activity_script_name = json.dumps(scr["name"])
    if lang == "he" and scr.get("help_title"):
        help_items_html = "".join(
            '<li style="margin-bottom:6px">' + esc(item) + '</li>'
            for item in scr.get("help_items", [])
        )
        help_trigger_html = (
            '<button type="button" onclick="openHelpModal()" style="display:inline-flex;align-items:center;gap:6px;border:1px solid #bfdbfe;background:#eff6ff;color:#1d4ed8;border-radius:999px;padding:6px 10px;font-size:12px;font-weight:700;cursor:pointer;font-family:inherit">'
            '<span style="display:inline-flex;align-items:center;justify-content:center;width:18px;height:18px;border-radius:999px;background:#dbeafe">?</span>'
            + esc(scr.get("help_label", "מידע נוסף"))
            + '</button>'
        )
        help_modal_html = (
            '<div class="modal-bg" id="helpModal" onclick="closeHelpModal(event)">'
            '<div class="modal-box" style="width:100%;max-width:560px;padding:1.5rem 1.5rem 1.25rem;border-radius:18px">'
            '<div style="font-size:18px;font-weight:800;color:#0f172a;margin-bottom:.75rem">' + esc(scr["help_title"]) + '</div>'
            '<div style="font-size:14px;line-height:1.8;color:#334155">'
            + esc(scr.get("help_intro", ""))
            + '<ul style="margin:.6rem 0 .75rem;padding-inline-start:1.2rem">'
            + help_items_html
            + '</ul>'
            + esc(scr.get("help_note", ""))
            + '</div>'
            + '<div style="margin-top:1rem"><button type="button" class="btn btn-blue" style="width:100%" onclick="closeHelpModal()">סגור</button></div>'
            + '</div></div>'
        )

    body = (
        '<a href="/dashboard" style="color:#2563eb;font-size:13px;text-decoration:none;display:block;margin-bottom:1rem">' + text["back_arrow"] + ' ' + scr["back_label"] + '</a>'
        + '<div class="card">'
        + '<div style="font-size:40px;margin-bottom:.5rem">' + scr["icon"] + '</div>'
        + '<div style="font-size:20px;font-weight:700;color:#1e3a8a;margin-bottom:4px">' + scr["name"] + '</div>'
        + '<div style="display:flex;align-items:center;justify-content:center;gap:10px;flex-wrap:wrap;margin-bottom:1.75rem">'
        + '<div style="font-size:13px;color:#64748b">' + scr["desc"] + '</div>'
        + help_trigger_html
        + '</div>'
        + content
        + '</div>'
        + help_modal_html
        + '<script>'
        + 'var fileInput=document.getElementById("fi");'
        + 'var label=document.getElementById("lbl");'
        + 'var button=document.getElementById("gb");'
        + 'var form=document.getElementById("uploadForm");'
        + 'function trackUserActivity(eventType, actionLabel, scriptId, scriptName, details){try{var data=new FormData();data.append("event_type",eventType||"");data.append("action_label",actionLabel||"");data.append("script_id",scriptId||"");data.append("script_name",scriptName||"");data.append("details",details||"");if(navigator.sendBeacon){navigator.sendBeacon("/activity",data);}else{fetch("/activity",{method:"POST",body:data,credentials:"same-origin",keepalive:true});}}catch(e){}}'
        + 'var activityScriptId=' + activity_script_id + ';'
        + 'var activityScriptName=' + activity_script_name + ';'
        + 'function openHelpModal(){var modal=document.getElementById("helpModal");if(modal){modal.style.display="flex";trackUserActivity("open_help_popup","פתח מידע והנחיות נוספות",activityScriptId,activityScriptName,"");}}'
        + 'function closeHelpModal(event){if(event && event.target && event.target.id!=="helpModal"){return;}var modal=document.getElementById("helpModal");if(modal){modal.style.display="none";}}'
        + 'if(fileInput && label){fileInput.addEventListener("change", function(){if(this.files && this.files.length){label.textContent=this.files[0].name;}});}'
        + 'if(form){form.addEventListener("submit", function(event){if(!fileInput || !fileInput.files || !fileInput.files.length){event.preventDefault();return false;}button.disabled=true;button.textContent="' + scr["processing_title"] + '";var box=document.getElementById("processingBox");if(box){box.classList.add("show");}return true;});}'
        + '</script>'
    )
    return render(
        scr["name"],
        body,
        lang=lang,
        topbar_greeting=text["topbar_greeting"],
        logout_label=text["logout"],
        show_lang_switch=True,
    )

@app.route("/download/<filename>")
@login_required
def download(filename):
    path = OUTPUT_FOLDER / filename
    if not path.exists():
        add_flash("File not found")
        return redirect("/dashboard")
    download_name = filename.split("_", 1)[-1] if "_" in filename else filename
    script = resolve_script_from_output_name(filename)
    if script:
        log_user_activity("download_report", "הוריד דוח", script.get("id", ""), script.get("name", ""), download_name)
    return send_file(path, as_attachment=True, download_name=download_name)


@app.route("/activity", methods=["POST"])
@login_required
def activity():
    if session.get("is_admin"):
        return "", 204
    event_type = request.form.get("event_type", "").strip()
    action_label = request.form.get("action_label", "").strip()
    if event_type and action_label:
        log_user_activity(
            event_type,
            action_label,
            request.form.get("script_id", "").strip(),
            request.form.get("script_name", "").strip(),
            request.form.get("details", "").strip(),
        )
    return "", 204


@app.route("/admin")
@login_required
@admin_required
def admin():
    activity_user_query = request.args.get("activity_user", "").strip()
    activity_range = request.args.get("activity_range", "all").strip() or "all"
    activity_from = request.args.get("activity_from", "").strip()
    activity_to = request.args.get("activity_to", "").strip()
    with get_db() as db:
        users = db.execute("SELECT * FROM users WHERE is_admin=0").fetchall()
        perms = db.execute("SELECT * FROM permissions").fetchall()
        all_activity_logs = db.execute("SELECT * FROM activity_logs ORDER BY created_at DESC, id DESC").fetchall()

    user_perms = {}
    for perm in perms:
        user_perms.setdefault(perm["user_id"], set()).add(perm["script_id"])

    rows = ""
    for user in users:
        uid = user["id"]
        status = get_account_status(user)
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
            "<td><strong>" + esc(user["company_name"] or user["full_name"] or user["username"]) + "</strong><br><span style=\"font-size:12px;color:#64748b\">ח.פ: " + esc(user["company_id"] or "לא הוגדר") + "</span></td>"
            '<td><div style="font-weight:700;color:#0f172a">' + esc(user["full_name"] or "לא הוגדר") + '</div><div style="font-size:12px;color:#64748b">@' + esc(user["username"]) + '</div><div style="font-size:12px;color:#64748b">' + esc(user["email"] or "ללא אימייל") + '</div><div style="font-size:12px;color:#64748b">' + esc(user["phone"] or "ללא טלפון") + '</div><div style="font-size:12px;color:#64748b">הצטרפות: ' + esc(format_ui_date(user["join_date"], "he")) + "</div></td>"
            '<td><div style="font-weight:700;color:#0f172a">' + esc(status["status_label_he"]) + '</div>'
            + ('<div style="font-size:12px;color:#64748b">נותרו ' + str(status["days_remaining"]) + ' ימים</div>' if status["days_remaining"] is not None else "")
            + ('<div style="font-size:12px;color:#64748b">בתוקף עד ' + esc(format_ui_date(status["renewal_date"], "he")) + '</div>' if status["renewal_date"] else "")
            + '<div style="font-size:12px;color:#64748b">חיוב: ' + esc(billing_mode_label(user["billing_mode"], "he")) + "</div></td>"
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
        "<table><thead><tr><th>Company</th><th>Contact</th><th>Service</th><th>Permissions</th><th>Password</th><th>Delete</th></tr></thead><tbody>"
        + rows
        + "</tbody></table>"
    ) if users else '<p style="color:#94a3b8;text-align:center;padding:2rem">No users yet</p>'

    filtered_activity_logs = []
    today = date.today()
    range_start = None
    range_end = None
    if activity_range == "last_30":
        range_start = datetime.combine(today - timedelta(days=29), datetime.min.time())
        range_end = datetime.combine(today, datetime.max.time())
    elif activity_range == "custom":
        parsed_from = parse_iso_date(activity_from)
        parsed_to = parse_iso_date(activity_to)
        if parsed_from:
            range_start = datetime.combine(parsed_from, datetime.min.time())
        if parsed_to:
            range_end = datetime.combine(parsed_to, datetime.max.time())

    for entry in all_activity_logs:
        username_text = str(entry["username"] or "")
        full_name_text = str(entry["full_name"] or "")
        if activity_user_query:
            query = activity_user_query.lower()
            if query not in username_text.lower() and query not in full_name_text.lower():
                continue
        created_at = parse_datetime_value(entry["created_at"])
        if range_start and created_at and created_at < range_start:
            continue
        if range_end and created_at and created_at > range_end:
            continue
        if range_start and not created_at:
            continue
        if range_end and not created_at:
            continue
        filtered_activity_logs.append(entry)

    total_activity = len(filtered_activity_logs)
    generated_reports = sum(1 for entry in filtered_activity_logs if entry["event_type"] == "generate_report")
    opened_tools = sum(1 for entry in filtered_activity_logs if entry["event_type"] == "open_script")
    help_opens = sum(1 for entry in filtered_activity_logs if entry["event_type"] == "open_help_popup")
    terms_opens = sum(1 for entry in filtered_activity_logs if entry["event_type"] == "open_service_terms")

    activity_rows = ""
    for entry in filtered_activity_logs:
        user_label = entry["full_name"] or entry["username"] or ("User #" + str(entry["user_id"]))
        activity_rows += (
            "<tr>"
            '<td>' + esc(format_ui_datetime(entry["created_at"])) + "</td>"
            '<td><div style="font-weight:700;color:#0f172a">' + esc(user_label) + '</div><div style="font-size:12px;color:#64748b">@' + esc(entry["username"] or "") + "</div></td>"
            '<td>' + esc(entry["action_label"]) + "</td>"
            '<td>' + esc(entry["script_name"] or "ללא כלי") + "</td>"
            '<td>' + esc(entry["details"] or "—") + "</td>"
            "</tr>"
        )
    activity_table = (
        "<table><thead><tr><th>When</th><th>User</th><th>Action</th><th>Tool</th><th>Details</th></tr></thead><tbody>"
        + activity_rows
        + "</tbody></table>"
    ) if filtered_activity_logs else '<p style="color:#94a3b8;text-align:center;padding:2rem">No activity matches the current filters</p>'
    activity_summary = (
        '<div style="display:grid;grid-template-columns:repeat(auto-fit,minmax(160px,1fr));gap:10px;margin-bottom:1rem">'
        + '<div style="background:#f8fafc;border:1px solid #e2e8f0;border-radius:12px;padding:12px"><div style="font-size:12px;color:#64748b;margin-bottom:6px">Total events</div><div style="font-size:20px;font-weight:800;color:#0f172a">' + str(total_activity) + '</div></div>'
        + '<div style="background:#f8fafc;border:1px solid #e2e8f0;border-radius:12px;padding:12px"><div style="font-size:12px;color:#64748b;margin-bottom:6px">Tools opened</div><div style="font-size:20px;font-weight:800;color:#0f172a">' + str(opened_tools) + '</div></div>'
        + '<div style="background:#f8fafc;border:1px solid #e2e8f0;border-radius:12px;padding:12px"><div style="font-size:12px;color:#64748b;margin-bottom:6px">Reports generated</div><div style="font-size:20px;font-weight:800;color:#0f172a">' + str(generated_reports) + '</div></div>'
        + '<div style="background:#f8fafc;border:1px solid #e2e8f0;border-radius:12px;padding:12px"><div style="font-size:12px;color:#64748b;margin-bottom:6px">Help opened</div><div style="font-size:20px;font-weight:800;color:#0f172a">' + str(help_opens) + '</div></div>'
        + '<div style="background:#f8fafc;border:1px solid #e2e8f0;border-radius:12px;padding:12px"><div style="font-size:12px;color:#64748b;margin-bottom:6px">Terms opened</div><div style="font-size:20px;font-weight:800;color:#0f172a">' + str(terms_opens) + '</div></div>'
        + '</div>'
    )
    activity_filter_bar = (
        '<form method="GET" action="/admin" style="display:grid;grid-template-columns:repeat(auto-fit,minmax(160px,1fr));gap:10px;margin-bottom:1rem">'
        '<div><label class="field-label">User</label><input type="text" name="activity_user" value="' + esc(activity_user_query) + '" placeholder="Name or username" style="margin-bottom:0"></div>'
        '<div><label class="field-label">Date range</label><select name="activity_range" style="padding:9px 12px;border:1.5px solid #e2e8f0;border-radius:8px;font-size:13px;font-family:inherit;outline:none;width:100%;margin-bottom:0;background:white">'
        + '<option value="all"' + (' selected' if activity_range == "all" else '') + '>All activity</option>'
        + '<option value="last_30"' + (' selected' if activity_range == "last_30" else '') + '>Last 30 days</option>'
        + '<option value="custom"' + (' selected' if activity_range == "custom" else '') + '>From date to date</option>'
        + '</select></div>'
        '<div><label class="field-label">From date</label><input type="text" name="activity_from" value="' + esc(activity_from) + '" placeholder="YYYY-MM-DD" style="margin-bottom:0"></div>'
        '<div><label class="field-label">To date</label><input type="text" name="activity_to" value="' + esc(activity_to) + '" placeholder="YYYY-MM-DD" style="margin-bottom:0"></div>'
        '<div style="display:flex;gap:8px;align-items:flex-end"><button type="submit" class="btn btn-blue" style="height:40px">Filter</button><a href="/admin" class="btn btn-gray" style="height:40px;display:inline-flex;align-items:center;text-decoration:none">Reset</a></div>'
        '</form>'
        '<div style="font-size:12px;color:#64748b;margin-bottom:1rem">Activity is available from the moment logging was enabled. Older button clicks and opens cannot be reconstructed retroactively.</div>'
    )

    body = (
        '<div class="card"><h2>&#10133; Add New User</h2><form method="POST" action="/admin/add_user"><div class="form-row">'
        '<div class="form-group"><label class="field-label">Full Name</label><input type="text" name="full_name" placeholder="Customer name" required style="margin-bottom:0"></div>'
        '<div class="form-group"><label class="field-label">Company Name</label><input type="text" name="company_name" placeholder="Company name" style="margin-bottom:0"></div>'
        '<div class="form-group"><label class="field-label">Company ID / ח.פ</label><input type="text" name="company_id" placeholder="Company ID" style="margin-bottom:0"></div>'
        '<div class="form-group"><label class="field-label">Username</label><input type="text" name="username" placeholder="Login username" required style="margin-bottom:0"></div>'
        '<div class="form-group"><label class="field-label">Password</label><input type="password" name="password" placeholder="Initial password" required style="margin-bottom:0"></div>'
        '</div><div class="form-row">'
        '<div class="form-group"><label class="field-label">Email</label><input type="text" name="email" placeholder="Email" style="margin-bottom:0"></div>'
        '<div class="form-group"><label class="field-label">Phone</label><input type="text" name="phone" placeholder="Phone" style="margin-bottom:0"></div>'
        '<div class="form-group"><label class="field-label">Billing Mode</label><select name="billing_mode" style="margin-bottom:0"><option value="monthly">Monthly</option><option value="yearly_prepaid">Yearly prepaid</option></select></div>'
        '<div class="form-group"><label class="field-label">Account Type</label><select name="account_type" style="margin-bottom:0"><option value="trial">30-day trial</option><option value="active">Active service</option></select></div>'
        '<div class="form-group"><label class="field-label">Valid Until</label><input type="text" name="service_valid_until" placeholder="YYYY-MM-DD" style="margin-bottom:0"></div>'
        '<button type="submit" class="btn btn-blue" style="height:40px;align-self:flex-end">Add</button></div></form></div>'
        '<div class="card"><h2>&#128101; Users In System</h2>'
        + table
        + '</div><div class="card"><h2>&#128221; User Activity Log</h2>'
        + activity_filter_bar
        + activity_summary
        + activity_table
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
    company_name = request.form.get("company_name", "").strip()
    company_id = request.form.get("company_id", "").strip()
    email = request.form.get("email", "").strip()
    phone = request.form.get("phone", "").strip()
    billing_mode = request.form.get("billing_mode", "monthly").strip() or "monthly"
    account_type = request.form.get("account_type", "trial").strip() or "trial"
    service_valid_until = request.form.get("service_valid_until", "").strip()
    join_date = date.today().isoformat()
    trial_start_date = join_date if account_type == "trial" else ""
    service_until_value = service_valid_until if account_type == "active" else ""
    try:
        with get_db() as db:
            db.execute(
                """INSERT INTO users(
                username,password,full_name,company_name,company_id,email,phone,join_date,trial_start_date,service_valid_until,billing_mode
                ) VALUES (?,?,?,?,?,?,?,?,?,?,?)""",
                (
                    username,
                    generate_password_hash(password),
                    full_name,
                    company_name,
                    company_id,
                    email,
                    phone,
                    join_date,
                    trial_start_date,
                    service_until_value,
                    billing_mode,
                ),
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

