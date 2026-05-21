"""Extract weekly timesheet data from Excel workbook into flat tabular output."""

import argparse
import re
import sys
from datetime import datetime, time
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
import pandas as pd

REFERENCE_SHEETS = {
    "Master Template (9) JOB",
    "Certified Codes",
    "Employees",
    "Job Details",
    "Job List",
    "ColumnLists",
}

DAY_BLOCKS = [
    ("MON",  "L",  "Q"),
    ("TUE",  "R",  "W"),
    ("WED",  "X",  "AC"),
    ("THUR", "AD", "AI"),
    ("FRI",  "AJ", "AO"),
    ("SAT",  "AP", "AU"),
    ("SUN",  "AV", "BA"),
]

HOUR_COLS = ["RT", "OT", "DT", "PD", "4D", "4A"]
TOTALS_COLS = ["BB", "BC", "BD", "BE", "BF", "BG"]

# Grand weekly total for the "Total Hours" column (row 9 across days, summed in BB9)
TOTAL_HOURS_GRAND_CELL = "BB9"

OUTPUT_COLUMNS = [
    "Folder Name", "Box Name", "Excel Name", "Sheet Name",
    "Employee Name", "EE ID", "WC State", "ST",
    "Date", "Day", "Start", "Lunch Out", "Lunch In", "Stop",
    "Total Hours", "RT", "OT", "DT", "PD", "4D", "4A", "PTO", "HP",
    "QA_Flag",
]

# Fill colours for QA sheet
_GREEN  = PatternFill("solid", fgColor="C6EFCE")
_RED    = PatternFill("solid", fgColor="FFC7CE")
_YELLOW = PatternFill("solid", fgColor="FFEB9C")
_GREY   = PatternFill("solid", fgColor="D9D9D9")
_BOLD   = Font(bold=True)


def col_letter_to_index(col: str) -> int:
    result = 0
    for ch in col.upper():
        result = result * 26 + (ord(ch) - ord("A") + 1)
    return result


def fmt_date(val) -> str:
    if val is None:
        return ""
    if isinstance(val, datetime):
        return val.strftime("%m/%d/%y")
    return str(val)


def fmt_time(val) -> str:
    """Return H:MM AM/PM string; blank when the value is not a recognisable time."""
    if val is None:
        return ""
    if isinstance(val, datetime):
        fmt = "%#I:%M %p" if sys.platform == "win32" else "%-I:%M %p"
        return val.strftime(fmt)
    if isinstance(val, time):
        dt = datetime.combine(datetime.today(), val)
        fmt = "%#I:%M %p" if sys.platform == "win32" else "%-I:%M %p"
        return dt.strftime(fmt)
    # Numeric fraction-of-day (Excel time serial stored as float)
    if isinstance(val, float) and 0.0 <= val < 1.0:
        total_minutes = round(val * 24 * 60)
        hours, minutes = divmod(total_minutes, 60)
        period = "AM" if hours < 12 else "PM"
        display_hour = hours % 12 or 12
        return f"{display_hour}:{minutes:02d} {period}"
    return ""


def coerce_hours(val) -> float:
    """Numeric cell → float; None / string / error → 0.0."""
    if val is None:
        return 0.0
    if isinstance(val, (int, float)):
        return float(val)
    return 0.0


def coerce_total_hours(val):
    """
    For 'Total Hours' specifically: preserve the distinction between
    a real numeric value, a blank cell, and a formula error.
        numeric  → float
        None     → None (blank)
        string   → None  (e.g. '#VALUE!' — flagged separately)
    """
    if val is None:
        return None
    if isinstance(val, (int, float)):
        return float(val)
    return None  # error string like '#VALUE!'


def infer_box_name(filename: str) -> str:
    stem = Path(filename).stem
    m = re.search(r"WE[\s_](\d{2})[\s_](\d{2})[\s_](\d{2})", stem, re.IGNORECASE)
    if m:
        return f"WE {m.group(1)} {m.group(2)} {m.group(3)}"
    return stem


def _is_time_value(v) -> bool:
    """True if v is a recognisable time: datetime, time, or float fraction-of-day."""
    if isinstance(v, (datetime, time)):
        return True
    if isinstance(v, float) and 0.0 <= v < 1.0:
        return True
    return False


def _row_qa_flag(time_cells: dict, total_hrs, total_hrs_cell: str) -> str:
    """
    Return a QA flag describing extraction issues on a day row.

    time_cells: ordered dict {field_name: (cell_address, value)} for Start/Lunch Out/Lunch In/Stop.
    total_hrs:  raw value of the row-9 Total Hours cell.
    total_hrs_cell: address of that cell (e.g. 'L9').
    """
    # All four time cells empty → employee was off / no entry
    if all(v is None for (_addr, v) in time_cells.values()):
        return "no time data"

    issues = []
    # Any time cell containing a non-time value (e.g. "OFF", "VAC", "-", "6am") is the
    # likely cause of a downstream formula error in the Total Hours cell.
    for field_name, (addr, val) in time_cells.items():
        if val is None or _is_time_value(val):
            continue
        issues.append(f"{field_name} cell {addr} = {val!r} — not a time")

    # If Total Hours itself is a string (formula error), report it. When we've already
    # named the bad input above, this is redundant noise, so we only add it if no
    # input cell explained the error.
    if isinstance(total_hrs, str) and not issues:
        issues.append(f"Total Hours cell {total_hrs_cell} = {total_hrs!r} (source formula error)")

    return "; ".join(issues) if issues else "OK"


def extract_sheet(ws, folder_name: str, box_name: str, excel_name: str) -> list[dict]:
    sheet_name = ws.title
    employee_name = ws["A1"].value or ""
    ee_id = ws["B2"].value
    wc_state = ws["G13"].value or ""
    st = ws["H13"].value or ""

    rows = []

    for day_label, start_col, _ in DAY_BLOCKS:
        start_idx = col_letter_to_index(start_col)

        date_val   = ws.cell(row=3, column=start_idx).value
        time_start = ws.cell(row=5, column=start_idx).value
        lunch_out  = ws.cell(row=6, column=start_idx).value
        lunch_in   = ws.cell(row=7, column=start_idx).value
        stop       = ws.cell(row=8, column=start_idx).value
        total_hrs  = ws.cell(row=9, column=start_idx).value

        day_hours = {}
        for i, col_name in enumerate(HOUR_COLS):
            day_hours[col_name] = coerce_hours(ws.cell(row=24, column=start_idx + i).value)

        time_cells = {
            "Start":     (f"{start_col}5", time_start),
            "Lunch Out": (f"{start_col}6", lunch_out),
            "Lunch In":  (f"{start_col}7", lunch_in),
            "Stop":      (f"{start_col}8", stop),
        }
        qa_flag = _row_qa_flag(time_cells, total_hrs, f"{start_col}9")

        rows.append({
            "Folder Name":   folder_name,
            "Box Name":      box_name,
            "Excel Name":    excel_name,
            "Sheet Name":    sheet_name,
            "Employee Name": employee_name,
            "EE ID":         ee_id,
            "WC State":      wc_state,
            "ST":            st,
            "Date":          fmt_date(date_val),
            "Day":           day_label,
            "Start":         fmt_time(time_start),
            "Lunch Out":     fmt_time(lunch_out),
            "Lunch In":      fmt_time(lunch_in),
            "Stop":          fmt_time(stop),
            "Total Hours":   coerce_total_hours(total_hrs),
            "RT":  day_hours["RT"],
            "OT":  day_hours["OT"],
            "DT":  day_hours["DT"],
            "PD":  day_hours["PD"],
            "4D":  day_hours["4D"],
            "4A":  day_hours["4A"],
            "PTO": "",
            "HP":  "",
            "QA_Flag": qa_flag,
        })

    # TOTALS row — grand weekly totals from BB24:BG24 (pay types) and BB9 (Total Hours)
    totals_hours = {}
    for col_name, col_letter in zip(HOUR_COLS, TOTALS_COLS):
        totals_hours[col_name] = coerce_hours(ws.cell(row=24, column=col_letter_to_index(col_letter)).value)

    grand_total_hours_raw = ws[TOTAL_HOURS_GRAND_CELL].value
    grand_total_hours = coerce_total_hours(grand_total_hours_raw)

    # BB9 has a formula error (e.g. #VALUE! because some days had "OFF" or other
    # non-time text in the time cells, making Excel's arithmetic fail).
    # Fall back to summing valid daily Total Hours already extracted.
    if grand_total_hours is None and grand_total_hours_raw is not None:
        grand_total_hours = round(
            sum(r["Total Hours"] for r in rows if r["Total Hours"] is not None), 4
        )
        totals_flag = (
            f"Total Hours cell {TOTAL_HOURS_GRAND_CELL} = {grand_total_hours_raw!r} "
            f"— recomputed from day rows"
        )
    else:
        totals_flag = "OK"

    rows.append({
        "Folder Name":   folder_name,
        "Box Name":      box_name,
        "Excel Name":    excel_name,
        "Sheet Name":    sheet_name,
        "Employee Name": employee_name,
        "EE ID":         ee_id,
        "WC State":      wc_state,
        "ST":            st,
        "Date":          "",
        "Day":           "TOTALS",
        "Start":         "",
        "Lunch Out":     "",
        "Lunch In":      "",
        "Stop":          "",
        "Total Hours":   grand_total_hours,
        "RT":  totals_hours["RT"],
        "OT":  totals_hours["OT"],
        "DT":  totals_hours["DT"],
        "PD":  totals_hours["PD"],
        "4D":  totals_hours["4D"],
        "4A":  totals_hours["4A"],
        "PTO": "",
        "HP":  "",
        "QA_Flag": totals_flag,
    })

    return rows


def check_accuracy(employee_rows: list[dict]) -> tuple[bool, dict]:
    """
    Returns (all_pass, results_dict).
    results_dict maps col → (daily_sum, grand_total, status).
        status: "MATCH" | "MISMATCH" | "SKIPPED (#VALUE in source)"
    Cross-check: sum of 7 extracted day values == Excel's pre-computed grand total.
        Pay types  → row 24 (daily) vs BB24:BG24 (grand)
        Total Hours → row 9  (daily) vs BB9       (grand)
    SKIPPED rows don't fail the overall check (the source itself has the error,
    not our extraction).
    """
    day_rows   = [r for r in employee_rows if r["Day"] != "TOTALS"]
    totals_row = next(r for r in employee_rows if r["Day"] == "TOTALS")

    results  = {}
    all_pass = True

    # Pay-type columns
    for col in HOUR_COLS:
        daily_sum   = round(sum(r[col] for r in day_rows), 4)
        grand_total = totals_row[col]
        passed = abs(daily_sum - grand_total) <= 0.01
        results[col] = (daily_sum, grand_total, "MATCH" if passed else "MISMATCH")
        if not passed:
            all_pass = False

    # Total Hours — both sides may contain None when source has #VALUE!
    daily_th_values = [r["Total Hours"] for r in day_rows]
    grand_th        = totals_row["Total Hours"]
    if any(v is None for v in daily_th_values) or grand_th is None:
        results["Total Hours"] = (
            round(sum(v for v in daily_th_values if v is not None), 4),
            grand_th if grand_th is not None else "#VALUE!",
            "SKIPPED (#VALUE in source)",
        )
    else:
        daily_sum   = round(sum(daily_th_values), 4)
        passed = abs(daily_sum - grand_th) <= 0.01
        results["Total Hours"] = (daily_sum, grand_th, "MATCH" if passed else "MISMATCH")
        if not passed:
            all_pass = False

    return all_pass, results


def _write_excel(out_path: Path, data_df: pd.DataFrame, qa_rows: list[dict], run_info: dict):
    """Write the three-sheet workbook with conditional formatting."""
    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        # ── Sheet 1: Data ──────────────────────────────────────────────────
        data_df.to_excel(writer, sheet_name="Data", index=False)

        # ── Sheet 2: QA_Summary ────────────────────────────────────────────
        qa_df = pd.DataFrame(qa_rows)
        qa_df.to_excel(writer, sheet_name="QA_Summary", index=False)

        # ── Sheet 3: Run_Info ──────────────────────────────────────────────
        run_df = pd.DataFrame([run_info])
        run_df.to_excel(writer, sheet_name="Run_Info", index=False)

    # Post-process: apply colours
    wb = openpyxl.load_workbook(out_path)

    # --- Data sheet: colour QA_Flag column ---
    # Grey tint for TOTALS rows is driven off the Day column, not QA_Flag —
    # QA_Flag now carries pure status only ("OK" or a specific issue).
    ws_data = wb["Data"]
    qa_col_idx  = data_df.columns.get_loc("QA_Flag") + 1  # 1-based
    day_col_idx = data_df.columns.get_loc("Day") + 1
    for row in ws_data.iter_rows(min_row=2, max_row=ws_data.max_row):
        cell = row[qa_col_idx - 1]
        val  = str(cell.value or "")
        is_totals_row = str(row[day_col_idx - 1].value or "") == "TOTALS"
        if val == "OK":
            cell.fill = _GREY if is_totals_row else _GREEN
        else:
            cell.fill = _RED
        cell.font = Font(bold=(val != "OK"))

    # Auto-width Data sheet
    for col_cells in ws_data.columns:
        max_len = max((len(str(c.value)) for c in col_cells if c.value is not None), default=8)
        ws_data.column_dimensions[col_cells[0].column_letter].width = min(max_len + 2, 30)

    # --- QA_Summary sheet: colour Overall column + per-col match cells ---
    ws_qa = wb["QA_Summary"]
    headers = [c.value for c in ws_qa[1]]

    overall_idx = headers.index("Overall") + 1 if "Overall" in headers else None

    for row in ws_qa.iter_rows(min_row=2, max_row=ws_qa.max_row):
        # Bold header (employee name)
        row[0].font = _BOLD

        for cell in row:
            col_header = headers[cell.column - 1] if cell.column <= len(headers) else ""
            val = str(cell.value or "")

            if col_header == "Overall":
                cell.fill = _GREEN if val == "PASS" else _RED
                cell.font = Font(bold=True)
            elif col_header.endswith("_match"):
                if val == "MATCH":
                    cell.fill = _GREEN
                elif val.startswith("SKIPPED"):
                    cell.fill = _YELLOW
                else:
                    cell.fill = _RED

    # Auto-width QA sheet
    for col_cells in ws_qa.columns:
        max_len = max((len(str(c.value)) for c in col_cells if c.value is not None), default=8)
        ws_qa.column_dimensions[col_cells[0].column_letter].width = min(max_len + 2, 30)

    # Bold header row on all sheets
    for sheet_name in ("Data", "QA_Summary", "Run_Info"):
        ws = wb[sheet_name]
        for cell in ws[1]:
            cell.font = Font(bold=True)
            cell.fill = PatternFill("solid", fgColor="BDD7EE")

    # Freeze top row on all sheets
    for sheet_name in ("Data", "QA_Summary", "Run_Info"):
        wb[sheet_name].freeze_panes = "A2"

    wb.save(out_path)


def extract(input_path: str, folder_name: str, box_name: str | None, out_path: str) -> int:
    input_path = Path(input_path)
    excel_name = input_path.name

    if box_name is None:
        box_name = infer_box_name(excel_name)

    wb = openpyxl.load_workbook(input_path, data_only=True)

    all_rows: list[dict]    = []
    qa_rows:  list[dict]    = []
    pass_count = 0
    total_employees = 0

    print(f"\n{'Employee':<25} {'RT':>8} {'OT':>8} {'DT':>8} {'PD':>8} {'4D':>8} {'4A':>8}  Status")
    print("-" * 85)

    for sheet_name in wb.sheetnames:
        if sheet_name in REFERENCE_SHEETS:
            continue
        ws = wb[sheet_name]
        emp_rows = extract_sheet(ws, folder_name, box_name, excel_name)
        all_rows.extend(emp_rows)

        total_employees += 1
        passed, results = check_accuracy(emp_rows)
        if passed:
            pass_count += 1

        emp_name = emp_rows[0]["Employee Name"]
        status   = "PASS" if passed else "FAIL"

        # Console line — short status per column
        short_status = {"MATCH": "OK", "MISMATCH": "!!", "SKIPPED (#VALUE in source)": "sk"}
        col_info = "  ".join(f"{col}:{short_status.get(r[2], '?')}" for col, r in results.items())
        print(f"{emp_name:<25} {col_info}  {status}")

        # Build QA_Summary row — one row per employee with full transparency
        qa_row: dict = {
            "Employee":    emp_name,
            "Sheet Name":  sheet_name,
            "QA Method":   "sum(MON..SUN) vs Excel grand total (pay types: BB24:BG24; Total Hours: BB9)",
            "Tolerance":   0.01,
        }
        for col, (daily_sum, grand_total, col_status) in results.items():
            qa_row[f"{col}_daily_sum"]   = daily_sum
            qa_row[f"{col}_grand_total"] = grand_total
            qa_row[f"{col}_match"]       = col_status
        qa_row["Overall"] = status
        qa_rows.append(qa_row)

    print(f"\n{pass_count}/{total_employees} employees passed accuracy check.\n")

    run_info = {
        "Input File":        excel_name,
        "Folder Name":       folder_name,
        "Box Name":          box_name,
        "Run Timestamp":     datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "Employees Checked": total_employees,
        "Employees Passed":  pass_count,
        "Employees Failed":  total_employees - pass_count,
        "Overall Result":    "PASS" if pass_count == total_employees else "FAIL",
        "QA Method":         "sum(MON..SUN per column) == Excel grand total (BB24:BG24), tolerance 0.01",
        "Script Version":    "1.1",
    }

    data_df  = pd.DataFrame(all_rows, columns=OUTPUT_COLUMNS)
    out_path = Path(out_path)
    out_path.parent.mkdir(parents=True, exist_ok=True)

    _write_excel(out_path, data_df, qa_rows, run_info)
    print(f"Output written to: {out_path}")
    print("  Sheet 1 — Data       : extracted rows + QA_Flag per row")
    print("  Sheet 2 — QA_Summary : per-employee accuracy comparison")
    print("  Sheet 3 — Run_Info   : run metadata and overall result\n")

    return 0 if pass_count == total_employees else 2


def main():
    parser = argparse.ArgumentParser(description="Extract timesheet data from Excel workbook.")
    parser.add_argument("input", help="Path to input Excel workbook")
    parser.add_argument("--folder", default="2022", help="Folder Name value (default: 2022)")
    parser.add_argument("--box",    default=None,   help="Box Name value (inferred from filename if omitted)")
    parser.add_argument("--out",    default="output/extracted.xlsx", help="Output Excel path")
    args = parser.parse_args()

    sys.exit(extract(args.input, args.folder, args.box, args.out))


if __name__ == "__main__":
    main()
