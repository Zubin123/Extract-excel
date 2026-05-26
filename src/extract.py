"""Extract weekly timesheet data from Excel workbook(s) into flat tabular output.

Phase 1 of the robust pipeline:
  - Config-driven profiles (config/schema.yaml)
  - Structural sheet classification (no hardcoded reference-sheet name list)
  - Pay-totals row discovered via sum cross-check
  - Sheets that fail anchor resolution are listed on the Unmatched_Sheets tab
"""

from __future__ import annotations

import argparse
import re
import sys
from datetime import datetime, time
from pathlib import Path

import openpyxl
import pandas as pd
import yaml
from openpyxl.styles import Font, PatternFill

from anchors import (
    classify_sheet,
    col_letter_to_index,
    resolve_anchor_cell,
    resolve_pay_totals_row,
    select_profile,
    sheet_has_day_grid,
)
from qa import (
    check_date_sequence,
    check_ee_id_consistency,
    check_filename_week_match,
    check_time_order,
    check_weekly_hours_range,
    two_pass_totals_row,
)

DEFAULT_DAY_LABELS = ["MON", "TUE", "WED", "THUR", "FRI", "SAT", "SUN"]
DEFAULT_PAY_TYPES  = ["RT", "OT", "DT", "PD", "4D", "4A"]


_GREEN  = PatternFill("solid", fgColor="C6EFCE")
_RED    = PatternFill("solid", fgColor="FFC7CE")
_YELLOW = PatternFill("solid", fgColor="FFEB9C")
_GREY   = PatternFill("solid", fgColor="D9D9D9")
_BOLD   = Font(bold=True)


# ── Formatters ──────────────────────────────────────────────────────────────

def fmt_date(val) -> str:
    if val is None:
        return ""
    if isinstance(val, datetime):
        return val.strftime("%m/%d/%y")
    return str(val)


def fmt_time(val) -> str:
    if val is None:
        return ""
    if isinstance(val, datetime):
        fmt = "%#I:%M %p" if sys.platform == "win32" else "%-I:%M %p"
        return val.strftime(fmt)
    if isinstance(val, time):
        dt = datetime.combine(datetime.today(), val)
        fmt = "%#I:%M %p" if sys.platform == "win32" else "%-I:%M %p"
        return dt.strftime(fmt)
    if isinstance(val, float) and 0.0 <= val < 1.0:
        total_minutes = round(val * 24 * 60)
        hours, minutes = divmod(total_minutes, 60)
        period = "AM" if hours < 12 else "PM"
        display_hour = hours % 12 or 12
        return f"{display_hour}:{minutes:02d} {period}"
    return ""


def coerce_hours(val) -> float:
    if val is None:
        return 0.0
    if isinstance(val, (int, float)):
        return float(val)
    return 0.0


def coerce_total_hours(val):
    if val is None:
        return None
    if isinstance(val, (int, float)):
        return float(val)
    return None


def infer_box_name(filename: str) -> str:
    stem = Path(filename).stem
    m = re.search(r"WE[\s_](\d{2})[\s_](\d{2})[\s_](\d{2})", stem, re.IGNORECASE)
    if m:
        return f"WE {m.group(1)} {m.group(2)} {m.group(3)}"
    return stem


def _is_time_value(v) -> bool:
    if isinstance(v, (datetime, time)):
        return True
    if isinstance(v, float) and 0.0 <= v < 1.0:
        return True
    return False


def _row_qa_flag(time_cells: dict, total_hrs, total_hrs_cell: str) -> str:
    if all(v is None for (_addr, v) in time_cells.values()):
        return "OK (no time data)"
    issues = []
    for field_name, (addr, val) in time_cells.items():
        if val is None or _is_time_value(val):
            continue
        issues.append(f"{field_name} cell {addr} = {val!r} — not a time")
    if isinstance(total_hrs, str) and not issues:
        issues.append(f"Total Hours cell {total_hrs_cell} = {total_hrs!r} (source formula error)")
    return "; ".join(issues) if issues else "OK"


# ── Sheet extraction ────────────────────────────────────────────────────────

def _anchor_issue(field: str, status: str) -> str | None:
    if status in ("ok", "noverify"):
        return None
    if status.startswith("relocated:"):
        return f"[CHECK] {field} anchor relocated to {status.split(':', 1)[1]}"
    if status == "missing":
        return f"[CHECK] {field} header not found in header_row"
    return f"[CHECK] {field} anchor unknown status {status!r}"


def extract_sheet(ws, profile: dict, pay_totals_row: int, folder_name: str,
                  box_name: str, excel_name: str, kind: str = "employee") -> list[dict]:
    sheet_name = ws.title
    anchors = profile["anchors"]
    raw_a1 = ws[anchors["employee_name"]].value
    if kind == "employee_alt_layout":
        ee_id = ws["C2"].value
    else:
        ee_id = ws[anchors["ee_id"]].value
    wc_state_val, wc_status, _ = resolve_anchor_cell(ws, anchors.get("wc_state"))
    st_val,       st_status, _ = resolve_anchor_cell(ws, anchors.get("st"))
    wc_state = wc_state_val if wc_state_val is not None else ""
    st       = st_val       if st_val       is not None else ""

    anchor_issues: list[str] = []
    if kind == "employee_placeholder":
        employee_name = ""
        anchor_issues.append(
            f"[CHECK] Employee Name placeholder in A1 ({raw_a1!r}) — left blank (no reliable source)"
        )
    elif kind == "employee_alt_layout":
        employee_name = ws["B1"].value or ""
        anchor_issues.append(
            f"[CHECK] Alternate layout — Employee Name from B1, EE ID from C2"
        )
    else:
        employee_name = raw_a1 or ""

    for field, status in (("WC State", wc_status), ("ST", st_status)):
        msg = _anchor_issue(field, status)
        if msg:
            anchor_issues.append(msg)

    db = profile["day_blocks"]
    day_labels = db["day_labels"]
    day_start_cols = db["day_start_cols"]
    pay_types = db["pay_types"]

    date_row = profile["date_row"]
    total_hours_row = profile["total_hours_row"]
    t_rows = profile["time_rows"]

    rows = []

    for day_label, start_col in zip(day_labels, day_start_cols):
        start_idx = col_letter_to_index(start_col)

        date_val   = ws.cell(row=date_row, column=start_idx).value
        time_start = ws.cell(row=t_rows["Start"],     column=start_idx).value
        lunch_out  = ws.cell(row=t_rows["Lunch Out"], column=start_idx).value
        lunch_in   = ws.cell(row=t_rows["Lunch In"],  column=start_idx).value
        stop       = ws.cell(row=t_rows["Stop"],      column=start_idx).value
        total_hrs  = ws.cell(row=total_hours_row,     column=start_idx).value

        day_hours = {}
        for i, col_name in enumerate(pay_types):
            day_hours[col_name] = coerce_hours(
                ws.cell(row=pay_totals_row, column=start_idx + i).value
            )

        time_cells = {
            "Start":     (f"{start_col}{t_rows['Start']}",     time_start),
            "Lunch Out": (f"{start_col}{t_rows['Lunch Out']}", lunch_out),
            "Lunch In":  (f"{start_col}{t_rows['Lunch In']}",  lunch_in),
            "Stop":      (f"{start_col}{t_rows['Stop']}",      stop),
        }
        base_flag = _row_qa_flag(time_cells, total_hrs, f"{start_col}{total_hours_row}")
        order_issues = check_time_order(time_cells)
        flag_parts = [p for p in [base_flag if base_flag != "OK" else None, *order_issues] if p]
        qa_flag = "; ".join(flag_parts) if flag_parts else "OK"

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
            **{pt: day_hours[pt] for pt in pay_types},
            "PTO": "",
            "HP":  "",
            "QA_Flag": qa_flag,
        })

    # TOTALS row — grand weekly totals from grand_totals.cols on pay_totals_row
    # and total_hours_col on total_hours_row.
    gt_cols = profile["grand_totals"]["cols"]
    th_col = profile["grand_totals"]["total_hours_col"]

    totals_hours = {}
    for col_name, col_letter in zip(pay_types, gt_cols):
        totals_hours[col_name] = coerce_hours(
            ws.cell(row=pay_totals_row, column=col_letter_to_index(col_letter)).value
        )

    grand_th_cell = f"{th_col}{total_hours_row}"
    grand_total_hours_raw = ws[grand_th_cell].value
    grand_total_hours = coerce_total_hours(grand_total_hours_raw)

    if grand_total_hours is None and grand_total_hours_raw is not None:
        grand_total_hours = round(
            sum(r["Total Hours"] for r in rows if r["Total Hours"] is not None), 4
        )
        totals_flag = (
            f"Total Hours cell {grand_th_cell} = {grand_total_hours_raw!r} "
            f"— recomputed from day rows"
        )
    else:
        totals_flag = "OK"

    if anchor_issues:
        merged = "; ".join(anchor_issues)
        totals_flag = merged if totals_flag == "OK" else f"{totals_flag}; {merged}"

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
        **{pt: totals_hours[pt] for pt in pay_types},
        "PTO": "",
        "HP":  "",
        "QA_Flag": totals_flag,
    })

    return rows


def extract_sheet_no_grid(ws, folder_name: str, box_name: str,
                          excel_name: str, profile: dict | None = None,
                          kind: str = "employee") -> list[dict]:
    """Emit 8 blank rows (MON..SUN + TOTALS) for an employee sheet that has
    no weekly day-label grid (e.g. 'Job' / 'OVERHEAD- CA').

    When `profile` is supplied, wc_state/st use the same header-verified
    resolver as the grid-bearing path; otherwise they fall back to G13/H13.
    When `kind == "employee_placeholder"`, sheet name is used as the
    employee name and a [CHECK] flag is emitted on the TOTALS row.
    """
    sheet_name = ws.title
    raw_a1 = ws["A1"].value
    if kind == "employee_placeholder":
        employee_name = ""
    elif kind == "employee_alt_layout":
        employee_name = ws["B1"].value or ""
    else:
        employee_name = raw_a1 or ""

    if kind == "employee_alt_layout":
        ee_id = ws["C2"].value
    else:
        ee_id = ws["B2"].value

    if profile is not None:
        anchors = profile.get("anchors", {})
        wc_val, _wc_status, _ = resolve_anchor_cell(ws, anchors.get("wc_state"))
        st_val, _st_status, _ = resolve_anchor_cell(ws, anchors.get("st"))
        wc_state = wc_val if wc_val is not None else ""
        st       = st_val if st_val is not None else ""
    else:
        wc_state = ws["G13"].value or ""
        st = ws["H13"].value or ""

    base_flag = "no time grid"
    if kind == "employee_placeholder":
        totals_flag = (
            f"{base_flag}; [CHECK] Employee Name placeholder in A1 ({raw_a1!r}) — left blank"
        )
    elif kind == "employee_alt_layout":
        totals_flag = (
            f"{base_flag}; [CHECK] Alternate layout — Employee Name from B1, EE ID from C2"
        )
    else:
        totals_flag = base_flag

    rows = []
    for day_label in DEFAULT_DAY_LABELS + ["TOTALS"]:
        row = {
            "Folder Name":   folder_name,
            "Box Name":      box_name,
            "Excel Name":    excel_name,
            "Sheet Name":    sheet_name,
            "Employee Name": employee_name,
            "EE ID":         ee_id,
            "WC State":      wc_state,
            "ST":            st,
            "Date":          "",
            "Day":           day_label,
            "Start":         "",
            "Lunch Out":     "",
            "Lunch In":      "",
            "Stop":          "",
            "Total Hours":   None,
            **{pt: "" for pt in DEFAULT_PAY_TYPES},
            "PTO": "",
            "HP":  "",
            "QA_Flag": totals_flag if day_label == "TOTALS" else base_flag,
        }
        rows.append(row)
    return rows


def check_accuracy(employee_rows: list[dict], pay_types: list[str],
                   tolerance: float) -> tuple[bool, dict]:
    day_rows   = [r for r in employee_rows if r["Day"] != "TOTALS"]
    totals_row = next(r for r in employee_rows if r["Day"] == "TOTALS")

    results  = {}
    all_pass = True

    for col in pay_types:
        daily_sum   = round(sum(r[col] for r in day_rows), 4)
        grand_total = totals_row[col]
        passed = abs(daily_sum - grand_total) <= tolerance
        results[col] = (daily_sum, grand_total, "MATCH" if passed else "MISMATCH")
        if not passed:
            all_pass = False

    daily_th_values = [r["Total Hours"] for r in day_rows]
    grand_th        = totals_row["Total Hours"]
    if any(v is None for v in daily_th_values) or grand_th is None:
        results["Total Hours"] = (
            round(sum(v for v in daily_th_values if v is not None), 4),
            grand_th if grand_th is not None else "#VALUE!",
            "SKIPPED (#VALUE in source)",
        )
    else:
        daily_sum = round(sum(daily_th_values), 4)
        passed = abs(daily_sum - grand_th) <= tolerance
        results["Total Hours"] = (daily_sum, grand_th, "MATCH" if passed else "MISMATCH")
        if not passed:
            all_pass = False

    return all_pass, results


# ── Writer ──────────────────────────────────────────────────────────────────

def _write_excel(out_path: Path, data_df: pd.DataFrame, qa_rows: list[dict],
                 run_info: dict, unmatched: list[dict], id_conflicts: list[dict]):
    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        data_df.to_excel(writer, sheet_name="Data", index=False)
        pd.DataFrame(qa_rows).to_excel(writer, sheet_name="QA_Summary", index=False)
        pd.DataFrame([run_info]).to_excel(writer, sheet_name="Run_Info", index=False)
        pd.DataFrame(unmatched).to_excel(writer, sheet_name="Unmatched_Sheets", index=False)
        pd.DataFrame(id_conflicts).to_excel(writer, sheet_name="ID_Conflicts", index=False)

    wb = openpyxl.load_workbook(out_path)

    # Data sheet colouring
    ws_data = wb["Data"]
    qa_col_idx  = data_df.columns.get_loc("QA_Flag") + 1
    day_col_idx = data_df.columns.get_loc("Day") + 1
    for row in ws_data.iter_rows(min_row=2, max_row=ws_data.max_row):
        cell = row[qa_col_idx - 1]
        val  = str(cell.value or "")
        is_totals_row = str(row[day_col_idx - 1].value or "") == "TOTALS"
        # Treat as OK: literal "OK", "OK (no time data)", or anything where only
        # [INFO] notes are attached and no [CHECK]/[WARN]/error prefix appears.
        is_ok = (
            val == "OK"
            or val.startswith("OK (")
            or ("[CHECK]" not in val and "[WARN]" not in val
                and not any(tok in val for tok in ("— not a time", "#VALUE", "no time grid")))
        )
        if is_ok:
            cell.fill = _GREY if is_totals_row else _GREEN
        else:
            cell.fill = _RED
        cell.font = Font(bold=not is_ok)

    for col_cells in ws_data.columns:
        max_len = max((len(str(c.value)) for c in col_cells if c.value is not None), default=8)
        ws_data.column_dimensions[col_cells[0].column_letter].width = min(max_len + 2, 30)

    # QA_Summary colouring
    ws_qa = wb["QA_Summary"]
    headers = [c.value for c in ws_qa[1]]
    for row in ws_qa.iter_rows(min_row=2, max_row=ws_qa.max_row):
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

    for col_cells in ws_qa.columns:
        max_len = max((len(str(c.value)) for c in col_cells if c.value is not None), default=8)
        ws_qa.column_dimensions[col_cells[0].column_letter].width = min(max_len + 2, 30)

    # Header row + freeze panes on all sheets
    for sheet_name in ("Data", "QA_Summary", "Run_Info", "Unmatched_Sheets", "ID_Conflicts"):
        if sheet_name not in wb.sheetnames:
            continue
        ws = wb[sheet_name]
        for cell in ws[1]:
            cell.font = Font(bold=True)
            cell.fill = PatternFill("solid", fgColor="BDD7EE")
        ws.freeze_panes = "A2"

    wb.save(out_path)


# ── Pipeline ────────────────────────────────────────────────────────────────

def load_config(path: Path) -> dict:
    with open(path, "r", encoding="utf-8") as f:
        return yaml.safe_load(f)


def process_workbook(input_path: Path, config: dict, folder_name: str,
                     box_name: str | None) -> tuple[list[dict], list[dict], list[dict]]:
    """Return (data_rows, qa_rows, unmatched_rows) for one workbook."""
    excel_name = input_path.name
    if box_name is None:
        box_name = infer_box_name(excel_name)

    profiles = config["profiles"]
    tolerance = config.get("tolerance", 0.01)

    wb = openpyxl.load_workbook(input_path, data_only=True)

    data_rows: list[dict] = []
    qa_rows: list[dict] = []
    unmatched: list[dict] = []

    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        kind, classify_reason = classify_sheet(ws, profiles)
        if kind == "reference":
            continue
        if kind == "unfilled_template":
            unmatched.append({
                "File": excel_name, "Sheet": sheet_name,
                "Reason": f"unfilled employee template: {classify_reason}",
            })
            continue

        # Employee sheets with no weekly grid (e.g. 'Job', 'OVERHEAD- CA') —
        # emit 8 blank rows so the employee is still represented in output.
        if not sheet_has_day_grid(ws, profiles):
            ng_profile = next(iter(profiles.values()), None)
            no_grid_rows = extract_sheet_no_grid(
                ws, folder_name, box_name, excel_name, ng_profile, kind=kind,
            )
            data_rows.extend(no_grid_rows)
            placeholder_issue = (
                f"[CHECK] Employee Name placeholder — using sheet name {sheet_name!r}"
                if kind == "employee_placeholder" else ""
            )
            qa_rows.append({
                "File":           excel_name,
                "Employee":       no_grid_rows[0]["Employee Name"],
                "Sheet Name":     sheet_name,
                "Profile":        "(no grid)",
                "Pay Totals Row": "",
                "QA Method":      "no weekly day-label grid — 8 blank rows emitted",
                "Tolerance":      tolerance,
                "Sheet Issues":   placeholder_issue,
                "Overall":        "NO_GRID",
            })
            continue

        profile_name = select_profile(ws, profiles)
        if profile_name is None:
            unmatched.append({
                "File": excel_name, "Sheet": sheet_name,
                "Reason": "no matching profile (day labels not at expected position)",
            })
            continue

        profile = profiles[profile_name]
        pay_totals_row, reason = resolve_pay_totals_row(ws, profile, tolerance=tolerance)
        if pay_totals_row is None:
            unmatched.append({
                "File": excel_name, "Sheet": sheet_name,
                "Reason": f"pay_totals_row unresolved: {reason}",
            })
            continue

        emp_rows = extract_sheet(ws, profile, pay_totals_row,
                                 folder_name, box_name, excel_name, kind=kind)

        # Phase 2 sheet-level QA layers
        day_rows = [r for r in emp_rows if r["Day"] != "TOTALS"]
        totals_row = next(r for r in emp_rows if r["Day"] == "TOTALS")

        sheet_issues: list[str] = []
        sheet_issues += two_pass_totals_row(ws, profile, pay_totals_row, tolerance)
        sheet_issues += check_weekly_hours_range(totals_row.get("Total Hours"))
        sheet_issues += check_date_sequence(day_rows)
        sheet_issues += check_filename_week_match(excel_name, day_rows)

        # Surface sheet-level issues on the TOTALS row's QA_Flag
        if sheet_issues:
            existing = totals_row.get("QA_Flag", "OK")
            extra = "; ".join(sheet_issues)
            totals_row["QA_Flag"] = extra if existing == "OK" else f"{existing}; {extra}"

        data_rows.extend(emp_rows)

        passed, results = check_accuracy(emp_rows, profile["day_blocks"]["pay_types"], tolerance)
        emp_name = emp_rows[0]["Employee Name"]
        qa_row = {
            "File":       excel_name,
            "Employee":   emp_name,
            "Sheet Name": sheet_name,
            "Profile":    profile_name,
            "Pay Totals Row": pay_totals_row,
            "QA Method":  f"sum(MON..SUN) vs grand totals (row {pay_totals_row})",
            "Tolerance":  tolerance,
        }
        for col, (daily_sum, grand_total, col_status) in results.items():
            qa_row[f"{col}_daily_sum"]   = daily_sum
            qa_row[f"{col}_grand_total"] = grand_total
            qa_row[f"{col}_match"]       = col_status
        totals_qa = str(totals_row.get("QA_Flag", ""))
        anchor_issue_tokens = [seg.strip() for seg in totals_qa.split(";")
                               if "[CHECK]" in seg and "anchor" in seg.lower()]
        all_issues_for_display = sheet_issues + anchor_issue_tokens
        qa_row["Sheet Issues"] = "; ".join(all_issues_for_display) if all_issues_for_display else ""
        # REVIEW only if there's an actionable [CHECK] or [WARN]; [INFO] alone passes.
        needs_review = (
            any(("[CHECK]" in s or "[WARN]" in s) for s in sheet_issues)
            or "[CHECK]" in totals_qa
            or "[WARN]" in totals_qa
        )
        qa_row["Overall"] = "PASS" if (passed and not needs_review) else ("REVIEW" if passed else "FAIL")
        qa_rows.append(qa_row)

    return data_rows, qa_rows, unmatched


def extract(input_arg: str, folder_name: str, box_name: str | None, out_path: str,
            config_path: str) -> int:
    config = load_config(Path(config_path))
    output_columns = config["output_columns"]

    in_path = Path(input_arg)
    files: list[Path]
    if in_path.is_dir():
        files = sorted([p for p in in_path.rglob("*.xlsx") if not p.name.startswith("~$")])
    else:
        files = [in_path]

    all_data: list[dict] = []
    all_qa: list[dict] = []
    all_unmatched: list[dict] = []

    print(f"\nProcessing {len(files)} workbook(s)...")
    for f in files:
        try:
            data, qa, um = process_workbook(f, config, folder_name, box_name)
        except Exception as e:
            all_unmatched.append({"File": f.name, "Sheet": "(workbook)", "Reason": f"ERROR: {e!r}"})
            print(f"  {f.name[:60]:<60}  ERROR {e!r}")
            continue
        all_data.extend(data)
        all_qa.extend(qa)
        all_unmatched.extend(um)
        passed = sum(1 for r in qa if r["Overall"] == "PASS")
        total = len(qa)
        print(f"  {f.name[:60]:<60}  emp={total:>3}  pass={passed:>3}  unmatched={len(um):>2}")

    total_emp = len(all_qa)
    total_pass    = sum(1 for r in all_qa if r["Overall"] == "PASS")
    total_no_grid = sum(1 for r in all_qa if r["Overall"] == "NO_GRID")
    total_grid    = total_emp - total_no_grid
    print(f"\n{total_pass}/{total_grid} grid-bearing sheets passed accuracy check.")
    print(f"{total_no_grid} employee sheet(s) had no time grid (8 blank rows each).")
    print(f"{len(all_unmatched)} sheet(s) on Unmatched_Sheets tab.\n")

    run_info = {
        "Input":           str(in_path),
        "Folder Name":     folder_name,
        "Box Name":        box_name or "(inferred per-file)",
        "Run Timestamp":   datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "Files Processed": len(files),
        "Sheets Extracted":     total_emp,
        "Sheets Passed":   total_pass,
        "Sheets Failed":   total_grid - total_pass,
        "Sheets No-Grid":  total_no_grid,
        "Sheets Unmatched": len(all_unmatched),
        "Sheets Needing Review": sum(1 for r in all_qa if r["Overall"] == "REVIEW"),
        "Overall Result":  "PASS" if total_pass == total_grid and not all_unmatched else "REVIEW",
        "Script Version":  "2.1 (Phase 2)",
        "Config":          str(config_path),
    }

    id_conflicts = check_ee_id_consistency(all_data)

    data_df = pd.DataFrame(all_data, columns=output_columns)
    out_p = Path(out_path)
    out_p.parent.mkdir(parents=True, exist_ok=True)
    _write_excel(out_p, data_df, all_qa, run_info, all_unmatched, id_conflicts)
    if id_conflicts:
        print(f"  ID_Conflicts tab: {len(id_conflicts)} EE-ID conflict(s) found")
    print(f"Output written to: {out_p}")
    print("  Sheet 1 — Data             : extracted rows + QA_Flag per row")
    print("  Sheet 2 — QA_Summary       : per-employee accuracy comparison")
    print("  Sheet 3 — Run_Info         : run metadata and overall result")
    print("  Sheet 4 — Unmatched_Sheets : sheets that couldn't be safely extracted")
    print("  Sheet 5 — ID_Conflicts     : same EE ID mapped to multiple names\n")

    return 0 if total_pass == total_grid and not all_unmatched else 2


def main():
    parser = argparse.ArgumentParser(description="Extract timesheet data from Excel workbook(s).")
    parser.add_argument("input", help="Path to input .xlsx file or folder of .xlsx files")
    parser.add_argument("--folder", default="2022", help="Folder Name value (default: 2022)")
    parser.add_argument("--box",    default=None,   help="Box Name (inferred from filename if omitted)")
    parser.add_argument("--out",    default="output/extracted.xlsx", help="Output Excel path")
    parser.add_argument("--config", default="config/schema.yaml",    help="Path to config YAML")
    args = parser.parse_args()
    sys.exit(extract(args.input, args.folder, args.box, args.out, args.config))


if __name__ == "__main__":
    main()
