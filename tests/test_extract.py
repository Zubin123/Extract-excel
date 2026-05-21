"""Tests for extract.py."""

import sys
from pathlib import Path

import pandas as pd
import pytest

sys.path.insert(0, str(Path(__file__).parent.parent / "src"))

from extract import check_accuracy, extract, extract_sheet, infer_box_name, REFERENCE_SHEETS

INPUT_FILE = Path(__file__).parent.parent / "data" / "WE 01 08 22 CA GF'S - WA 4939-0899-6263_1.xlsx"


@pytest.fixture(scope="module")
def output_wb(tmp_path_factory):
    """Run extraction once into a temp file; return all three sheets as DataFrames."""
    out = tmp_path_factory.mktemp("output") / "extracted.xlsx"
    extract(str(INPUT_FILE), "2022", "WE 01 08 22", str(out))
    data = pd.read_excel(out, sheet_name="Data")
    qa   = pd.read_excel(out, sheet_name="QA_Summary")
    info = pd.read_excel(out, sheet_name="Run_Info")
    return {"data": data, "qa": qa, "info": info, "path": out}


# ── Sheet existence ────────────────────────────────────────────────────────────

def test_three_sheets_exist(output_wb):
    import openpyxl
    wb = openpyxl.load_workbook(output_wb["path"])
    assert set(wb.sheetnames) >= {"Data", "QA_Summary", "Run_Info"}


# ── Data sheet ─────────────────────────────────────────────────────────────────

def test_row_count(output_wb):
    assert len(output_wb["data"]) == 80


def test_column_order(output_wb):
    expected = [
        "Folder Name", "Box Name", "Excel Name", "Sheet Name",
        "Employee Name", "EE ID", "WC State", "ST",
        "Date", "Day", "Start", "Lunch Out", "Lunch In", "Stop",
        "Total Hours", "RT", "OT", "DT", "PD", "4D", "4A", "PTO", "HP",
        "QA_Flag",
    ]
    assert list(output_wb["data"].columns) == expected


def test_spot_check_case_tue_rt(output_wb):
    df = output_wb["data"]
    row = df[(df["Sheet Name"].str.strip() == "Case") & (df["Day"] == "TUE")]
    assert len(row) == 1
    assert float(row["RT"].iloc[0]) == pytest.approx(8.0)


def test_spot_check_case_totals_rt(output_wb):
    df = output_wb["data"]
    row = df[(df["Sheet Name"].str.strip() == "Case") & (df["Day"] == "TOTALS")]
    assert len(row) == 1
    assert float(row["RT"].iloc[0]) == pytest.approx(32.0)


def test_qa_flag_column_known_issues(output_wb):
    """
    QA flags correctly surface real source-data issues found in the sample:
      - Case MON / Schmidt MON: Start cell L5 = 'OFF' (text in a time cell)
      - MacKenzie SUN: no time data (employee off that day)
      - Covey all 7 days: no time data (employee had no hours this week)
      - Case/Schmidt TOTALS: BB9 was #VALUE!, recomputed from day rows
    """
    df = output_wb["data"]

    # Spot-check the specific cell+value flags
    case_mon = df[(df["Employee Name"] == "Case; Kory V") & (df["Day"] == "MON")]
    assert case_mon["QA_Flag"].iloc[0] == "Start cell L5 = 'OFF' — not a time"

    schmidt_mon = df[(df["Employee Name"] == "Schmidt; Eric M") & (df["Day"] == "MON")]
    assert schmidt_mon["QA_Flag"].iloc[0] == "Start cell L5 = 'OFF' — not a time"

    covey = df[(df["Employee Name"] == "Covey; Lawrence Peter") & (df["Day"] != "TOTALS")]
    assert (covey["QA_Flag"] == "no time data").all()

    case_totals = df[(df["Employee Name"] == "Case; Kory V") & (df["Day"] == "TOTALS")]
    assert case_totals["QA_Flag"].iloc[0] == "Total Hours cell BB9 = '#VALUE!' — recomputed from day rows"


# ── QA_Summary sheet ───────────────────────────────────────────────────────────

def test_all_employees_pass_qa_summary(output_wb):
    qa = output_wb["qa"]
    failed = qa[qa["Overall"] != "PASS"]
    assert len(failed) == 0, f"Employees failed: {failed['Employee'].tolist()}"


def test_qa_summary_has_ten_rows(output_wb):
    assert len(output_wb["qa"]) == 10


def test_qa_summary_match_columns_present(output_wb):
    for col in ["RT_match", "OT_match", "DT_match", "PD_match", "4D_match", "4A_match"]:
        assert col in output_wb["qa"].columns


def test_qa_summary_all_match(output_wb):
    qa = output_wb["qa"]
    for col in ["RT_match", "OT_match", "DT_match", "PD_match", "4D_match", "4A_match"]:
        mismatches = qa[~qa[col].isin(["MATCH", "SKIPPED (#VALUE in source)"])]
        assert len(mismatches) == 0, f"{col} mismatches: {mismatches['Employee'].tolist()}"


def test_qa_summary_has_total_hours_check(output_wb):
    """Total Hours should be cross-checked too."""
    qa = output_wb["qa"]
    assert "Total Hours_daily_sum"   in qa.columns
    assert "Total Hours_grand_total" in qa.columns
    assert "Total Hours_match"       in qa.columns

    # Sanders: clean week, sum should equal BB9
    sanders = qa[qa["Employee"] == "Sanders; Travis P."].iloc[0]
    assert sanders["Total Hours_match"] == "MATCH"
    assert float(sanders["Total Hours_daily_sum"]) == pytest.approx(111.0)

    # Case: has #VALUE! in MON, so the check is SKIPPED — not failed
    case = qa[qa["Employee"] == "Case; Kory V"].iloc[0]
    assert case["Total Hours_match"] == "SKIPPED (#VALUE in source)"


def test_totals_row_total_hours_populated(output_wb):
    """TOTALS rows should now carry the grand total from BB9 (when valid)."""
    df = output_wb["data"]
    sanders_totals = df[(df["Employee Name"] == "Sanders; Travis P.") & (df["Day"] == "TOTALS")].iloc[0]
    assert float(sanders_totals["Total Hours"]) == pytest.approx(111.0)

    # Case has #VALUE! in BB9, so Total Hours is computed from valid daily rows — must be a real number
    case_totals = df[(df["Employee Name"] == "Case; Kory V") & (df["Day"] == "TOTALS")].iloc[0]
    assert not pd.isna(case_totals["Total Hours"]), "Expected computed Total Hours for Case TOTALS, got NaN"


# ── Run_Info sheet ─────────────────────────────────────────────────────────────

def test_run_info_overall_pass(output_wb):
    info = output_wb["info"]
    assert info["Overall Result"].iloc[0] == "PASS"


def test_run_info_counts(output_wb):
    info = output_wb["info"]
    assert int(info["Employees Checked"].iloc[0]) == 10
    assert int(info["Employees Passed"].iloc[0])  == 10
    assert int(info["Employees Failed"].iloc[0])  == 0


# ── Unit tests (no file I/O) ───────────────────────────────────────────────────

def test_infer_box_name_spaces():
    assert infer_box_name("WE 01 08 22 CA GF'S - WA 4939-0899-6263_1.xlsx") == "WE 01 08 22"


def test_infer_box_name_underscores():
    assert infer_box_name("WE_01_08_22_CA_GF_S.xlsx") == "WE 01 08 22"


def _mk_day(day, rt=0.0, total_hours=8.0):
    return {"Day": day, "RT": rt, "OT": 0.0, "DT": 0.0, "PD": 0.0, "4D": 0.0, "4A": 0.0,
            "Total Hours": total_hours}


def test_check_accuracy_passes():
    rows = [_mk_day(d, rt=8.0, total_hours=8.0) for d in ["MON", "TUE", "WED", "THUR"]] + \
           [_mk_day(d, rt=0.0, total_hours=0.0) for d in ["FRI", "SAT", "SUN"]]
    rows.append({"Day": "TOTALS", "RT": 32.0, "OT": 0.0, "DT": 0.0, "PD": 0.0,
                 "4D": 0.0, "4A": 0.0, "Total Hours": 32.0})
    passed, results = check_accuracy(rows)
    assert passed
    assert results["RT"] == (32.0, 32.0, "MATCH")
    assert results["Total Hours"] == (32.0, 32.0, "MATCH")


def test_check_accuracy_fails_on_mismatch():
    rows = [_mk_day(d, rt=8.0, total_hours=8.0)
            for d in ["MON", "TUE", "WED", "THUR", "FRI", "SAT", "SUN"]]
    rows.append({"Day": "TOTALS", "RT": 99.0, "OT": 0.0, "DT": 0.0, "PD": 0.0,
                 "4D": 0.0, "4A": 0.0, "Total Hours": 56.0})
    passed, results = check_accuracy(rows)
    assert not passed
    assert results["RT"][2] == "MISMATCH"


def test_check_accuracy_skips_total_hours_when_value_error():
    """When source has #VALUE!, Total Hours check is SKIPPED, not failed."""
    rows = [_mk_day("MON", rt=8.0, total_hours=None)]  # source had #VALUE!
    rows += [_mk_day(d, rt=8.0, total_hours=8.0)
             for d in ["TUE", "WED", "THUR", "FRI", "SAT", "SUN"]]
    rows.append({"Day": "TOTALS", "RT": 56.0, "OT": 0.0, "DT": 0.0, "PD": 0.0,
                 "4D": 0.0, "4A": 0.0, "Total Hours": None})
    passed, results = check_accuracy(rows)
    assert passed  # pay types still match
    assert results["Total Hours"][2] == "SKIPPED (#VALUE in source)"
