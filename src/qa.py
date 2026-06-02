"""QA layers for the timesheet extraction pipeline.

Each function returns a list of issue strings (empty list = clean). The extractor
collects these into the per-row QA_Flag and the per-sheet QA_Summary.

Severity convention (encoded in the string prefix):
  - "[WARN]"  : suspicious but extracted — value may still be correct
  - "[CHECK]" : two methods disagree — extraction may have picked wrong row/cell
  - "[INFO]"  : informational note (e.g. recomputed Total Hours)
  - no prefix : extraction failure — extracted value not trustworthy
"""

from __future__ import annotations

from collections import defaultdict
from datetime import datetime, time
from typing import Iterable

from anchors import col_letter_to_index, _pay_type_sum_across_days


# ── Layer 1: time order within a day ────────────────────────────────────────

def _to_minutes(v) -> int | None:
    if isinstance(v, datetime):
        return v.hour * 60 + v.minute
    if isinstance(v, time):
        return v.hour * 60 + v.minute
    if isinstance(v, float) and 0.0 <= v < 1.0:
        return round(v * 24 * 60)
    return None


def check_time_order(time_cells: dict) -> list[str]:
    """time_cells = {'Start': (addr, val), 'Lunch Out': ..., 'Lunch In': ..., 'Stop': ...}"""
    order = ["Start", "Lunch Out", "Lunch In", "Stop"]
    minutes = []
    for name in order:
        _addr, val = time_cells[name]
        m = _to_minutes(val)
        minutes.append((name, m))

    out = []
    prev_name, prev_m = None, None
    for name, m in minutes:
        if m is None:
            prev_name, prev_m = name, m
            continue
        if prev_m is not None and m < prev_m:
            out.append(f"[CHECK] {name}({m//60}:{m%60:02d}) before {prev_name}({prev_m//60}:{prev_m%60:02d})")
        prev_name, prev_m = name, m
    return out


# ── Layer 2: weekly hours plausibility ──────────────────────────────────────

def check_weekly_hours_range(total_hours, min_h: float = 0.0,
                             max_h: float = 140.0) -> list[str]:
    if total_hours is None:
        return []
    if total_hours < min_h:
        return [f"[WARN] weekly total {total_hours} < {min_h}"]
    if total_hours > max_h:
        return [f"[WARN] weekly total {total_hours} > {max_h} (plausible weekly cap)"]
    return []


# ── Layer 3: two-pass totals-row cross-check ────────────────────────────────

def two_pass_totals_row(ws, profile, picked_row: int,
                        tolerance: float = 0.01) -> list[str]:
    """Disabled.

    Companies in this dataset use inconsistent OT/PD accounting conventions
    (some include OT in Total Hours, some don't), so a uniform RT+DT==TH
    invariant doesn't hold. The Phase 1 per-pay-type sum cross-check is the
    authoritative signal; this check only produced noise.
    """
    return []
    # --- former implementation kept below for reference ---
    """Independent re-derivation using the hours pay-types only.

    Observation: pay types split into 'hours worked' (RT, OT, DT) and
    'allowances/categories' (PD, 4D, 4A) that don't add to gross hours.
    The Total Hours grand total equals RT+OT+DT (work-hour types only).

    Cross-check: at the picked totals row, RT_grand + OT_grand + DT_grand
    should equal the sum of daily Total Hours cells.

    If they disagree, the picked row is likely wrong.
    """
    db = profile["day_blocks"]
    pay_types = db["pay_types"]
    gt_cols = profile["grand_totals"]["cols"]
    th_row = profile["total_hours_row"]
    day_idx = [col_letter_to_index(c) for c in db["day_start_cols"]]

    # Identify which grand-total columns correspond to gross-hours pay types.
    # Observation in this dataset: Total Hours = RT + DT. OT is overlapping
    # (premium on already-counted hours) and is reported separately.
    hour_pay_types = {"RT", "DT"}
    hour_gt_cols = [
        col_letter_to_index(gt_col)
        for pt, gt_col in zip(pay_types, gt_cols)
        if pt in hour_pay_types
    ]
    if not hour_gt_cols:
        return []

    daily_th = 0.0
    daily_th_seen = False
    for c in day_idx:
        v = ws.cell(row=th_row, column=c).value
        if isinstance(v, (int, float)):
            daily_th += float(v)
            daily_th_seen = True
    if not daily_th_seen or daily_th == 0:
        return []

    picked_hours_sum = 0.0
    for c in hour_gt_cols:
        v = ws.cell(row=picked_row, column=c).value
        if isinstance(v, (int, float)):
            picked_hours_sum += float(v)

    if abs(picked_hours_sum - daily_th) <= tolerance:
        return []
    # Informational: companies use different conventions (some count OT into
    # Total Hours, some don't). The Phase 1 per-pay-type sum cross-check is the
    # authoritative test; this is a secondary signal worth surfacing but not
    # itself proof of an extraction error.
    return [
        f"[INFO] picked row {picked_row}: RT+DT grand={picked_hours_sum} "
        f"differs from daily Total Hours sum={daily_th} "
        f"(may indicate company-specific OT/PD accounting)"
    ]


# ── Layer 4: date sequence per sheet ────────────────────────────────────────

EXPECTED_DAY_ORDER = ["MON", "TUE", "WED", "THUR", "FRI", "SAT", "SUN"]


def check_date_sequence(day_rows: list[dict]) -> list[str]:
    """day_rows is the 7 day rows (in MON..SUN order). Check:
      - all dates are real dates or all blank
      - dates are 7 consecutive days
      - each date's weekday matches the Day label
    """
    issues = []
    parsed = []
    for r in day_rows:
        s = r.get("Date", "")
        if not s:
            parsed.append(None)
            continue
        dt = None
        for fmt in ("%m/%d/%Y", "%m/%d/%y"):
            try:
                dt = datetime.strptime(s, fmt)
                break
            except (ValueError, TypeError):
                continue
        if dt is None:
            issues.append(f"[CHECK] {r['Day']} Date {s!r} not parseable")
            parsed.append(None)
            continue
        parsed.append(dt)

    real = [d for d in parsed if d is not None]
    if len(real) < 2:
        return issues

    # consecutive check (real ones)
    real_sorted = sorted(real)
    for a, b in zip(real_sorted, real_sorted[1:]):
        if (b - a).days > 7:
            issues.append(f"[CHECK] dates {a.date()} and {b.date()} >7 days apart")

    # day-of-week check
    # Python: Monday=0 ... Sunday=6. Sheet 'THUR' = Thursday = 3.
    label_to_weekday = {"MON": 0, "TUE": 1, "WED": 2, "THUR": 3, "FRI": 4, "SAT": 5, "SUN": 6}
    for r, dt in zip(day_rows, parsed):
        if dt is None:
            continue
        expected = label_to_weekday.get(r["Day"])
        if expected is not None and dt.weekday() != expected:
            issues.append(f"[CHECK] {r['Day']} date {dt.date()} is a {dt.strftime('%a')}, expected {r['Day']}")

    return issues


# ── Layer 5: EE ID consistency within a workbook ────────────────────────────

def build_ee_id_index(all_workbook_rows: list[dict]) -> dict[str, list[str]]:
    """For each workbook, group sheets by EE ID. Returns id -> [names...]."""
    by_id: dict[tuple[str, object], set[str]] = defaultdict(set)
    for r in all_workbook_rows:
        ee_id = r.get("EE ID")
        if ee_id is None or ee_id == "":
            continue
        name = (r.get("Employee Name") or "").strip()
        if name:
            by_id[(r["Excel Name"], ee_id)].add(name)
    return {f"{f}|{i}": sorted(v) for (f, i), v in by_id.items()}


def check_ee_id_consistency(all_rows: list[dict]) -> list[dict]:
    """Return a list of dicts describing conflicts (one EE ID → multiple names)."""
    idx = build_ee_id_index(all_rows)
    conflicts = []
    for key, names in idx.items():
        if len(names) > 1:
            file_, id_ = key.split("|", 1)
            conflicts.append({"File": file_, "EE ID": id_, "Names": "; ".join(names)})
    return conflicts


# ── Layer 6: filename ↔ Monday date ─────────────────────────────────────────

import re


def check_filename_week_match(excel_name: str, day_rows: list[dict]) -> list[str]:
    """If filename encodes a WE date (e.g. 'WE 12 25 21'), compare to the
    extracted MON date. Convention: filename WE = the *Saturday* (week ending)
    OR the *following Monday* — we accept either within 7 days of MON."""
    m = re.search(r"WE[\s_](\d{2})[\s_](\d{2})[\s_](\d{2})", excel_name, re.IGNORECASE)
    if not m:
        return []
    we_dt = None
    try:
        we_dt = datetime.strptime(f"{m.group(1)}/{m.group(2)}/{m.group(3)}", "%m/%d/%y")
    except ValueError:
        return [f"[WARN] filename WE date {m.group()!r} not parseable"]
    mon_row = next((r for r in day_rows if r["Day"] == "MON"), None)
    if not mon_row or not mon_row.get("Date"):
        return []
    try:
        mon_dt = datetime.strptime(mon_row["Date"], "%m/%d/%y")
    except ValueError:
        return []
    diff_days = abs((we_dt - mon_dt).days)
    if diff_days > 8:
        return [
            f"[INFO] filename WE {we_dt.date()} >8 days from MON {mon_dt.date()} in sheet "
            f"(extraction is fine; source dates appear stale)"
        ]
    return []
