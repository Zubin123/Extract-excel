"""Structural primitives for locating timesheet landmarks.

Every function here is read-only and side-effect free. No hardcoded coordinates
that vary between files — those live in config/schema.yaml.
"""

from __future__ import annotations

from datetime import datetime, time
from typing import Iterable

DAY_TOKENS = {"MON", "TUE", "WED", "THUR", "THU", "FRI", "SAT", "SUN"}


def col_letter_to_index(col: str) -> int:
    result = 0
    for ch in col.upper():
        result = result * 26 + (ord(ch) - ord("A") + 1)
    return result


def col_index_to_letter(idx: int) -> str:
    s = ""
    while idx > 0:
        idx, r = divmod(idx - 1, 26)
        s = chr(ord("A") + r) + s
    return s


def is_time_value(v) -> bool:
    if isinstance(v, (datetime, time)):
        return True
    if isinstance(v, float) and 0.0 <= v < 1.0:
        return True
    return False


def find_day_label_row(ws, max_scan_row: int = 20) -> tuple[int, int] | None:
    """Return (row, first_col_idx) of the row containing >=3 day tokens.

    Returns None if no such row found within max_scan_row.
    """
    max_c = min(ws.max_column or 0, 80)
    max_r = min(ws.max_row or 0, max_scan_row)
    for r in range(1, max_r + 1):
        first_col = None
        hits = 0
        for c in range(1, max_c + 1):
            v = ws.cell(row=r, column=c).value
            if isinstance(v, str) and v.strip().upper() in DAY_TOKENS:
                hits += 1
                if first_col is None:
                    first_col = c
        if hits >= 3 and first_col is not None:
            return r, first_col
    return None


def looks_like_person_name(v) -> bool:
    """Heuristic: 'Last; First M' or 'Last, First M'."""
    if not isinstance(v, str):
        return False
    s = v.strip()
    if len(s) < 3:
        return False
    if (";" in s or "," in s) and any(c.isalpha() for c in s):
        # at least two letter chunks
        chunks = [p for p in s.replace(";", ",").split(",") if any(ch.isalpha() for ch in p)]
        return len(chunks) >= 2
    return False


def classify_sheet(ws, profiles: dict) -> str:
    """Return one of: 'employee', 'reference'.

    A sheet is an employee sheet iff A1 looks like a person name AND B2 is an
    integer ID. The presence of a weekly day-label grid is NOT required —
    sheets like 'Job' / 'OVERHEAD- CA' that name an employee but carry no
    weekly time grid are still employee sheets (they just produce blank rows).
    """
    a1 = ws["A1"].value
    if not looks_like_person_name(a1):
        return "reference"
    b2 = ws["B2"].value
    if not isinstance(b2, int) and not (isinstance(b2, float) and b2.is_integer()):
        return "reference"
    return "employee"


def sheet_has_day_grid(ws, profiles: dict) -> bool:
    """True iff the sheet has the standard weekly day-label grid (MON/TUE/...)."""
    return find_day_label_row(ws) is not None


def select_profile(ws, profiles: dict) -> str | None:
    """Pick the profile whose detect rules match this sheet."""
    day = find_day_label_row(ws)
    if day is None:
        return None
    day_row, first_col_idx = day
    first_col_letter = col_index_to_letter(first_col_idx)

    for name, prof in profiles.items():
        d = prof.get("detect", {})
        if d.get("day_label_row") != day_row:
            continue
        if d.get("day_label_first_col", "").upper() != first_col_letter:
            continue
        return name
    return None


def _block_sum(ws, row: int, start_col_idx: int, width: int) -> float:
    total = 0.0
    for k in range(width):
        v = ws.cell(row=row, column=start_col_idx + k).value
        if isinstance(v, (int, float)) and not (isinstance(v, float) and 0.0 <= v < 1.0):
            # exclude time-fractions
            total += float(v)
    return total


def _pay_type_sum_across_days(
    ws, row: int, day_start_cols: list[int], pay_type_offset: int
) -> tuple[float, bool]:
    """Sum one pay-type column across the 7 days on the given row.

    Pay-type cells can carry sub-hour values (e.g. 0.5h) so we do NOT filter
    out the 0<=v<1 range here — that filter is only appropriate for time-of-day
    cells (rows 5-8), never for hour-quantity cells.

    Returns (sum, has_any_numeric).
    """
    total = 0.0
    has_num = False
    for c in day_start_cols:
        v = ws.cell(row=row, column=c + pay_type_offset).value
        if isinstance(v, (int, float)):
            total += float(v)
            has_num = True
    return total, has_num


def resolve_pay_totals_row(
    ws,
    profile: dict,
    tolerance: float = 0.01,
    scan_from: int = 10,
    scan_to: int = 50,
) -> tuple[int | None, str]:
    """Find the row whose day-block sums equal the grand-totals row sums.

    The grand totals (BB..BG in standard profile) sit on the same row as the
    daily totals. So we look for a row r where, for every pay type i:
        sum(ws[day_start_col + i, r] for day_start_col in day_start_cols)
            ≈ ws[grand_total_cols[i], r]

    Returns (row_index, reason). row_index is None when no row satisfies the
    cross-check — the sheet should then be flagged.
    """
    db = profile["day_blocks"]
    gt_cols = profile["grand_totals"]["cols"]
    width = db["pay_block_width"]
    n_pay = len(db["pay_types"])

    day_start_idx = [col_letter_to_index(c) for c in db["day_start_cols"]]
    gt_idx = [col_letter_to_index(c) for c in gt_cols]

    max_r = min(ws.max_row or 0, scan_to)
    matching: list[tuple[int, float]] = []   # (row, total_magnitude) — pick largest
    near_miss: list[tuple[int, int]] = []    # (row, mismatched_pay_types)

    for r in range(scan_from, max_r + 1):
        gt_numeric_count = 0
        mismatches = 0
        total_magnitude = 0.0
        for i in range(n_pay):
            day_sum, _has_day = _pay_type_sum_across_days(ws, r, day_start_idx, i)
            gt_val = ws.cell(row=r, column=gt_idx[i]).value
            gt_num = isinstance(gt_val, (int, float))
            if gt_num:
                gt_numeric_count += 1
                if abs(day_sum - float(gt_val)) > tolerance:
                    mismatches += 1
                else:
                    total_magnitude += abs(float(gt_val))

        # A valid totals row has all grand-total cells populated and all cross-check.
        if gt_numeric_count == n_pay and mismatches == 0:
            matching.append((r, total_magnitude))
        elif gt_numeric_count >= n_pay - 1:
            near_miss.append((r, mismatches))

    if matching:
        # Prefer the row with the largest total magnitude — that's the row that
        # actually carries the weekly totals, not an empty upstream summary row.
        # Ties go to the latest row (typical totals row is below summary rows).
        best_row = max(matching, key=lambda x: (x[1], x[0]))[0]
        return best_row, "OK"

    if near_miss:
        best = min(near_miss, key=lambda x: x[1])
        return None, (
            f"no row passed sum cross-check; closest candidate row {best[0]} "
            f"with {best[1]} pay-type mismatch(es)"
        )
    return None, "no row had both day-block and grand-total numeric values"
