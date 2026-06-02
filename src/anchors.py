"""Structural primitives for locating timesheet landmarks.

Every function here is read-only and side-effect free. No hardcoded coordinates
that vary between files — those live in config/schema.yaml.
"""

from __future__ import annotations

import re
from datetime import datetime, time
from typing import Iterable

# Day vocabulary — short and full forms. Same field, minor label variation.
# Add new localized variants here, not in code branches.
DAY_TOKENS = {"MON", "TUE", "WED", "THUR", "THU", "FRI", "SAT", "SUN"}
DAY_TOKENS_FULL = {
    "MONDAY", "TUESDAY", "WEDNESDAY", "THURSDAY", "FRIDAY", "SATURDAY", "SUNDAY"
}
# Canonical MON..SUN order, mapping every accepted token to its weekday index.
_DAY_ORDER = ["MON", "TUE", "WED", "THUR", "FRI", "SAT", "SUN"]
_DAY_ALIASES = {
    "MON": 0, "MONDAY": 0,
    "TUE": 1, "TUESDAY": 1,
    "WED": 2, "WEDNESDAY": 2,
    "THUR": 3, "THU": 3, "THURSDAY": 3,
    "FRI": 4, "FRIDAY": 4,
    "SAT": 5, "SATURDAY": 5,
    "SUN": 6, "SUNDAY": 6,
}
ALL_DAY_TOKENS = DAY_TOKENS | DAY_TOKENS_FULL

# Pay-type header tokens followed by a day index, e.g. RT1, OT3, PTO7, 4D5.
# The trailing digit is the day number (1=first day .. 7=last day).
# Tokens cover every pay-type seen across all template families:
#   RT, OT, DT, PD, 4D, 4A (standard); PTO, HP (Field Mechanic + drift variants)
PAYTYPE_DAY_RE = re.compile(r"^(RT|OT|DT|PD|4D|4A|PTO|HP|PP)(\d)$", re.IGNORECASE)

# Same vocabulary but with the day-1 suffix optional — some templates
# (Rancho Runners) drop the trailing '1' on the first day-block column,
# writing 'PP' instead of 'PP1' while still using 'PP2'..'PP7' for later
# days. Match group 2 is None for the suppressed-suffix form, which the
# caller interprets as day index 0.
PAYTYPE_DAY_RE_OPTIONAL_SUFFIX = re.compile(
    r"^(RT|OT|DT|PD|4D|4A|PTO|HP|PP)(\d)?$", re.IGNORECASE
)


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
            if isinstance(v, str) and v.strip().upper() in ALL_DAY_TOKENS:
                hits += 1
                if first_col is None:
                    first_col = c
        if hits >= 3 and first_col is not None:
            return r, first_col
    return None


def resolve_anchor_cell(ws, spec, search_cols: int = 14,
                        row_scan_radius: int = 1, data_scan_depth: int = 5
                        ) -> tuple[object, str, str]:
    """Read an anchor cell, verifying its column against a header label.

    `spec` is either a plain cell address like 'G13' (no verification) or a
    dict with keys: cell, header_row, header_label.

    For the dict form, the column letter of `cell` is the *expected* location
    and the row digit is the *expected* data row. The resolver:
      1. Searches header_row ± row_scan_radius for the header label (handles
         OVERHEAD-style sheets where the header sits one row above/below).
      2. Inside the matching header row, finds the column whose cell equals
         the label.
      3. Scans the column downward from header_row+1 for up to
         data_scan_depth rows to find the first non-empty cell (handles
         Ibarra-style sheets where the data sits two rows below the header).

    Returns (value, status, source_addr) where status is one of:
      "ok"                — header at expected row+col, value at expected row
      "relocated:<addr>"  — header or data found at a different position
      "missing"           — label not found anywhere within scan range
      "noverify"          — spec was a plain string; no header check
    """
    if isinstance(spec, str):
        return ws[spec].value, "noverify", spec

    if not isinstance(spec, dict) or "cell" not in spec:
        return None, "missing", ""

    cell_addr = spec["cell"]
    expected_header_row = spec.get("header_row")
    header_label = spec.get("header_label")
    if expected_header_row is None or header_label is None:
        return ws[cell_addr].value, "noverify", cell_addr

    # Split cell_addr (e.g. 'G13' -> 'G', 13)
    col_letters = "".join(ch for ch in cell_addr if ch.isalpha())
    row_digits = "".join(ch for ch in cell_addr if ch.isdigit())
    expected_col_idx = col_letter_to_index(col_letters)
    expected_data_row = int(row_digits)
    label_upper = str(header_label).strip().upper()

    def _norm(v):
        return v.strip().upper() if isinstance(v, str) else None

    def _first_non_empty(col_idx, start_row, depth):
        for r in range(start_row, start_row + depth):
            v = ws.cell(row=r, column=col_idx).value
            if v is not None and not (isinstance(v, str) and v.strip() == ""):
                return r, v
        return None, None

    # Build list of header rows to try: expected first, then ±1, ±2, ...
    rows_to_try = [expected_header_row]
    for delta in range(1, row_scan_radius + 1):
        rows_to_try.extend([expected_header_row - delta, expected_header_row + delta])
    rows_to_try = [r for r in rows_to_try if r >= 1]

    max_c = min(ws.max_column or 0, search_cols)

    for hr in rows_to_try:
        # First check the expected column at this header row
        if _norm(ws.cell(row=hr, column=expected_col_idx).value) == label_upper:
            found_col = expected_col_idx
        else:
            # Scan the row for the label in any other column
            found_col = None
            for c in range(1, max_c + 1):
                if _norm(ws.cell(row=hr, column=c).value) == label_upper:
                    found_col = c
                    break
        if found_col is None:
            continue

        # Header found at (hr, found_col). Now find first non-empty data row.
        data_row_used, value = _first_non_empty(found_col, hr + 1, data_scan_depth)
        if data_row_used is None:
            # Header present but column completely empty within scan depth.
            value = None
            data_row_used = hr + 1

        relocated_addr = f"{col_index_to_letter(found_col)}{data_row_used}"
        if hr == expected_header_row and found_col == expected_col_idx and data_row_used == expected_data_row:
            return value, "ok", cell_addr
        return value, f"relocated:{relocated_addr}", relocated_addr

    return None, "missing", ""


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


EMPLOYEE_NAME_PLACEHOLDERS = {"INPUT EMPLOYEE NUMBER"}


def _b2_is_int_ee_id(b2) -> bool:
    if isinstance(b2, bool):
        return False
    if isinstance(b2, int):
        return True
    if isinstance(b2, float) and b2.is_integer():
        return True
    return False


def classify_sheet(ws, profiles: dict) -> tuple[str, str]:
    """Return (kind, reason). kind is one of:

    - 'employee'             : A1 is a person name AND B2 is an int EE ID
                               → extract normally
    - 'employee_placeholder' : B2 is an int EE ID but A1 is a template
                               placeholder (or empty) → extract; leave
                               Employee Name blank (sheet name is not the
                               real employee name)
    - 'employee_alt_layout'  : A1 is empty but B1 is a person name AND C2
                               is an int EE ID (Hall-style OVERHEAD sheets
                               where the template uses B1/C2 instead of
                               A1/B2)
    - 'unfilled_template'    : sheet has employee-template shape but neither
                               name nor ID was filled in → log to
                               Unmatched_Sheets so it's visible
    - 'reference'            : structural reference sheet (Job List,
                               Equipment, ColumnLists, etc.) → silent skip
    """
    a1 = ws["A1"].value
    b1 = ws["B1"].value
    b2 = ws["B2"].value
    c2 = ws["C2"].value
    a1_str = a1.strip() if isinstance(a1, str) else ""
    a1_upper = a1_str.upper()
    b2_str = b2.strip() if isinstance(b2, str) else ""
    b2_upper = b2_str.upper()

    if looks_like_person_name(a1) and _b2_is_int_ee_id(b2):
        return "employee", ""

    # Hall-style alternate layout: name in B1, EE ID in C2, A1 empty.
    if (a1 is None or a1_str == "") and looks_like_person_name(b1) and _b2_is_int_ee_id(c2):
        return "employee_alt_layout", (
            f"Hall-style layout: name at B1 ({b1!r}), EE ID at C2 ({c2})"
        )

    # Has a real EE ID but A1 is empty / template placeholder → recoverable.
    # Don't use sheet name as fallback — it isn't the employee name. Leave blank.
    if _b2_is_int_ee_id(b2) and (
        a1 is None
        or a1_upper in EMPLOYEE_NAME_PLACEHOLDERS
        or a1_upper.startswith("INPUT ")
    ):
        return "employee_placeholder", (
            f"A1 placeholder ({a1!r}) — Employee Name left blank"
        )

    # Looks like an unfilled employee template (B2 carries the 'EE: # ' literal —
    # placeholder syntax with ':' or '#' — not a bare header label like 'EE ID').
    if a1 is None and isinstance(b2, str) and (":" in b2 or "#" in b2) and b2_upper.startswith("EE"):
        return "unfilled_template", (
            f"A1=None, B2={b2!r} — employee template never filled in"
        )
    if a1_upper.startswith("INPUT ") and b2 is None:
        return "unfilled_template", (
            f"A1={a1!r}, B2=None — employee template never filled in"
        )

    return "reference", ""


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
    # candidate: (row, total_magnitude, day_cell_coverage, mismatches)
    candidates: list[tuple[int, float, int, int]] = []

    for r in range(scan_from, max_r + 1):
        gt_numeric_count = 0
        mismatches = 0
        total_magnitude = 0.0
        day_cell_coverage = 0
        for i in range(n_pay):
            day_sum, _has_day = _pay_type_sum_across_days(ws, r, day_start_idx, i)
            # Count populated day cells (numeric, any value incl 0) as coverage.
            for c in day_start_idx:
                v = ws.cell(row=r, column=c + i).value
                if isinstance(v, (int, float)):
                    day_cell_coverage += 1
            gt_val = ws.cell(row=r, column=gt_idx[i]).value
            gt_num = isinstance(gt_val, (int, float))
            if gt_num:
                gt_numeric_count += 1
                if abs(day_sum - float(gt_val)) > tolerance:
                    mismatches += 1
                else:
                    total_magnitude += abs(float(gt_val))

        # A row qualifies as a candidate if it has all grand-totals cells
        # populated AND at most 1 pay-type mismatch. Strict 0-mismatch rows
        # rank above 1-mismatch ones; among equal-strictness candidates, the
        # row with the most populated day cells (the SUMMING row in multi-
        # line-item sheets) wins — closes Federline / Tweed silent-zero class
        # where row 14 (one line item) was preferred over row 25 (week sum).
        if gt_numeric_count == n_pay and mismatches <= 1:
            candidates.append((r, total_magnitude, day_cell_coverage, mismatches))

    if candidates:
        # Rank by (coverage, magnitude, later row) and tolerate one mismatch on
        # the winning row. The summing row of a multi-line-item sheet always
        # has the highest coverage (every day × every pay-type), even when one
        # of its grand-totals cells disagrees due to source-data inconsistency
        # (Federline RT 32 vs grand-total 35 case). A strict 0-mismatch
        # upstream line-item row has much lower coverage and would otherwise
        # win — losing the data on days it doesn't carry.
        best = max(candidates, key=lambda x: (x[2], x[1], x[0]))
        return best[0], "OK" if best[3] == 0 else (
            f"OK (best candidate row {best[0]} has 1 pay-type cross-check miss — "
            f"source data inconsistency; picked by cell-coverage)"
        )

    return None, "no row had both day-block and grand-total numeric values within tolerance"


# ── Field Mechanic discovery (label-driven; no hardcoded cells) ───────────────
#
# Field Mechanic sheets are line-item tables, not day-block grids. Every
# landmark below is found by the label it carries, so column/row drift between
# files is tolerated. See CLAUDE.md §6.3.

def find_label_in_grid(ws, label: str, max_r: int = 30, max_c: int = 20
                       ) -> tuple[int, int] | None:
    """Return (row, col) of the first cell whose text equals `label` (case-
    insensitive, trimmed). Searches the top-left region only. None if absent."""
    target = label.strip().upper()
    rmax = min(ws.max_row or 0, max_r)
    cmax = min(ws.max_column or 0, max_c)
    for r in range(1, rmax + 1):
        for c in range(1, cmax + 1):
            v = ws.cell(row=r, column=c).value
            if isinstance(v, str) and v.strip().upper() == target:
                return r, c
    return None


def day_label_row_style(ws, max_scan_row: int = 20) -> str | None:
    """Return 'full' if the day-label row uses MONDAY..SUNDAY, 'short' if it
    uses MON..SUN, else None. This is the structural discriminator between the
    Field Mechanic family (full names, row 3) and standard/shifted (short, row
    4) — the day-name style determines the whole grid geometry.
    """
    max_c = min(ws.max_column or 0, 80)
    max_r = min(ws.max_row or 0, max_scan_row)
    for r in range(1, max_r + 1):
        short = full = 0
        for c in range(1, max_c + 1):
            v = ws.cell(row=r, column=c).value
            if isinstance(v, str):
                u = v.strip().upper()
                if u in DAY_TOKENS:
                    short += 1
                elif u in DAY_TOKENS_FULL:
                    full += 1
        if full >= 3:
            return "full"
        if short >= 3:
            return "short"
    return None


def find_field_mechanic_header_row(ws, max_r: int = 30) -> int | None:
    """The line-item header row: contains 'WC STATE' AND >=5 'RT#/OT#/...'
    pay-type-with-day tokens. Discovered by label, never hardcoded.

    Guarded by the day-name style: standard/shifted sheets ALSO carry
    'WC STATE' + 'RT1/OT1...' headers, so the pay-type signature alone is NOT
    unique. Field Mechanic is distinguished by FULL day names (MONDAY) in its
    day-label row; standard/shifted use short names (MON). Only full-name
    sheets take the Field Mechanic path.

    Returns the 1-based row index, or None if this isn't a Field Mechanic sheet.
    """
    if day_label_row_style(ws) != "full":
        return None

    rmax = min(ws.max_row or 0, max_r)
    cmax = min(ws.max_column or 0, 120)
    for r in range(1, rmax + 1):
        has_wc = False
        paytype_hits = 0
        for c in range(1, cmax + 1):
            v = ws.cell(row=r, column=c).value
            if isinstance(v, str):
                u = v.strip().upper()
                if u == "WC STATE":
                    has_wc = True
                elif PAYTYPE_DAY_RE.match(u):
                    paytype_hits += 1
        if has_wc and paytype_hits >= 5:
            return r
    return None


def discover_paytype_columns_by_label(
    ws, header_row: int, day_start_cols: list[int]
) -> tuple[dict[tuple[int, str], int], list[str]]:
    """For every column under each day-block, read its label in `header_row`
    and map it to the output pay-type field by label.

    Returns:
      grid:    {(day_idx, pay_type): column_index} where pay_type is the
               canonical output field name (RT/OT/DT/PD/4D/4A/PTO/HP).
      issues:  list of [CHECK] messages for any column whose label couldn't
               be matched (unknown tokens, missing day index, etc).

    This unifies the standard and Field Mechanic paths: each column is
    routed to the output field that matches its actual label, regardless
    of how the YAML profile is configured. Closes the silent position-
    drift class (e.g. Wirth M12='PTO1' was being read as OT positionally).

    Day indices in the returned grid are 0..n-1 corresponding to
    day_start_cols in order. We don't fix a block width — we scan from
    each day's start column up to (the next day's start column - 1) or,
    for the last day, a sensible bound.
    """
    OUTPUT_TYPES = {"RT", "OT", "DT", "PD", "4D", "4A", "PTO", "HP", "PP"}
    grid: dict[tuple[int, str], int] = {}
    issues: list[str] = []

    # Compute scan extent for each day block from the day_start_cols list.
    # For the last day, use the same stride as the previous gap so we don't
    # spill into the grand-totals summary columns (cols 54+ on standard
    # layouts) and surface false 'unknown header' flags for legitimate
    # summary labels like '1', 'PTO', 'HP', 'Total'.
    if len(day_start_cols) >= 2:
        last_stride = day_start_cols[-1] - day_start_cols[-2]
    else:
        last_stride = 6
    bounds = []
    for i, start in enumerate(day_start_cols):
        if i + 1 < len(day_start_cols):
            end = day_start_cols[i + 1] - 1
        else:
            end = start + last_stride - 1
        bounds.append((start, end))

    for day_idx, (start, end) in enumerate(bounds):
        for c in range(start, end + 1):
            v = ws.cell(row=header_row, column=c).value
            if not isinstance(v, str):
                continue
            label_raw = v.strip()
            label_up = label_raw.upper()
            m = PAYTYPE_DAY_RE.match(label_up)
            if not m:
                # Header text present but doesn't match a known pay-type label.
                # Only flag if it LOOKS like a malformed pay-type token (starts
                # with a known prefix and has trailing digits) — Morgan-style
                # 'DT22' typos in the PD2 slot ate real data without warning.
                # We can't safely route the value, but we must tell the manager.
                looks_like_paytype = False
                for prefix in ("RT", "OT", "DT", "PD", "4D", "4A", "PTO", "HP", "PP"):
                    if (label_up.startswith(prefix)
                            and len(label_up) > len(prefix)
                            and label_up[len(prefix):].isdigit()):
                        looks_like_paytype = True
                        break
                if looks_like_paytype:
                    addr = f"{col_index_to_letter(c)}{header_row}"
                    issues.append(
                        f"[CHECK] header {addr}={label_raw!r} looks like a "
                        f"pay-type label but doesn't match the day-1..7 form "
                        f"(day {day_idx+1} block) — likely typo; value will "
                        f"not be routed to any output field"
                    )
                continue
            label_pt = m.group(1).upper()
            label_day = int(m.group(2))
            # Validate the day index matches where the column sits.
            if label_day != day_idx + 1:
                issues.append(
                    f"[CHECK] header col {col_index_to_letter(c)}{header_row}={v!r} "
                    f"but it sits under day {day_idx+1} block — label/position mismatch"
                )
                continue
            if label_pt not in OUTPUT_TYPES:
                issues.append(
                    f"[CHECK] header col {col_index_to_letter(c)}{header_row}={v!r} "
                    f"— unknown pay-type token, value will not be extracted"
                )
                continue
            grid[(day_idx, label_pt)] = c
    return grid, issues


def resolve_state_per_day_from_line_items(
    ws, label: str, header_row: int, day_start_cols: list[int],
    pay_block_width: int = 6, max_scan_rows: int = 30,
) -> tuple[dict[int, object], object | None]:
    """For each day, return the WC State / ST value taken from the line-item
    row that actually carries that day's hours.

    Returns (per_day, fallback) where:
      - per_day[day_idx] = the state code (string) for line items with hours
        on that day. Missing if no line item carries hours on that day.
      - fallback = the first non-empty value in the label column, used for
        days with no hours.

    The manager's rule: each output row's WC State / ST should reflect the
    work-state of THAT day's line item, not a single per-employee value.
    Hilyard 01/19/23 worked at OR (one line item); other days at CA.
    """
    label_up = label.strip().upper()
    label_col = None
    cmax = min(ws.max_column or 0, 80)
    for c in range(1, cmax + 1):
        v = ws.cell(row=header_row, column=c).value
        if isinstance(v, str) and v.strip().upper() == label_up:
            label_col = c
            break
    if label_col is None:
        return {}, None

    last_row = min((ws.max_row or header_row) + 1, header_row + max_scan_rows)
    per_day: dict[int, object] = {}
    fallback: object | None = None

    for r in range(header_row + 1, last_row + 1):
        state_val = ws.cell(row=r, column=label_col).value
        if state_val is None or (isinstance(state_val, str) and state_val.strip() == ""):
            continue
        if fallback is None:
            fallback = state_val
        # Sum each day-block to find days with hours on this line item.
        for day_idx, start_col in enumerate(day_start_cols):
            block_total = 0.0
            for off in range(pay_block_width):
                cell = ws.cell(row=r, column=start_col + off).value
                if isinstance(cell, (int, float)) and not (
                    isinstance(cell, float) and 0.0 <= cell < 1.0
                ):
                    block_total += abs(float(cell))
            if block_total > 0 and day_idx not in per_day:
                per_day[day_idx] = state_val

    return per_day, fallback


def resolve_state_for_hours_row(
    ws, label: str, header_row: int, day_start_cols: list[int],
    max_scan_rows: int = 30,
) -> tuple[object, str]:
    """Find the value of a per-line-item state column (e.g. 'WC STATE', 'ST') by
    picking the line-item row that has the most worked hours.

    On sheets with multiple line items (Federline-style storm-week sheets), the
    WC STATE column carries one value per line item. The manager's rule:
    use the WC State of the row that actually carries worked hours, not the
    first non-empty cell. We pick the row whose day-block cells (any of
    RT/OT/DT/PD/4D/4A/PTO/HP/PP at the recognised pay-type columns) sum to the
    largest magnitude. If no row has any hours, fall back to the first non-
    empty value below the header.

    Returns (value, status). status is 'ok' for hours-driven pick, 'fallback'
    for first-non-empty fallback, 'missing' if the column wasn't found.
    """
    # 1. Locate the label column in the header row.
    label_up = label.strip().upper()
    label_col = None
    cmax = min(ws.max_column or 0, 80)
    for c in range(1, cmax + 1):
        v = ws.cell(row=header_row, column=c).value
        if isinstance(v, str) and v.strip().upper() == label_up:
            label_col = c
            break
    if label_col is None:
        return None, "missing"

    # 2. Walk line-item rows below the header, collecting (row, value, hours).
    last_row = min((ws.max_row or header_row) + 1, header_row + max_scan_rows)
    candidates: list[tuple[int, object, float]] = []
    for r in range(header_row + 1, last_row + 1):
        v = ws.cell(row=r, column=label_col).value
        if v is None or (isinstance(v, str) and v.strip() == ""):
            continue
        # Skip aggregate / summary rows where the column carries a known
        # non-state literal (rare; usually the value is the state code).
        # Sum hours across all day-block columns for this row.
        hours = 0.0
        for c in day_start_cols:
            # Scan a small window from each day-start to capture every
            # pay-type cell in the block. Stride between days bounds the scan.
            # Width = up to 8 (max pay-type variants per day).
            for off in range(8):
                cell = ws.cell(row=r, column=c + off).value
                if isinstance(cell, (int, float)) and not (
                    isinstance(cell, float) and 0.0 <= cell < 1.0
                ):
                    hours += abs(float(cell))
        candidates.append((r, v, hours))

    if not candidates:
        return None, "missing"

    # 3. Prefer the row with the most hours; if all are zero, fall back to
    #    the first non-empty value.
    with_hours = [c for c in candidates if c[2] > 0]
    if with_hours:
        best = max(with_hours, key=lambda x: (x[2], -x[0]))
        return best[1], "ok"
    return candidates[0][1], "fallback"


def find_fm_job_col(ws, header_row: int) -> int | None:
    """Locate the 'JOB' column in the Field Mechanic header row by label.

    Field Mechanic sheets have a row labelled 'JOB' in the header. Each
    real line item below carries a job code in that column; echo/import-
    shadow rows (which duplicate per-day pay-type values for accounting
    purposes) leave it blank. Filtering by this column eliminates
    double-counting without any hardcoded cell.
    """
    cmax = min(ws.max_column or 0, 30)
    for c in range(1, cmax + 1):
        v = ws.cell(row=header_row, column=c).value
        if isinstance(v, str) and v.strip().upper() == "JOB":
            return c
    return None


def discover_field_mechanic_paytype_cols(ws, header_row: int
                                         ) -> tuple[dict, list[int], list[str]]:
    """Map each (day_index, pay_type) -> column index from the header row labels.

    Returns (grid, day_indices, pay_types_seen) where:
      - grid[(day_idx, PT)] = column index (1-based); day_idx is 0..6 (MON..SUN)
        derived from the trailing digit (1->0 .. 7->6).
      - day_indices = sorted list of day_idx values present.
      - pay_types_seen = pay-type tokens encountered, in first-seen order.
    """
    grid: dict[tuple[int, str], int] = {}
    day_indices: set[int] = set()
    pay_types_seen: list[str] = []

    # Derive the day-1 column bounds from the day-label row. A bare-suffix
    # token (e.g. 'PP') is only treated as day-1 when it sits inside the
    # day-1 block — otherwise grand-totals summary labels like bare 'PTO'
    # / 'HP' / 'PP' at cols 50+ would overwrite legitimate 'PTO1'/'HP1'
    # mappings (Adam Peart, Rancho Office Brothers regression).
    day_info = find_day_label_row(ws)
    day1_start: int | None = None
    day1_end: int | None = None
    if day_info is not None:
        day_label_row, first_day_col = day_info
        # Collect day columns in row-order to compute stride.
        day_cols: list[int] = []
        cmax_day = min(ws.max_column or 0, 120)
        for cc in range(1, cmax_day + 1):
            dv = ws.cell(row=day_label_row, column=cc).value
            if isinstance(dv, str) and dv.strip().upper() in ALL_DAY_TOKENS:
                day_cols.append(cc)
        if len(day_cols) >= 2:
            day1_start = day_cols[0]
            day1_end = day_cols[1] - 1

    cmax = min(ws.max_column or 0, 120)
    for c in range(1, cmax + 1):
        v = ws.cell(row=header_row, column=c).value
        if not isinstance(v, str):
            continue
        # Accept the bare-token form (e.g. 'PP' on Rancho Runners day-1) by
        # treating missing suffix as day index 1, BUT only when the column
        # sits inside the day-1 block. Bare tokens in the grand-totals area
        # (typical at cols 47+ on standard layouts) are summary headers,
        # not per-day labels — Adam Peart / Brothers had bare 'PTO'/'HP'
        # there overriding the real PTO1/HP1 mappings.
        m = PAYTYPE_DAY_RE_OPTIONAL_SUFFIX.match(v.strip().upper())
        if not m:
            continue
        pt = m.group(1).upper()
        if m.group(2) is None:
            if day1_start is None or not (day1_start <= c <= day1_end):
                continue
            day_num = 1
        else:
            day_num = int(m.group(2))
        if not (1 <= day_num <= 7):
            continue
        day_idx = day_num - 1
        grid[(day_idx, pt)] = c
        day_indices.add(day_idx)
        if pt not in pay_types_seen:
            pay_types_seen.append(pt)
    return grid, sorted(day_indices), pay_types_seen


def day_index_to_label(day_idx: int) -> str:
    """0->MON .. 6->SUN (canonical short label for output)."""
    return _DAY_ORDER[day_idx] if 0 <= day_idx < len(_DAY_ORDER) else f"DAY{day_idx+1}"


def day_alias_index(token: str) -> int | None:
    """Map any accepted day token (MON / MONDAY / THU / ...) to 0..6, else None."""
    return _DAY_ALIASES.get(token.strip().upper()) if isinstance(token, str) else None
