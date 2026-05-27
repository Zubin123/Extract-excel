# Next-Phase Plan — Scaling Beyond Hard-Coded Profiles

> Captured 2026-05-26 during review of the 40-workbook corpus. Goal: move
> from layout-per-template hard-coding to a self-describing extractor that
> handles 5,000+ files across template families we haven't seen yet.

## Problem statement

The current extractor encodes layouts as YAML profiles (`standard`, `shifted`)
and resolves a handful of fields by header label. Every new template family
(e.g. the **Field Mechanic** template discovered during this review —
MONDAY/TUESDAY at row 3, columns N/S/X/AC/AH/AM/AR, stride 5 not 6) requires:

1. Manual inspection of a representative sheet
2. A new YAML profile entry
3. Possibly a code change (e.g. day-token vocabulary)
4. A new release

At 5,000+ files / multiple years / multiple companies, this is not viable.
The user can guarantee templates are "similar but not the same cell
reference" — so cells will drift, but the structural pattern (a weekly
grid with day labels, time labels, pay types, grand totals) holds.

## Confirmed gaps as of this review

- 29 NO_GRID sheets are actually real Field Mechanic timesheets, currently
  emitted as 8 blank rows. ~15 employees × ~4 weeks of work silently zeroed.
- WC State / ST resolver works for the standard family but is a half-step;
  every other anchor (EE ID, name, date row, time rows, pay types) is still
  positional.
- The arithmetic cross-check (day-sum vs grand-total) is sound and
  self-validating — this is the strongest correctness gate we have and
  must be preserved through any refactor.

## Recommended architecture — fully structural layout discovery

Stop hard-coding any cell positions. Detect every landmark by content shape,
then validate via arithmetic. The vocabulary (day-name patterns, time-label
patterns, pay-type tokens) is small, bounded, and separate from layout.

### Landmark detection rules

| Landmark | Detection (no cell hard-coding) |
|---|---|
| Day-label row | Scan first 30 rows; pick the row with ≥5 cells matching any day-name pattern (MON/MONDAY/Mon/etc., regex, case-insensitive). Record exact columns where each day sits. |
| Day-block stride | Compute the gap between consecutive day-label columns (5 today, 6 today — could be anything). |
| Time-label column | Find a column where ≥3 cells say `Start` / `Lunch Out` / `Lunch In` / `Stop` (or variants) within ~6 rows of the day-row. |
| Time-data rows | Rows where the time-label column has those labels. |
| Pay-type columns | Look in the row above/below the day-label for cells matching `RT`, `OT`, `DT`, `PD`, `4D`, `4A`. Order = whatever's found. Width = count. |
| Grand-totals columns | The pay-type-count columns immediately right of the last day block. Verify by sum-cross-check. |
| Pay-totals row | Discovered by sum-cross-check (already implemented and self-validating — keep). |
| WC State / ST / EE ID / name | Resolve by header label search (already done for WC/ST; extend to all fields). |

### Phased delivery (~1 week)

1. **Phase 1 (1 day)** — Day-token vocabulary expansion (MON+MONDAY+Mon+
   localized variants) + stride-aware grid finder + dynamic day-start
   column discovery. Removes the need for `day_label_first_col` and
   `day_start_cols` in YAML. **First measurable win: the 29 Field
   Mechanic sheets extract correctly.**

2. **Phase 2 (1 day)** — Extend header-label resolution to every anchor.
   YAML moves from `cell: G13` to `near_header: "WC STATE"`. Keep cell
   hints only as last-resort fallback.

3. **Phase 3 (1 day)** — Roster cross-check via the `Employees` sheet
   present in every workbook. EE ID → name / WC State / ST becomes
   independent ground truth. Mismatch promotes from `[CHECK]` to `FAIL`.

4. **Phase 4 (½ day)** — Vocabulary YAML: accepted day-name patterns,
   time-label patterns, pay-type tokens. Operator-editable, no code
   changes for vocabulary updates.

5. **Phase 5 (½ day)** — Operator's guide + per-run sanity report:
   - "X sheets passed all gates"
   - "Y sheets had landmarks at unusual positions (listed for spot-check)"
   - "Z sheets failed → Unmatched_Sheets tab; review required"

### What this does NOT promise

- Zero human involvement. The goal is **bounded and auditable** human
  involvement — spot-check the Unmatched_Sheets tab once per run, not
  inspect every file.
- 100% accuracy on Day 1 against unseen templates. A genuinely new
  template will fail loudly (good) rather than silently emit wrong data.
- A replacement for the arithmetic cross-check. That stays as the
  ultimate correctness gate.

## Open questions for manager

1. What does "similar but not same cell reference" actually cover?
   - Whole-grid column shifts only?
   - Different templates per company/department?
   - Template revisions over time?
   - All of the above? (probable answer)

2. Is the `Employees` roster sheet reliably present in every workbook?
   If yes, it's our single best cross-check anchor.

3. Operator failure mode preference:
   - (a) Refuse to extract an unverifiable sheet → Unmatched_Sheets
   - (b) Extract best-effort and flag uncertainty (current behavior)
   - (c) Halt processing on first unverifiable sheet

## Resume point for tomorrow

Start with Phase 1. Concrete first step:

1. In [src/anchors.py](src/anchors.py), broaden `DAY_TOKENS` to include
   full names (MONDAY..SUNDAY). Add a `DAY_TOKEN_PATTERN` regex.
2. Replace the column-letter list in profiles with runtime-discovered
   `day_start_cols` from `find_day_label_row`.
3. Confirm by running on the 29 NO_GRID sheets — they should re-classify
   as grid-bearing employees and extract real hours.
4. Re-run sum cross-check QA — if it passes, we have a real Phase-1 win.

Field Mechanic structure to verify against once Phase 1 runs:
- Day labels: row 3, cols N S X AC AH AM AR (stride 5)
- Date row: row 2
- Time labels: col M, rows 4 5 6 7
- Pay-type columns and totals row: TBD by inspection
