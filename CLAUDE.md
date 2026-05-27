# CLAUDE.md — Project Instructions & Handoff Notes

A robust, config-driven Python pipeline that extracts weekly timesheet data
from multi-sheet Excel workbooks into a flat tabular output with layered QA.

> **This file is the source of truth for the project.** Read it end-to-end
> when picking up the project on a new machine or in a new Claude session.

---

## 1. What this project does

- **Input:** any number of weekly-timesheet `.xlsx` workbooks (sample folder
  or a folder of 100s/1000s of files). Each workbook may contain reference
  sheets + multiple per-employee timesheet sheets.
- **Output:** a single Excel workbook with 5 tabs (Data, QA_Summary, Run_Info,
  Unmatched_Sheets, ID_Conflicts) matching the header order of
  `data/Header Template New_Rev 1.xlsx`.
- **Goal:** near-100% extraction accuracy with self-validating checks; any
  sheet that can't be safely extracted is flagged loudly rather than producing
  silently-wrong numbers.

---

## 2. Files in this repo

```
Extract-excel/
├── CLAUDE.md                       ← this file (read first)
├── pyproject.toml                  ← deps (managed by `uv`)
├── uv.lock
├── .gitignore                      ← excludes data/sample*, most outputs, .venv
├── config/
│   └── schema.yaml                 ← profiles + output columns (edit to add variants)
├── src/
│   ├── extract.py                  ← orchestrator + CLI
│   ├── anchors.py                  ← structural primitives
│   ├── qa.py                       ← layered QA checks
│   └── probe.py                    ← dev tool to scan new file batches
├── data/
│   ├── WE 01 08 22 CA GF'S - WA 4939-0899-6263_1.xlsx   ← baseline input (tracked)
│   ├── Header Template New_Rev 1.xlsx                    ← output column template (tracked)
│   ├── sample/                                           ← NOT in git; ship via OneDrive
│   └── sample1/                                          ← NOT in git; ship via OneDrive
└── output/
    ├── phase2_corpus_v2.xlsx       ← latest 40-file corpus result (tracked)
    └── probe_report.xlsx           ← latest probe scan output (tracked)
```

When starting on a new machine, copy `data/sample/` and `data/sample1/` from
the shared drive into this repo's `data/` directory.

---

## 3. Quick start (new machine)

```bash
# 1. Install uv if you don't have it: https://docs.astral.sh/uv/
# 2. Create venv and install deps
uv venv
uv sync

# 3. Copy data/sample/ and data/sample1/ from the shared drive into ./data/

# 4. Run the extractor
.venv\Scripts\python.exe src\extract.py data\ --out output\run.xlsx

# 5. Or run the probe first on a new batch to check for layout drift
.venv\Scripts\python.exe src\probe.py    # writes output/probe_report.xlsx
```

**Windows / PowerShell:** call `.venv\Scripts\python.exe` directly. `uv run`
may not resolve if `uv` isn't on PATH globally.

---

## 4. How the pipeline works (high level)

### Phase 1 — Anchor-based extraction (shipped)
Layout discovery from structure, not hardcoded coordinates:

1. **Sheet classifier** — a sheet is an employee sheet iff `A1` looks like a
   person name AND `B2` is an integer EE ID. No reference-sheet name list.
2. **Profile detection** — find `MON` in row 4; the column determines the
   layout (`standard` at L vs `shifted` at M). Profiles declared in
   `config/schema.yaml`.
3. **Pay-totals row resolution** — scan rows 10–50; the right row is the one
   where the 6 day-block pay-type sums equal the grand totals in BB..BG (or
   BC..BH for shifted). Prefer the row with the largest non-zero magnitude
   so we don't pick decoy all-zero summary rows. **Self-validating.**
4. **No-grid employee sheets** — `Job`, `OVERHEAD- CA`, `Girtman - CA`, etc.
   carry a real EE name+ID but no weekly grid. Each produces 8 placeholder
   rows (MON–SUN + TOTALS) with `QA_Flag = "no time grid"`.

### Phase 2 — Layered QA (shipped)
Each layer surfaces a specific class of issue without blocking extraction:

| Layer | Severity | Purpose |
|---|---|---|
| Per-pay-type daily sum vs grand total | fatal if mismatch | Authoritative; if this fails, extraction is wrong |
| Time order (Start ≤ LO ≤ LI ≤ Stop) | `[CHECK]` | Swapped time cells |
| Weekly hours plausibility (0–140h) | `[WARN]` | Typos / unit errors |
| Date sequence + day-of-week alignment | `[CHECK]` | Off-by-day errors |
| EE-ID consistency per workbook | tab | Same EE ID with different names |
| Filename WE date vs MON in sheet | `[INFO]` | Stale template carry-overs |

Severity legend in QA_Flag strings:
- `[INFO]` — informational; extraction is fine, ignore unless auditing
- `[WARN]` — unusual, glance at it
- `[CHECK]` — two independent signals disagree, review
- no prefix — extraction itself may be wrong (e.g. `#VALUE!`, OFF in time cell)

### Sheet status (Overall column in QA_Summary)
- `PASS` — daily sums match grand totals AND no `[CHECK]`/`[WARN]` flag
- `REVIEW` — daily sums match BUT a `[CHECK]`/`[WARN]` flag is present
- `FAIL` — daily sums don't match grand totals (extraction is wrong)
- `NO_GRID` — no-grid employee sheet (8 blank rows)

### Output tabs
- **Data** — row-per-day (7 + 1 TOTALS per employee × all employees)
- **QA_Summary** — one row per employee sheet; pay-type-by-pay-type comparison
- **Run_Info** — run metadata, totals, overall result
- **Unmatched_Sheets** — sheets the classifier accepted but extraction couldn't safely complete
- **ID_Conflicts** — same EE ID → multiple names within one workbook

---

## 5. Latest results & current state (as of 2026-05-27, label-driven Phase 4 in progress)

**Guiding principle (carry forward):** extract every output field by **finding its
label**, regardless of template family. Do not branch logic on "is this Field
Mechanic vs standard." If a field's label is present → extract it; if absent →
leave blank + emit a `[CHECK]` flag. Only hardcode logic that is a 100%
structural guarantee (the arithmetic sum-cross-check; invariants confirmed
across the whole corpus). Never emit a guessed value.

Authoritative output: `output/phase2_corpus_v4.xlsx` (local-only, gitignored).

### Current v4 result (after the empty-employee drop filter, commit a90cd9c)
| Status | Count | Notes |
|---|---|---|
| PASS | ~251 | grid-bearing, pay-type sums verified, no `[CHECK]` |
| REVIEW | ~20 | extracted correctly but flagged |
| Dropped (empty) | 218 | zero-activity employees removed from output |
| Unmatched | 41 | unfilled templates — visible, not silently dropped |
| **In progress** | 29 Field Mechanic sheets | real timesheets, a 3rd family — see §6.3 |

The 271 sheets currently in v4 are independently audited correct
(`scripts/audit_v4.py`): pay-type grand totals 271/271, WC State 271/271, ST
271/271, Total Hours effectively 271/271 (64 are `#VALUE!` in source, recovered
by summing day rows).

**History that shaped this:** the v2 baseline (438/438 PASS) was *silently wrong*
on 13 sheets — WC State read `'N'` from the FSL column in WE 12 25 21 6818
because anchors were positional. Header-verified anchors (§6.1) fixed that.
All 208 annotations in `phase2_corpus_v3 - Feedback.xlsx` were resolved
(WC State / ST missing → header resolver; Employee Name from sheet-name → now
blank+flag; Hall all-fields-missing → `employee_alt_layout`). `Case TUE RT =
8.0` and `Case TOTALS RT = 32.0` ✓.

---

## 6. Layout reference (verified across 40 workbooks, 531 employee sheets)

**Fixed (never varies):**
- Day-label row 4, date row 3, time rows 5–8, daily Total Hours row 9
- Each day block is 6 columns wide in order `RT, OT, DT, PD, 4D, 4A`
- Grand-total columns sit immediately right of the last day block

**Variable (discovered at runtime):**
- Pay-totals row: 22 different positions observed (13, 14, 15, 17, 18, 20, 22,
  23, 24, 25, 28, 29, 30, 32, 33, 34, 37, 40, 41, 42)
- **THREE template families** (not two — corrects an earlier claim):
  `standard` (MON at L, ~460 sheets), `shifted` (MON at M, 13 sheets), and
  `Field Mechanic` (full day names at row 3, 29 sheets — see §6.3)
- **WC State / ST column positions** drift across templates (see §6.1)
- **Sheet header layout** varies (A1/B2 standard vs B1/C2 Hall-style — see §6.2)

### 6.1 Header-verified anchors (Phase 3)

`resolve_anchor_cell` in `src/anchors.py` finds fields by **header label**, not by fixed cell position:

1. Reads the configured `header_row` ± 1 looking for the `header_label` text (case-insensitive)
2. Once found, scans downward up to 5 rows for the first non-empty cell
3. Returns `(value, status, addr)` where status is `"ok"`, `"relocated:<addr>"`, `"missing"`, or `"noverify"`

The config YAML (`config/schema.yaml`) declares anchors like:
```yaml
wc_state:
  cell: G13                # expected location (fallback)
  header_row: 12           # where to find the label
  header_label: WC STATE   # text to match
```

This handles three real-world drifts seen in the corpus:
- **Column drift** — FSL inserted before WC STATE shifts the column from G to H (WE 12 25 21 6818, 13 sheets)
- **Header row drift** — OVERHEAD-CA style sheets have headers at row 11 instead of 12 (Hall + single-employee Job/OVERHEAD sheets, ~14 sheets)
- **Data row drift** — Ibarra/Schnider OH have header at row 12 but the value sits at row 15 (~6 sheets)

Any time the resolver falls back, a `[CHECK]` flag is emitted on the TOTALS row's QA_Flag and the sheet's QA_Summary Overall becomes `REVIEW`.

### 6.2 Sheet classifier categories (Phase 3)

`classify_sheet` returns `(kind, reason)`. Each kind routes through a different extract path:

| kind | Detection | Where it goes |
|---|---|---|
| `employee` | `A1` is a person name AND `B2` is an int EE ID | normal extraction |
| `employee_placeholder` | `B2` is an int EE ID but `A1` is `"Input Employee Number"` or empty | extracted; Employee Name left blank with `[CHECK]` flag |
| `employee_alt_layout` | `A1` is empty AND `B1` is a person name AND `C2` is an int EE ID (Hall-style template) | extracted using B1 for name, C2 for EE ID |
| `unfilled_template` | `B2` carries `'EE: # '` template literal AND `A1` is None | logged to Unmatched_Sheets with reason |
| `reference` | anything else (Job List, Equipment, ColumnLists, etc.) | silently skipped |

The classifier does NOT use a hardcoded list of reference-sheet names — it's purely structural.

### 6.3 Field Mechanic family (3rd template — in-progress extraction)

29 sheets across ~14 workbooks (Coulson, Pottage, Beeler, Medina, Hall, Smith
Dustin, Girtman). Title cell `Z1 = "FIELD MECHANIC TIME SHEET"`. **Structurally
different — a line-item table, not a day-block grid.** Verified by raw-grid
inspection (`scripts/inspect_fm_full.py`, `scripts/probe_fm_paytypes.py`).

Two stacked grids on one sheet:
1. **Clock-time summary**, rows 4–10: time labels in col M (Start=4, Lunch
   Out=5, Lunch In=6, Stop=7), `Total Hours`=row 8. Day columns N,S,X,AC,AH,AM,AR
   (stride 5) for MON..SUN; week total in AW.
2. **Pay-coding line-item table**: header at **row 11** (found by label —
   `WC STATE` at L11, `ST` at M11, then **7 days × 5 pay-types interleaved**:
   `RT1 OT1 DT1 PTO1 HP1 | RT2 ... | ... RT7 ... HP7`). Each row 12+ is a
   job/cost-code **line item**, not a day. Per-line total in `BB`.

Key differences from standard:
- Pay-types are **RT / OT / DT / PTO / HP** (not RT/OT/DT/PD/4D/4A). **PTO & HP
  are real here** and map to the previously-always-blank PTO/HP output columns.
- Day labels are **full names** (MONDAY..SUNDAY) at **row 3** (not MON/TUE).
- Per-day pay-type totals are recoverable by summing each labelled `(day,
  pay-type)` column down the line items (35 = 7×5 columns found by label on
  every sheet).
- Identity: Coulson/Pottage use A1/B2; Hall uses alt-layout B1/C2 (A1 empty).

The manager's `phase2_corpus_v4 - Feedback.xlsx` deliberately had **no expected
rows** for these — they were dropped until extraction was designed. The Phase-4
label-driven path extracts them by discovering every field by label (no
per-family branching), validated by arithmetic.

---

## 7. Pay-type semantics (important for any future QA work)

The pay model is **not uniform** across companies/sheets in this dataset:

- `RT` = Regular Time. Hours worked at standard rate.
- `OT` = Overtime. **Sometimes a premium on already-counted hours (so
  Total Hours = RT + DT), sometimes additive (Total Hours = RT + OT + DT).**
  This varies by company; do NOT assume a single formula.
- `DT` = Double Time. Hours worked at 2× rate.
- `PD` = Per Diem. Allowance, not always hours.
- `4D`, `4A` = labelled allowance categories. Treat as opaque numbers.

An earlier RT+DT vs Total Hours two-pass check produced false positives and
was removed. Phase 1's per-pay-type sum-cross-check is the only universal
invariant — that's the check `Overall=PASS` is built on.

---

## 8. Adding a new template variant (no code changes needed)

If a new layout shows up (e.g. MON starts at column N), add a profile to
`config/schema.yaml`:

```yaml
profiles:
  alt_profile:
    detect:
      day_label_row: 4
      day_label_first_col: N
    day_label_row: 4
    date_row: 3
    time_rows: {Start: 5, "Lunch Out": 6, "Lunch In": 7, Stop: 8}
    total_hours_row: 9
    day_blocks:
      day_labels:     [MON, TUE, WED, THUR, FRI, SAT, SUN]
      day_start_cols: [N,   T,   Z,   AF,   AL,  AR,  AX]
      pay_block_width: 6
      pay_types: [RT, OT, DT, PD, 4D, 4A]
    grand_totals:
      cols:            [BD, BE, BF, BG, BH, BI]
      total_hours_col: BD
    anchors:
      employee_name: A1
      ee_id:         B2
      wc_state:      G13
      st:            H13
```

Run `python src/probe.py` first to confirm the new column letters.

---

## 9. Outstanding work / known limitations

### Resolved in Phase 3 (commit 64367e3)
- ✅ **Hall files extract `emp=0`** — fixed via `employee_alt_layout` classifier; B1 holds the name, C2 holds the EE ID
- ✅ **Silent skip of sheets with `A1='Input Employee Number'`** — fixed via `employee_placeholder` classifier (Delgado/Griego/Nunez/McDonald/Benedict/Rivera/Hanson now extract correctly)
- ✅ **WC State read from FSL column on WE 12 25 21 6818** — fixed via header-verified `resolve_anchor_cell` (was 104 silently wrong rows, now 0)
- ✅ **Unmatched_Sheets visibility** — sheets previously dropped silently are now logged with reasons

### Still open
1. **No parallelization.** Single-process; 5,000 files take ~4h sequential.
   Adding `multiprocessing.Pool` over workbooks is a 1-day task for ~10× speedup.

2. **Single large output workbook.** At 5,000 files, the output `.xlsx` is
   too large for Excel. Switch to CSV/Parquet output for large runs.

3. **No incremental processing.** Reruns re-extract everything. Adding
   filename+mtime hashing to skip unchanged files is straightforward.

4. **Operator's guide doesn't exist yet.** Non-engineers running at scale
   need a one-page README of "run this command, look at these tabs, do X when
   Y is flagged."

5. **Other free-floating field anchors still positional.** Header verification
   is currently applied only to `wc_state` and `st`. The same pattern should
   be extended to `employee_name`, `ee_id`, date row, time rows, and pay-type
   column order to close the entire "silent position drift" failure class.
   See conversation thread on residual risks for details. ~1 day of work.

6. **Hall ST value depends on header presence in row 11.** The Hall template
   has WC STATE at L11 but no ST header. ST currently resolves to whatever
   the resolver finds when scanning H12 ± 1 — which happens to be 'CA' in
   the corpus. If a Hall file ever has a different ST value, we'd miss it.
   Roster lookup via the workbook's Employees sheet (EE 4876 → ST='CA') would
   be the robust fix.

---

## 10. How to extend (good defaults)

- **Adding an output column:** add to `output_columns` in
  `config/schema.yaml`, then populate that key in the row dict in
  `extract.py` (both `extract_sheet` and `extract_sheet_no_grid`).
- **Adding a QA check:** write a small function in `src/qa.py` returning
  `list[str]` of issue messages. Call it from `process_workbook` in
  `extract.py`. Use the `[INFO]/[WARN]/[CHECK]` severity prefix convention.
- **Adding a layout profile:** edit `config/schema.yaml`. No Python changes.
- **Debugging a specific file:** copy into `data/`, run
  `python src/extract.py "data\<file>.xlsx" --out output\debug.xlsx`, then
  inspect the output's Data + QA_Summary tabs.

---

## 11. Rules (carry forward)

- **No LLM at runtime.** Pure Python; deterministic.
- **Never silently extract wrong data.** If a row/sheet can't be safely
  resolved, flag it on Unmatched_Sheets or with a non-OK QA_Flag.
- **Pay-type sub-hour values are real hours.** Never filter `0 <= v < 1` on
  pay-type cells (the time-fraction filter belongs only on time-of-day cells,
  rows 5–8).
- **Preserve trailing spaces in sheet names.** They're load-bearing.
- **Don't hand-edit `pyproject.toml`** — use `uv add` / `uv remove`.
- **Don't modify** `data/Header Template New_Rev 1.xlsx`. Read-only reference.
- **Don't ignore empty employee sheets.** If `A1` is a name and `B2` is an
  ID, the sheet must produce 8 output rows even if the employee had zero
  hours that week.
- **Reference sheets are detected structurally, not by name list.**
- **When adding a new field anchor, use `resolve_anchor_cell` with a header label** — don't read fixed cells like `G13` directly. Add `cell` + `header_row` + `header_label` to the YAML profile.
- **`employee_placeholder` sheets must leave Employee Name blank** — never use the sheet tab name as a fallback. The sheet name isn't reliably the employee's real name.

---

## 12. Session continuity for Claude

**Working directory:** `C:\dev\Extract-excel` (NOT the OneDrive folder — OneDrive corrupts git pack files; project was moved here in Phase 3).

When resuming on a new machine, in priority order:

1. Read this file (§4–§11 especially). The Phase 3 changes in §5, §6.1, §6.2, §9 are the most recent.
2. Inspect `output/phase2_corpus_v4.xlsx` Run_Info tab — that's the latest run's metadata (local-only, gitignored).
3. Inspect `output/probe_report.xlsx` for the structural patterns observed in the corpus.
4. Read `src/anchors.py` (the key technical artifact — `resolve_anchor_cell` and `classify_sheet` are the two most important functions) and `src/qa.py`.
5. Skim `config/schema.yaml` — all profile/column declarations live there; the Python code is driven from it. Anchors now use `{cell, header_row, header_label}` dict form for wc_state and st.

**Git state at last update:**
- Branch: `main`
- HEAD: `64367e3` "Fix layout-drift bugs: header-verified anchors + Hall alt-layout + placeholder handling"
- Remote: https://github.com/Zubin123/Extract-excel.git

**Conversation history that shaped Phase 3:**
- User reported 6 specific claims about missing/wrong extractions in early WE 12 18 21 + WE 12 25 21 corpus
- Investigation found root causes were positional anchors (G13/H13 hardcoded), placeholder sheets being silently dropped, and Hall files using a different template (B1/C2 instead of A1/B2)
- Fix was header-verified `resolve_anchor_cell` + expanded classifier (4 kinds instead of 2)
- All 6 claims and all 208 feedback-file annotations were resolved (verified via `scripts/validate_against_feedback.py`)

Git history has the chronological story: Phase 1 (anchor-based extraction),
Phase 2 (layered QA), Phase 2 v2 (noise cleanup — `[INFO]` downgrades,
"OK (no time data)" relabel).
