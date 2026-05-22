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

## 5. Latest results (40-workbook corpus)

| Status | Count | Notes |
|---|---|---|
| PASS | **438** | grid-bearing sheets, all numbers verified |
| REVIEW | 0 | clean after Phase 2 noise cleanup (see §7) |
| NO_GRID | 23 | employee sheets without weekly time grid |
| Unmatched | 0 | nothing fell through |
| **Total output rows** | **~3,690** | 461 employees × 8 rows |

Authoritative output: [output/phase2_corpus_v2.xlsx](output/phase2_corpus_v2.xlsx).

`Case TUE RT = 8.0` and `Case TOTALS RT = 32.0` ✓ (original spec checks pass).

---

## 6. Layout reference (verified across 40 workbooks, 531 employee sheets)

**Fixed (never varies):**
- Day-label row 4, date row 3, time rows 5–8, daily Total Hours row 9
- Each day block is 6 columns wide in order `RT, OT, DT, PD, 4D, 4A`
- Grand-total columns sit immediately right of the last day block

**Variable (discovered at runtime):**
- Pay-totals row: 22 different positions observed (13, 14, 15, 17, 18, 20, 22,
  23, 24, 25, 28, 29, 30, 32, 33, 34, 37, 40, 41, 42)
- Column profile: `standard` (MON at L) covers 464 sheets; `shifted` (MON at
  M) covers 13 sheets

Reference sheets (auto-rejected by structural classifier, no name list):
`Job List`, `Equipment`, `Employees`, `Template`, `ColumnLists`, plus any
sheet where `A1` isn't a person name or `B2` isn't an int.

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

1. **`Hall` / `Smith, Dustin` files extract `emp=0`.** Their JOB / OVERHEAD-CA
   sheets have `A1=None` and `B2='EE: # '` (template placeholder). Would need
   filename-based employee inference. Not yet implemented.

2. **No parallelization.** Single-process; 5,000 files take ~4h sequential.
   Adding `multiprocessing.Pool` over workbooks is a 1-day task for ~10× speedup.

3. **Single large output workbook.** At 5,000 files, the output `.xlsx` is
   too large for Excel. Switch to CSV/Parquet output for large runs.

4. **No incremental processing.** Reruns re-extract everything. Adding
   filename+mtime hashing to skip unchanged files is straightforward.

5. **Operator's guide doesn't exist yet.** Non-engineers running at scale
   need a one-page README of "run this command, look at these tabs, do X when
   Y is flagged."

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

---

## 12. Session continuity for Claude

When resuming on a new machine, in priority order:

1. Read this file (§4–§11 especially).
2. Inspect [output/phase2_corpus_v2.xlsx](output/phase2_corpus_v2.xlsx)
   Run_Info tab — that's the latest run's metadata.
3. Inspect [output/probe_report.xlsx](output/probe_report.xlsx) for the
   structural patterns observed in the corpus.
4. Read `src/anchors.py` (the key technical artifact) and `src/qa.py`.
5. Skim `config/schema.yaml` — all profile/column declarations live there;
   the Python code is driven from it.

Git history has the chronological story: Phase 1 (anchor-based extraction),
Phase 2 (layered QA), Phase 2 v2 (noise cleanup — `[INFO]` downgrades,
"OK (no time data)" relabel).
