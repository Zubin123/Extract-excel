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
│   ├── sample1/                                          ← NOT in git; ship via OneDrive
│   └── sample2/                                          ← NOT in git; 63 wb, 5 WE folders, 2023 weeks
├── output/
│   ├── corpus_v5.xlsx              ← sample/sample1 manager-validated (gitignored)
│   ├── corpus_sample2_v3.xlsx      ← sample2 manager-validated (gitignored)
│   ├── phase2_corpus_v2.xlsx       ← legacy 40-file corpus result (tracked, stale)
│   └── probe_report.xlsx           ← latest probe scan output (tracked)
├── corpus_sample2 - Feedback.xlsx  ← manager-annotated truth for sample2 (gitignored, in repo root)
└── phase2_corpus_v4 - Feedback.xlsx ← manager-annotated truth for sample/sample1 (gitignored)
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

## 5. Current state (2026-06-02, sample2 manager-validated)

**Guiding principle (THE rule):** extract every output field by **finding its
label** in the source. No fixed cell addresses outside structural invariants.
No per-family branching. If a field's label is present → extract it; if absent
→ leave blank + emit a `[CHECK]`. Only hardcode logic that is a 100%
structural guarantee (the arithmetic sum-cross-check; date row = day-label
row − 1; the output-schema column list). Never emit a guessed value.

**Authoritative outputs:**
- `output/corpus_v5.xlsx` — sample/sample1 (2022 source), manager-accepted.
- `output/corpus_sample2_v3.xlsx` — sample2 (2023 weeks), validated against
  `corpus_sample2 - Feedback.xlsx`. **13 of 17 fields at 100%**.

### Final result vs `corpus_sample2 - Feedback.xlsx` (annotated truth)
| Field | Match rate |
|---|---|
| WC State, ST, Date | **100%** |
| Total Hours | **100%** |
| RT, OT, DT, PD, 4D, 4A, PTO, HP, PP | **100% each** |
| Start, Lunch In | 99.6% |
| Lunch Out, Stop | 99.1% |

The Start/Lunch Out/Lunch In/Stop misses (13–29 cells per column out of 3,202)
are not data errors — they are a layout-interpretation difference on
~13 employees where the source has only Start + one mid-time cell with no
Stop, and the manager places that mid-time in the Stop column while we keep
the source positions faithfully. Not flagged in any feedback round so far;
leaving as source-faithful unless explicitly requested.

### Earlier validation vs `phase2_corpus_v4 - Feedback.xlsx` (sample/sample1)
**15 of 16 fields at 100%** (single 0.8% miss = Wirth PTO 19.98 vs manager-
rounded 20.00). Confirmed by `scripts/compare_v4_feedback.py`.

### Run shape (sample2, current)
- 63 workbooks → 741 employee sheets extracted
- 434 PASS / 307 REVIEW / 0 FAIL (most REVIEWs = informational anchor flags)
- 303 dropped by empty-employee filter
- 66 on `Unmatched_Sheets` (unfilled templates — visible)

### Independent audit (`scripts/audit_v4.py`, sample2)
- Pay-type grand totals match source: **735/741**
- WC State / ST: **733/741** — 8 mismatches are the per-day resolver picking
  the line-item-with-hours value (working as designed for Hilyard/Driessen)
- 237 source `#VALUE!` Total Hours correctly recovered
- 12 suspicious values (11 negative RT, 1 negative 4A) — faithful to source

### History (briefly)
- v2 baseline (438/438 PASS) was *silently wrong* on 13 sheets — WC State read
  `'N'` from the FSL column. Header-label anchors fixed that.
- v3 feedback (208 annotations) resolved by header-label resolver + Hall
  `employee_alt_layout` + `employee_placeholder`.
- v4 + label-driven (commits `00af7f1`, `a6bae9a`): Field Mechanic family,
  PD/4D/4A in regex, all 8 pay-types label-driven per sheet.
- **sample2 fixes (commit `9b5bc8c`, 2026-06-02)** — see §14 for the full
  list of bugs closed.
- `Case TUE RT = 8.0` and `Case TOTALS RT = 32.0` ✓.

---

## 6. Layout reference (verified across 40 workbooks, 531 employee sheets)

**Universal invariants (the only safe-to-hardcode facts):**
- Date row = day-label row − 1
- Time labels start at day-label row + 1
- The output schema (`config/schema.yaml` `output_columns`) is the contract
- Arithmetic sum-cross-check is the correctness gate (day-sums = grand-totals)

**Everything else is discovered at runtime by label:**
- Day-label row position: scan rows 1–20 for ≥3 day tokens (`MON…SUN` short or
  `MONDAY…SUNDAY` full — vocabulary lives in `anchors.ALL_DAY_TOKENS`)
- Day-block start columns: wherever each day token actually sits
- Pay-type columns (RT/OT/DT/PD/4D/4A/PTO/HP): per-sheet, by reading row 12's
  actual labels (`L12='RT1'` → RT col for day 1, etc.). See §6.4.
- Pay-totals row: found by `resolve_pay_totals_row` via sum-cross-check
  (22 distinct positions seen: 13–42)
- WC State / ST / EE ID / Name: by header-label search

**Template families (informational, NOT branched on in extraction):**
- `standard` (MON at L, ~460 sheets), `shifted` (MON at M, 13 sheets),
  `field_mechanic` (full day names at row 3, 19 with-data + 10 empty,
  see §6.3). All three go through the same label-driven extraction —
  the family label is only useful for diagnostics.
- Sheet header layout varies: A1/B2 standard, B1/C2 Hall alt-layout —
  the classifier detects this (§6.2).

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

### 6.3 Field Mechanic family (now extracted via the unified label-driven path)

19 with-data sheets across ~14 workbooks (Coulson, Pottage, Beeler, Medina,
Hall, Smith Dustin, Girtman). Title cell `Z1 = "FIELD MECHANIC TIME SHEET"`.
**Line-item table, not a day-block grid.** Verified by raw-grid inspection
(`scripts/inspect_fm_full.py`, `scripts/probe_fm_paytypes.py`).

Two stacked grids on one sheet:
1. **Clock-time summary**, rows 4–10: time labels in col M (Start=4, Lunch
   Out=5, Lunch In=6, Stop=7), `Total Hours`=row 8. Day columns at full-name
   day-label row (typically row 3, cols N,S,X,AC,AH,AM,AR stride 5).
2. **Pay-coding line-item table**: header at **row 11** (discovered by label —
   contains `WC STATE` + ≥5 `RT#/OT#/...` tokens). Each row 12+ is a
   job/cost-code **line item**; rows where the label-discovered `JOB` column
   is empty are echo/import-shadow rows and are **skipped** to avoid double-
   counting (`find_fm_job_col`).

Routing rules (discriminator + invariants):
- A sheet is Field Mechanic iff its day-label row uses **full day names**
  (MONDAY/TUESDAY/...). Standard/shifted use short names (MON/TUE/...). This
  is the only structural discriminator — pay-type labels (`RT1/OT1`) appear
  on both families and are NOT unique.
- Identity: Coulson/Pottage use A1/B2 (standard layout); Hall/Smith Dustin
  use alt-layout B1/C2 (`employee_alt_layout`).
- All Field Mechanic sheets go through `extract_field_mechanic_sheet` —
  every field still resolved by label (same `PAYTYPE_DAY_RE` regex used for
  standard sheets), so no per-family branching of *field resolution*; the
  only family-specific code is the line-item-table loop.

### 6.4 Label-driven pay-type discovery (the central architectural rule)

Every output pay-type column (RT/OT/DT/PD/4D/4A/PTO/HP) is resolved by reading
each column's actual label in the header row of the sheet, *per sheet*.
There is no positional assumption.

`anchors.discover_paytype_columns_by_label(ws, header_row, day_start_cols)`:
1. For each day-block (defined by the day columns discovered at runtime),
   scan that block's columns in `header_row` (the row containing `WC STATE`).
2. Match each cell against `PAYTYPE_DAY_RE = ^(RT|OT|DT|PD|4D|4A|PTO|HP)(\d)$`.
3. Return `{(day_idx, pay_type) -> column_idx}`. Unknown tokens or
   day-index/position mismatches surface as `[CHECK]`.

`extract_sheet` then reads each `(day, pay_type)` value from its discovered
column, routing it to the correct output field regardless of where the
column physically sits. The YAML `pay_types` list is a fallback used only
when no labels are found.

**Why this matters (the silent-mislabel bugs this closes):**
- v2 baseline read `'N'` from the FSL column as WC State because anchors
  were positional → 13 silently-wrong sheets.
- Wirth had `M12='PTO1'` instead of `M12='OT1'`; the old code read 19.98h
  of PTO into the OT output column.
- Impastato had `P12='HP1', Q12='PTO1'` instead of `4D1/4A1`; holiday/PTO
  hours were lost to the 4D/4A columns.
- Multiple PG&E sheets encode Friday holiday hours by putting the literal
  `'HP'` in row 5's Friday day-column; label discovery captures these into
  the HP output column.

All four bug classes are now structurally impossible — a mislabel either
surfaces as `[CHECK]` (unknown token) or routes to the correct output field.

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

## 8. Handling a new dataset (no code or YAML changes expected)

Because every field is label-driven (§6.4), the extractor should already
handle new workbooks that use the same labels, even if the column positions
differ. The expected workflow:

```powershell
# 1. Drop the new files into data/ (or data/sample*/ — anything recursive)
#    .venv\Scripts\python.exe is the canonical interpreter call on Windows.

# 2. Run the extractor against the whole data/ folder
.venv\Scripts\python.exe src\extract.py data\ --out output\corpus_v6.xlsx

# 3. Check the Run_Info + QA_Summary tabs for any FAIL / REVIEW counts
#    Look at Unmatched_Sheets for anything that classified-out

# 4. Run the independent audit (re-reads source, doesn't trust QA_Flag)
.venv\Scripts\python.exe scripts\audit_v4.py
#    (after pointing audit_v4.OUT at the new output, or symlink)

# 5. When a new feedback file lands, run the cell-by-cell comparator
.venv\Scripts\python.exe scripts\compare_v4_feedback.py
#    (after pointing CURRENT + FEEDBACK at the new files in the script)
```

**When you'd actually need a code change:**
- A new pay-type token appears (not in RT/OT/DT/PD/4D/4A/PTO/HP). Add it to
  `anchors.PAYTYPE_DAY_RE` AND to the output column list. Likely also add to
  `_NUMERIC_OUTPUT_COLS` in `extract.py` for `'0.00'` formatting.
- A new day-name spelling appears (e.g. localized). Add to
  `anchors.DAY_TOKENS` / `DAY_TOKENS_FULL` / `_DAY_ALIASES` — single point of
  truth, no code-path branching.
- A genuinely new layout family. Should still work via label discovery if
  the field labels match; if not, inspect with `scripts/inspect_fm_full.py`
  (general-purpose raw-grid dumper despite the name) first, then add a new
  branch in `process_workbook` only as a last resort.

---

## 9. Outstanding work / known limitations

### Resolved (commits 64367e3, a90cd9c, 00af7f1, a6bae9a)
- ✅ Hall alt-layout (B1/C2)
- ✅ Placeholder sheets (`A1='Input Employee Number'`)
- ✅ WC State / ST positional reads (now header-label resolved)
- ✅ Unmatched_Sheets visibility
- ✅ Empty-employee drop filter (zero hours + no time entries)
- ✅ Field Mechanic family (line-item table, RT/OT/DT/PTO/HP)
- ✅ Field Mechanic line-item doubling (JOB-column filter)
- ✅ Numeric output rendered as `'0.00'` strings matching feedback format
- ✅ **All pay-type columns label-driven per sheet** (closes silent
  position-drift class for RT/OT/DT/PD/4D/4A/PTO/HP)
- ✅ **sample2 feedback (commit `9b5bc8c`, 2026-06-02)** — see §14:
  - Per-day WC State / ST from the line-item with worked hours (Hilyard /
    Driessen multi-state weeks)
  - PP token + Rancho Runners family (5-col stride, bare-suffix day-1)
  - Pay-totals row prefers the summing row in multi-line-item sheets
    (Federline DT/PD per day now populates)
  - `#VALUE!` Total Hours falls back to the row below (Tweed FRI = 8.00)
  - Date format `MM/DD/YYYY` (+ `qa.check_date_sequence` updated to parse
    both 2- and 4-digit years)
  - Zero pay-type cells render blank (manager rule)
  - Time cells written as `datetime.time` objects, not strings
  - Typo / unknown-label `[CHECK]` flag (Morgan `DT22` no longer silent)

### Still open
1. **No parallelization.** Single-process; 5,000 files take ~4h sequential.
   Adding `multiprocessing.Pool` over workbooks is a 1-day task for ~10× speedup.

2. **Single large output workbook.** At 5,000 files, the output `.xlsx` is
   too large for Excel. Switch to CSV/Parquet output for large runs.

3. **No incremental processing.** Reruns re-extract everything. Adding
   filename+mtime hashing to skip unchanged files is straightforward.

4. **Operator's guide doesn't exist yet.** Non-engineers running at scale
   need a one-page README of "run this command, look at these tabs, do X
   when Y is flagged."

5. **Hall ST value depends on row-11 header layout.** Hall has WC STATE at
   L11 but no separate ST header. ST currently resolves to whatever the
   resolver finds when scanning H12 ± 1 — happens to be 'CA' in the corpus.
   If a Hall file ever has a different ST value, we'd miss it. Roster
   lookup via the workbook's `Employees` sheet would be the robust fix.

6. **`audit_v4.py` positional cross-check produces false positives** on
   sheets where the manager hand-entered holiday hours into a positionally-
   wrong grand-total cell (7 sheets). The label-driven output is right;
   the audit's positional sum-check disagrees with the source's own
   inconsistent positional sum. Either: (a) make the audit label-driven
   too, or (b) treat these as informational, not regressions. The
   feedback-file comparison (`compare_v4_feedback.py`) is the true accuracy
   metric — 15/16 fields at 100%.

7. **YAML drift toward irrelevance.** `pay_types`, `day_start_cols`,
   `grand_totals.cols` are now fallback hints only. Could be removed once
   the label-driven path has been exercised against a wider corpus.

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
- **Label-driven extraction is THE rule for all pay-types** (RT/OT/DT/PD/4D/4A/PTO/HP/PP). Don't add positional reads to `extract_sheet`. New pay-type tokens get added to `anchors.PAYTYPE_DAY_RE`, not to extract code.
- **WC State / ST is per-day, not per-employee.** When a sheet has multiple line items with different state codes, each day row carries the state of the line item that owns that day's hours. Use `resolve_state_per_day_from_line_items`, not `resolve_anchor_cell`, for these fields.
- **Bare-suffix pay-type tokens (`PP`, `PTO`, `HP`) are only day-1 inside the day-1 block.** Outside that block they're weekly-totals headers and must be skipped. The bug class that hits when this is ignored: bare `PTO`/`HP` at cols 50+ silently overrides legitimate `PTO1`/`HP1` mappings.
- **Compare against the manager's feedback file as the source of truth.** `Overall=PASS` in QA_Summary is a derived self-check, not validation. The cell-by-cell comparator (`compare_v4_feedback.py`) is the real accuracy metric.

---

## 12. Session continuity for Claude

**Working directory:** `C:\dev\Extract-excel` (NOT the OneDrive folder —
OneDrive corrupts git pack files).

When resuming, read in this order:
1. This file §5 (current state), §6.4 (the label-driven rule), §8 (new-dataset
   workflow), §13 (verification recipe), §14 (sample2 fix log), §15
   (discussion notes), §16 (large-dataset handling).
2. `SESSION_HANDOFF.md` — most recent session's notes, including any
   in-flight context.
3. `src/anchors.py` — discovery primitives. The functions that matter:
   `discover_paytype_columns_by_label`, `discover_field_mechanic_paytype_cols`,
   `resolve_state_per_day_from_line_items`, `resolve_pay_totals_row`,
   `find_field_mechanic_header_row`, `find_fm_job_col`, `resolve_anchor_cell`,
   `classify_sheet`.
4. `src/extract.py` — `extract_sheet`, `extract_field_mechanic_sheet`,
   `_format_numeric_cells`, `_drop_empty_employees`, `process_workbook`,
   `_write_excel` (the time-cell post-write).
5. `output/corpus_v5.xlsx` and `output/corpus_sample2_v3.xlsx` — the two
   manager-validated outputs.
6. `corpus_sample2 - Feedback.xlsx` (repo root) — annotated truth for sample2.

**Git state at last update (2026-06-02):**
- Branch: `main`
- HEAD: `9b5bc8c` "Resolve sample2 feedback: per-day WC/ST, PP family, blank zeros, MM/DD/YYYY dates"
- Remote: https://github.com/Zubin123/Extract-excel.git
- Recent commits worth reading in order:
  - `64367e3` header-verified anchors + Hall + placeholders (Phase 3)
  - `a90cd9c` empty-employee drop filter + empty-value flagging
  - `00af7f1` Field Mechanic extraction
  - `a6bae9a` all pay-types label-driven (sample/sample1 architecture commit)
  - `8b67410`, `4c5245e` rename to `corpus_v5.xlsx`
  - `211e330` 2026-05-29 docs refresh
  - `9b5bc8c` **sample2 feedback resolved** (the current head — see §14)

---

## 13. Verifying a new dataset (the accuracy-check recipe)

When new workbooks arrive (manager drops them into `data/sample*/` or any
subfolder of `data/`), this is the exact sequence to confirm the extractor
is still hitting accuracy. The same checks that proved 15/16 fields at 100%
on the v4 feedback apply unchanged to any new dataset.

### Step 1 — Extract
```powershell
.venv\Scripts\python.exe src\extract.py data\ --out output\corpus_v6.xlsx
```
- Watch the per-file `emp= pass= unmatched=` log for sheets that didn't
  extract. Look at the `Dropped` count (empty-employee filter).
- Open `Run_Info` tab: `Sheets Extracted`, `Sheets Passed`,
  `Sheets Needing Review`, `Sheets Unmatched`, `Overall Result`.

### Step 2 — Independent audit (re-reads source, doesn't trust QA)
Edit the line `OUT = ROOT / "output" / "phase2_corpus_v4.xlsx"` in
`scripts/audit_v4.py` to point at the new output, then:
```powershell
.venv\Scripts\python.exe scripts\audit_v4.py
```
**What 100% looks like:**
- `Pay-type grand totals match source: N/N`
- `WC State matches header-resolved source value: N/N`
- `ST matches header-resolved source value: N/N`
- `Suspicious value counts: (none)`

Source `#VALUE!` Total Hours are correctly categorised as
`source_th_value_error` (informational — not a regression). The 7
positional-grand-total mismatches we already see (Wirth/Abella/etc.) will
persist as informational unless the audit is upgraded to label-driven
totals.

### Step 3 — Cell-by-cell vs the manager's expected output (the truth)
When a new feedback file arrives (manager's annotated version of the
output), edit `scripts/compare_v4_feedback.py`:
- Set `CURRENT` to the new output path
- Set `FEEDBACK` to the new feedback file path
- Verify `EXPECTED_COLS` mapping still matches the feedback file's column
  layout (run a quick header inspection first — see the comment block in
  the script for the layout we used)

Then:
```powershell
.venv\Scripts\python.exe scripts\compare_v4_feedback.py
```

**What 100% looks like:**
- Field-by-field accuracy lines all `100.0%` or `99+%` with the only
  remaining mismatches being rounding differences (e.g. source `19.98` vs
  manager-rounded `20.00`).
- `Expected but missing now`: count should be near zero. Any non-zero
  count = real coverage gap; investigate.
- `Present now but not expected`: expected to be non-zero ONLY for newly
  added templates (e.g. Field Mechanic was 88 of these in the v4 feedback,
  because the manager intentionally excluded them at the time).

### Step 4 — If anything is not 100%
The mismatches sort cleanly into a small number of root causes:
- **Source value differs from manager expected (rounding / hand-correction)**
  — accept; we report source faithfully.
- **Manager specified a label our regex doesn't know yet** — add to
  `anchors.PAYTYPE_DAY_RE` or the day-token vocabulary.
- **Sheet's identity layout is new** (A1/B2/B1/C2 doesn't apply) — extend
  `classify_sheet` in `anchors.py`.
- **A new structural family appears** (not standard / shifted / Field
  Mechanic) — inspect with `scripts/inspect_fm_full.py`, then decide
  whether label discovery alone suffices or whether a new extract path
  is needed.

Never accept a "PASS that doesn't match the feedback" — the feedback file
is the ground truth, the in-output QA_Summary is a derived signal.


---

## 14. sample2 fixes (commit `9b5bc8c`, 2026-06-02)

Manager raised six items on the first sample2 extract; this section captures
what each one was and the structural fix shipped. Read this if a future
feedback round reopens any of these classes.

### #1 — WC State / ST per-day from line-item-with-hours
**Symptom:** Hilyard / Delair / Knowland (PacifiCorp 01/22/23) and Driessen
(PacifiCorp 01/01/23) had multi-state weeks (CA most days, OR one day).
Old code took the *first non-empty* value under the WC STATE header — a
single code per employee — losing the OR/CA split.

**Fix:** `anchors.resolve_state_per_day_from_line_items` walks every line
item below the header, sums each day-block on that row, and for any day
with hours > 0 records that line item's WC/ST value. `extract_sheet` and
`extract_field_mechanic_sheet` look up `per_day[day_idx]` when emitting
each day row, falling back to the first non-empty value for days with no
hours (weekends). Closes the Hilyard Thursday `OR` and Driessen TUE-FRI
`OR` cases. **Never assumes a single state per employee again.**

### #2/#3/#4 — PP token + Rancho Runners family without breaking standard sheets
**Symptom:** Adding `PP` to the regex (for Rancho Runners, where the day-1
column is bare `PP` not `PP1`) broke Adam Peart and Rancho Office: bare
`PTO` / `HP` labels in the grand-totals summary area (cols 50+) were being
interpreted as day-1 labels and *overrode* the legitimate `PTO1`/`HP1`
mappings, swapping PTO and HP for day 1.

**Fix:** `anchors.PAYTYPE_DAY_RE_OPTIONAL_SUFFIX` accepts bare tokens, BUT
`discover_field_mechanic_paytype_cols` only treats a bare token as day 1
when its column sits *inside the day-1 block*. The day-1 block is derived
from the day-label row (find MON & TUE columns; gap = block width).
Bare tokens outside the day-1 block are skipped — they are weekly summary
headers, not per-day labels.

### #5 — Pay-totals row selection for multi-line-item sheets
**Symptom:** Federline (storm week, 3 line-item rows + 1 sum row) and Tweed
were reading from one line-item row instead of the summing row. Days the
line item didn't cover (MON, SAT, SUN for Federline) showed as 0.

**Fix:** `anchors.resolve_pay_totals_row` ranks candidate rows by populated
day-cell coverage and tolerates 1 cross-check mismatch on the winning row.
The summing row always has the highest coverage; it wins even when its
grand-total cell disagrees with its day-sum due to source-data
inconsistency (Federline RT 32 vs grand-total 35 = manager hand-correction).

### #6 — Total Hours `#VALUE!` fallback
**Symptom:** Tweed FRI 01/27/2023 showed TH=0.00. Source row 9 (Total Hours)
FRI cell = `#VALUE!` formula error; row 10 (Time Sheet Hours label) FRI = 8.

**Fix:** In `extract_sheet`, when the Total Hours cell is a `#`-prefixed
error string, fall back to the cell one row below (the recovered value).

### #7 — Date format MM/DD/YYYY
**Symptom:** Output had `12/26/22`; manager wants `12/26/2022`.
**Fix:** `fmt_date` uses the 4-digit-year strftime directive instead of the
2-digit one.
**Companion fix in qa.py:** `check_date_sequence` was parsing dates with
`%m/%d/%y` only — after the format change every day row got a `[CHECK] Date
'12/26/2022' not parseable` flag and the entire dataset went REVIEW. Fixed
by trying `%m/%d/%Y` first, then falling back to `%m/%d/%y` for backwards
compatibility. Any new date-format change needs the same dual-format parse
treatment.

### #8 — Blank zero pay-type cells
**Symptom:** Output rendered zero pay-type values as `'0.00'`; manager wants
blank cells (`None`) when there's no real value.
**Fix:** `_format_numeric_cells.fmt_nonzero_blank` returns `None` for zero/
blank inputs on day rows. Total Hours still renders `'0.00'` so the zero-hour
row remains visible. TOTALS rows keep their numeric grand totals.

### #9 — Time cells as `datetime.time` objects, not strings
**Symptom:** Pandas serializes `datetime.time` to `'07:00:00'` strings on
write; manager wants raw time-of-day values so Excel renders them as times.
**Fix:** Snapshot the four clock columns before pandas writes, then overwrite
those cells with `datetime.time` values via openpyxl post-write and set
`number_format = "h:mm AM/PM"`.

### #10 — Typo / unknown-label `[CHECK]` flag
**Symptom:** Morgan sheet had `'DT22'` and `'DT32'` typos in the PD2/PD3
slots. The strict regex skipped them silently; the underlying cell value
(if any) was lost without a flag.
**Fix:** `discover_paytype_columns_by_label` emits `[CHECK] header U12='DT22'
looks like a pay-type label but doesn't match the day-1..7 form ...` whenever
a day-block header cell starts with a known pay-type prefix but fails the
strict regex. Surfaces silent data loss instead of hiding it.

---

## 15. Discussion notes from the sample2 round

These are the design decisions the manager and Claude reached that future
sessions should preserve:

- **`4S` exists but isn't per-day.** Found at Spaulding template col 57, in
  the grand-totals summary row alongside `1, 3, 2, HP, PTO, Total`. Not added
  to the output schema (manager said skip for now). If a future dataset shows
  per-day `4S1..4S7` labels, add `4S` to `PAYTYPE_DAY_RE` and the output
  columns.
- **`12, 4A22, 4A24, 4A25, 4D52, DT22, DT32, 3.5` are typos.** Manager
  confirmed during sample2 review. Don't add to the regex; the typo `[CHECK]`
  flag (#10 above) surfaces them.
- **Manager comments on sample2 feedback file:** 25 annotations across 4
  categories — "Full Excel Extraction missed" (17 = Manuals 4 employees),
  "DT entry was missed" (3 = Federline), "PD entry was missed" (4 = Federline),
  "Total Hours was missing" (1 = Tweed FRI 01/27/23). All resolved.
- **Per-day WC State is NOT "all distinct codes joined with `;`".** Initial
  interpretation was wrong. Manager wants the per-row code of the line item
  that owns that day's hours. Row-level, not employee-level.
- **The `OFF` sentinel in the Start cell is NOT a hours-zeroing signal.** When
  Start = `'OFF'`, the day may still have legitimate hours filed under a pay-
  type cell or recovered from `Time Sheet Hours`. The flag we emit is
  informational only.
- **Time cells without lunch — open interpretation question.** ~13 employees
  have rows with Start + one mid-time value (Lunch Out filled, Lunch In + Stop
  blank). Manager's feedback file places that mid-time in the Stop column;
  we keep it in Lunch Out (source-faithful). Not yet flagged by the manager
  as a bug; left as source-faithful unless explicitly confirmed.

---

## 16. Large-dataset handling (years 2022, 2024, 2025, 2026)

The full corpus is **~1-1.4 GB per year folder, ~4-5.6 GB total** uncompressed
`.xlsx` files (the 500-800 MB figure earlier was the compressed/zipped form).
Estimated ~250-500 workbooks per year, ~1,000-2,000 workbooks total.

### Storage
- **Don't put it in the repo working tree.** Git `status` walks the working
  tree even when files are gitignored — 5 GB of files = thousands of inode
  stats per command. Use an external path like `D:\timesheet-data\2024\` and
  pass it via CLI:
  ```powershell
  .venv\Scripts\python.exe src\extract.py "D:\timesheet-data\2024" --out output\corpus_2024.xlsx
  ```
- **Avoid OneDrive / SharePoint / cloud-synced paths.** Same corruption risk
  that motivated moving the working tree off OneDrive (§12). External SSD
  over USB is fine; mounted SharePoint / OneDrive is not.

### Output sizing
- **One output workbook per year.** Extrapolating from sample2 (63 wb →
  ~3,200 data rows): roughly ~50,000 rows per year, ~5-15 MB output file.
  Comfortable for Excel.
- Excel row limit = 1,048,576. ~250-400 employees × 8 rows × ~50 weeks per
  year ≈ 80-160k rows per year — well within limits.
- Combined output across 4 years would be ~200k rows total. Still openable,
  but per-year files give faster iteration if the manager flags issues.

### Performance
- Serial extraction is **~1.5-2.5 sec/workbook** (sample2: 63 wb in ~2-3 min).
- Full corpus serial estimate: **4-7 hours** for ~1,500 workbooks.
- The `multiprocessing.Pool` task in §9 item 1 becomes high-leverage at this
  scale — with 4-8 cores would drop to ~30-60 min total. Not yet implemented;
  worth doing before the next full-corpus rerun.
- **Don't run multiple extractions in parallel on the same machine** unless
  you have 16+ GB RAM free. openpyxl loads each workbook fully into memory
  (data-only mode reads the whole sheet); some workbooks are 50+ MB each.

### Verification
- Run the audit + cell-by-cell feedback comparator **per year**. Per-year
  feedback files keep the comparison tractable and let the manager review
  in chunks.
- If a single year extraction crashes mid-run, the extractor processes
  workbooks serially and writes the output only at the end — there's no
  partial recovery. Investigate the failing workbook and rerun. (Adding
  incremental processing is in §9 item 3.)
