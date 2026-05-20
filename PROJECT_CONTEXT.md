# Project Context — Excel Timesheet Extractor

> This document captures everything a new Claude/developer session needs to resume the project on any machine. **Claude Code sessions are local — they do not transfer via git.** This file is the handoff.

---

## 1. What the project is

A **deterministic Python script** that reads weekly-timesheet Excel workbooks (one sheet per employee, fixed grid layout) and outputs a flat tabular Excel with three sheets:

1. **`Data`** — 80 rows (10 employees × 8 rows: MON-SUN + TOTALS) with a `QA_Flag` column
2. **`QA_Summary`** — per-employee accuracy comparison with visible numbers
3. **`Run_Info`** — run metadata (timestamp, counts, QA method, version)

No AI / LLM at runtime. All cell coordinates are fixed and known (see [CLAUDE.md](CLAUDE.md) Section 4).

---

## 2. Current state (as of last commit)

- **All 19 tests pass** (`uv run pytest`)
- **10/10 employees pass accuracy check** on the sample input file
- Output Excel at [output/extracted.xlsx](output/extracted.xlsx) reflects the latest run
- Input sample at [data/WE 01 08 22 CA GF'S - WA 4939-0899-6263_1.xlsx](data/)
- Template at [data/Header Template New_Rev 1.xlsx](data/)

---

## 3. Key design decisions (history of "why")

### 3.1 Three-sheet output instead of single sheet
**Why:** Stakeholders (non-technical) need to *see* the accuracy claim, not just trust a console message. The `QA_Summary` sheet shows side-by-side `daily_sum` vs `grand_total` vs `match` for every employee and every pay type. Anyone can audit it without reading code.

### 3.2 Accuracy check methodology
The script cross-checks **its own extracted values** against **Excel's own pre-computed formula totals**:

| Column | Daily source | Grand-total source |
|---|---|---|
| RT, OT, DT, PD, 4D, 4A | Row 24, cols L–BA (per day block) | Row 24, cols BB–BG |
| Total Hours | Row 9, day-block start cols | Cell `BB9` |

Tolerance: `0.01`. Status per column: `MATCH` / `MISMATCH` / `SKIPPED (#VALUE in source)`.

**SKIPPED ≠ FAIL.** When the source Excel itself has a `#VALUE!` formula error (e.g. Case MON, Schmidt MON in the sample), the comparison is skipped — the source itself is broken, not our extraction. We surface it via the `QA_Flag` column on the affected rows.

### 3.3 `coerce_total_hours` vs `coerce_hours`
- `coerce_hours` returns `0.0` for blank/error cells (safe for sum aggregation of pay types).
- `coerce_total_hours` returns `None` for blank/error cells — preserves the distinction so the QA check can detect them.

### 3.4 `QA_Flag` column on every data row
Surfaces real source-data problems:
- `OK` — clean row
- `TOTALS` — clean TOTALS row
- `no time data` — day has no Start/Lunch/Stop values (employee off)
- `total_hrs='#VALUE!'` — source has a formula error in the day's Total Hours cell
- `TOTALS; grand_total_hours='#VALUE!'` — BB9 has a formula error

Colour-coded in Excel: green / grey / red.

### 3.5 `tmp_path` for tests
Tests write to a pytest temp directory, never to `output/`. This avoids `PermissionError` when the user has the output Excel open in Excel while running tests.

---

## 4. Where extraction CAN silently break (audit list)

Documented during stakeholder discussion — also relevant for future code review:

1. **`data_only=True` trap** — openpyxl reads cached formula values. If source file was generated programmatically without an Excel save, all formula cells return `None` → QA check passes with all zeros. **Most dangerous.**
2. **Template drift** — if a column is inserted in the source template, `DAY_BLOCKS` shifts and every cell read is wrong, no error thrown.
3. **Circular validation** — daily subtotals and grand totals are both formulas in the same source file. A formula bug in the source would consistently lie to both sides of our check.
4. **`coerce_hours` swallows strings** — non-numeric pay-type cells silently become 0.0. Detected by QA only if grand total disagrees.
5. **QA does NOT verify** — time fields (Start/Lunch/Stop), employee name, EE ID, WC State, ST. These are pulled but not cross-checked.
6. **New pay types beyond 4A** — silently ignored.

These are documented limitations, not bugs. See conversation history in section 8.

---

## 5. Project layout

```
Excel_extract/
├── CLAUDE.md                 ← original spec (verified, do not re-derive)
├── PROJECT_CONTEXT.md        ← this file
├── PLAN.md                   ← what's next (see Section 7)
├── pyproject.toml            ← uv-managed
├── uv.lock
├── .gitignore                ← excludes .venv/, .pytest_cache/, ~$*.xlsx
├── data/
│   ├── WE 01 08 22 CA GF'S - WA 4939-0899-6263_1.xlsx   ← sample input
│   └── Header Template New_Rev 1.xlsx                    ← schema source
├── src/
│   └── extract.py            ← main script + library
├── output/
│   └── extracted.xlsx        ← latest 3-sheet output (committed for review)
├── scripts/
│   └── inspect_totals.py     ← one-off probe used during BB9 investigation
└── tests/
    └── test_extract.py       ← 19 tests
```

> **Filename note:** the actual files in `data/` use *spaces and an apostrophe* (`WE 01 08 22 CA GF'S - ...`), not the underscored names mentioned in CLAUDE.md. The script handles both via `infer_box_name`.

---

## 6. How to run (any machine, after `git clone`)

```bash
# 1. Install uv if not present: https://docs.astral.sh/uv/

# 2. Restore environment (reads pyproject.toml + uv.lock)
uv sync

# 3. Run extraction
uv run python src/extract.py "data/WE 01 08 22 CA GF'S - WA 4939-0899-6263_1.xlsx"

# 4. Run tests
uv run pytest tests/ -v
```

Expected console output: `10/10 employees passed accuracy check.` plus per-employee status table.

CLI flags: `--folder "2022"` (default), `--box "WE 01 08 22"` (inferred from filename), `--out output/extracted.xlsx` (default).

Exit codes: `0` on all-pass, `2` if any employee fails.

---

## 7. Future plans

See [PLAN.md](PLAN.md) for the detailed scaling roadmap. High-level:

- **Phase 1 (done):** Single-file extraction + QA + tests.
- **Phase 2:** Batch mode — process a folder of N files into one combined output (multiprocessing).
- **Phase 3:** Schedule / pipeline — for ongoing weekly files. Options range from Windows Task Scheduler to Azure Functions + SQL Server + Power BI.

The user's stakeholder context: ~5,000 historical files need extraction. See PLAN.md for timing estimates and infrastructure options.

---

## 8. Conversation history summary (for next session)

Key user feedback shaped the design:

1. **"Add QA transparency to the output Excel, not just console"** → led to 3-sheet design with `QA_Summary` sheet showing visible daily-sum vs grand-total comparison.
2. **"Flag rows with extraction problems"** → led to `QA_Flag` column on every data row.
3. **"Do we even need a QA test for deterministic code? Where can it break?"** → led to the documented audit list (Section 4). The honest answer: determinism guarantees repeatability, not correctness; the QA check catches wrong-cell-address bugs and source-data issues that would otherwise be silent.
4. **"You ignored Total Hours grand total even when it's present"** → led to reading `BB9` cell and adding Total Hours to the accuracy check with `SKIPPED` tri-state for when source has `#VALUE!`.

---

## 9. For the next Claude session (resuming on another machine)

Read in this order:
1. **[CLAUDE.md](CLAUDE.md)** — original spec (cell coordinates are verified; don't re-derive)
2. **This file** — context, design decisions, current state
3. **[PLAN.md](PLAN.md)** — what's pending
4. **[src/extract.py](src/extract.py)** — the implementation
5. **[tests/test_extract.py](tests/test_extract.py)** — what's covered

Then run `uv sync && uv run pytest` to confirm a clean baseline before changes.
