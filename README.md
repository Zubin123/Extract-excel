# Excel Timesheet Extractor

Deterministic Python script that converts weekly-timesheet Excel workbooks into a flat, audited tabular output.

## Quick start

```bash
# Install uv first: https://docs.astral.sh/uv/

uv sync
uv run python src/extract.py "data/WE 01 08 22 CA GF'S - WA 4939-0899-6263_1.xlsx"
uv run pytest tests/ -v
```

Expected: `10/10 employees passed accuracy check`, all 19 tests green.

## Output

Three-sheet Excel at `output/extracted.xlsx`:

1. **`Data`** — 80 rows (10 employees × 7 day rows + 1 TOTALS row), with a `QA_Flag` column flagging any source-data issues per row.
2. **`QA_Summary`** — per-employee comparison: `daily_sum` vs `grand_total` vs `match` for every pay type and Total Hours. Colour-coded (green = MATCH, yellow = SKIPPED, red = MISMATCH).
3. **`Run_Info`** — run metadata.

## Documentation

- [CLAUDE.md](CLAUDE.md) — original verified spec (cell coordinates, sheet layout)
- [PROJECT_CONTEXT.md](PROJECT_CONTEXT.md) — design decisions, "where it can break", state of the project
- [PLAN.md](PLAN.md) — what's next (batch mode for 5,000+ files, pipeline options)

## How the accuracy check works

The script cross-checks **its own extracted values** against **Excel's own pre-computed formula totals**:

| Column | Daily source | Grand-total source |
|---|---|---|
| RT, OT, DT, PD, 4D, 4A | Row 24, cols L–BA | Row 24, cols BB–BG |
| Total Hours | Row 9, day-block start cols | Cell `BB9` |

Tolerance: `0.01`. If the sum of 7 extracted day values matches Excel's pre-computed grand total, the extraction is verified for that column.

## Stack

- Python 3.12
- [uv](https://docs.astral.sh/uv/) for dependency management
- `openpyxl`, `pandas` (runtime)
- `pytest` (dev)
