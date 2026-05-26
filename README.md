# Excel Timesheet Extractor

Deterministic Python pipeline that converts weekly-timesheet Excel workbooks into a flat, audited tabular output with layered QA and header-verified field anchors.

## Quick start

```powershell
# Install uv first: https://docs.astral.sh/uv/

uv sync

# Run on a folder of workbooks
.venv\Scripts\python.exe src\extract.py data\ --out output\run.xlsx

# Or a single file
.venv\Scripts\python.exe src\extract.py "data\WE 01 08 22 CA GF'S - WA 4939-0899-6263_1.xlsx"
```

## Output (5 tabs)

| Tab | What's in it |
|---|---|
| `Data` | Row-per-day per employee + TOTALS row, with QA_Flag column |
| `QA_Summary` | Per-employee pay-type accuracy comparison + `Overall` status (PASS / REVIEW / FAIL / NO_GRID) + Sheet Issues column |
| `Run_Info` | Run metadata, totals, overall result |
| `Unmatched_Sheets` | Sheets the pipeline refused to extract, with reason |
| `ID_Conflicts` | Same EE ID → multiple names within one workbook |

## Documentation

- [CLAUDE.md](CLAUDE.md) — **source of truth.** Layout reference, classifier behaviour, header-verified anchors, outstanding work, session-continuity notes
- [PROJECT_CONTEXT.md](PROJECT_CONTEXT.md) — design decisions, "where it can break"
- [PLAN.md](PLAN.md) — phase status and roadmap

## How extraction stays correct under template drift

The pipeline uses **header-verified anchors** instead of fixed cell positions for free-floating fields:

- `resolve_anchor_cell` finds each field by reading its header label, then walks downward for the first non-empty data cell
- Handles column drift (FSL inserted before WC STATE), header-row drift (row 11 vs row 12), and data-row drift (data at row 15 vs row 13)
- Any layout fallback emits a `[CHECK]` flag — silent misreads are structurally impossible for header-verified fields

Pay-type hours (RT/OT/DT/PD/4D/4A) are still verified by sum cross-check: `sum(MON..SUN per pay type)` must equal the grand-total cell. Tolerance: 0.01. See [CLAUDE.md](CLAUDE.md) §6 for full details.

## Stack

- Python 3.12
- [uv](https://docs.astral.sh/uv/) for dependency management
- `openpyxl`, `pandas` (runtime)
- `pytest` (dev)
