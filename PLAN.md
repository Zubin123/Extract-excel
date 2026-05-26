# Project Plan — Excel Timesheet Extractor

## Status

| Phase | State | Notes |
|---|---|---|
| 1. Single-file extraction + QA + tests | **DONE** | 19/19 tests passing, 10/10 employees pass accuracy |
| 2. Batch mode + layered QA | **DONE** | Folder-walking CLI, 5-tab output, layered QA flags |
| 3. Layout-drift fixes (header-verified anchors) | **DONE** | Commit `64367e3`. 208/208 feedback issues resolved. See [CLAUDE.md](CLAUDE.md) §5–§6 for details. |
| 4. Parallelization | TODO | `multiprocessing.Pool` for ~10× speedup on 5,000-file batches |
| 5. Pipeline / scheduling | TODO (decision pending) | See Section 3 for options |

---

## 1. Phase 2 — Batch mode (the next concrete task)

### Goal
Process N input files in a folder into a single combined output Excel with all employees from all files in `Data`, and one row per (file × employee) in `QA_Summary`.

### Proposed CLI
```bash
uv run python src/extract.py --batch data/2022/ --out output/2022_combined.xlsx
```

### Implementation sketch
1. Refactor `extract()` to return DataFrames instead of writing directly.
2. Add `extract_batch(input_dir, out_path)`:
   - Glob `**/*.xlsx` (excluding `~$*` lock files)
   - Use `multiprocessing.Pool` to process files in parallel
   - Concatenate all DataFrames
   - Write one workbook with 3 sheets (same shape as single-file output)
3. Add a per-file row in `Run_Info` (input filename, pass/fail count) instead of one row.

### Estimated performance

| File count | Single-threaded | 8 cores (multiprocessing) |
|---|---|---|
| 100 | ~17 min | ~2 min |
| 1,000 | ~2.8 hr | ~21 min |
| 5,000 | ~14 hr | ~1.75 hr |
| 10,000 | ~28 hr | ~3.5 hr |

(Assumes ~10s per file from current single-file benchmark; openpyxl is the bottleneck.)

### Risks for batch mode
- **Filename variation** — `infer_box_name` regex assumes `WE MM DD YY` pattern. Files outside that convention need a fallback.
- **Schema drift across years** — if 2022 vs 2025 files have different column positions, the script silently extracts wrong cells. Mitigation: run a structural pre-check (verify A1 is non-empty, B2 is int, row 4 has `MON`/`TUE`/...) per file and reject on mismatch.
- **Memory** — 10,000 files × 80 rows = 800k rows. Still <200MB in pandas. Fine.

---

## 2. Outstanding gaps from Phase 1 (small fixes)

- [ ] **Schema pre-check** before extraction (verify the template hasn't drifted). Function: `validate_sheet_structure(ws) -> list[str]` returning a list of structural problems. Reject the sheet if any are found rather than silently extracting wrong cells.
- [ ] **Sheet name normalisation** — currently relies on exact match (`Raper ` with trailing space). A future file with `Raper  ` (two trailing spaces) would slip past `REFERENCE_SHEETS` and be treated as data. Mitigation: include known-employee-name check using the `Employees` reference sheet.
- [ ] **Logging** — replace `print()` with `logging` module so batch runs can be silenced or piped to file.
- [ ] **JSON report** alongside Excel — easier for downstream tooling to consume.

---

## 3. Phase 3 — Pipeline / scheduling options

Decision pending on volume and frequency. Three tiers:

### Tier A: One-time migration of historical files
- Use Phase 2 batch mode locally / on any server
- One run, walk away, ~2 hours for 5,000 files
- **Cost: $0. Infra: laptop/server.**

### Tier B: Recurring weekly ingestion
- ~50 new files per week
- Windows Task Scheduler (or `cron`) running the batch script
- Writes timestamped output Excel to a shared folder
- **Cost: $0. Infra: any always-on Windows/Linux machine.**

### Tier C: Production pipeline (cloud)
For continuous ingestion, audit history, and dashboards:

```
File drop (SharePoint / S3 / Azure Blob)
    ↓ (event trigger)
Azure Function / AWS Lambda
    ↓ (runs extract.py)
SQL Server / Postgres (Data + QA_Summary + Run_Info tables)
    ↓
Power BI / Tableau dashboard
```

- Each file processed in <30s on serverless
- Database: query "show me all 2023 timesheets where Total Hours had a #VALUE!"
- Dashboard: per-week pass rate, per-employee flags
- **Cost: ~$5–300/month depending on volume.**

### Recommendation for now
Build Phase 2 (batch mode) first. Pick a Tier only when the actual file volume and cadence is known.

---

## 4. Long-term hardening ideas (post-MVP)

- **Diff-based QA on consecutive weeks** — flag suspicious week-over-week deltas per employee (e.g. RT jumped from 40 → 80).
- **Source-file fingerprinting** — hash the template region (header rows, formula cells) to detect template drift before extraction.
- **Auto-recovery for `#VALUE!` cells** — for Total Hours, recompute from Start/Stop/Lunch times instead of trusting the broken Excel formula.
- **PTO / HP support** — currently always blank because the sample didn't have them. Confirm cell location when a file with PTO/HP arrives.

---

## 5. Open questions for the stakeholder

1. What's the actual file volume? (One-time 5k? Weekly 50? Monthly drip?)
2. Where do the files live today? (SharePoint, file server, email attachments…)
3. Where should the output go? (One mega-Excel? Database? Dashboard?)
4. Is the source template version-controlled or could the cell layout change year-over-year?
5. Who needs to read the QA_Summary — payroll, finance, audit?

The answers to (1) and (3) determine which Tier (A/B/C) we build in Phase 3.
