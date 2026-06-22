# Excel Timesheet Extractor — A Plain-English White Paper

*Turning 5,000 timesheet workbooks into one verified payroll spreadsheet*

How we built a reliable, auditable pipeline that extracts weekly payroll data with zero silent failures — and why this approach beats AI for payroll-grade reliability.

---

## At a glance

| | |
|---|---|
| **5,000+** | files per run, no manual touch |
| **3,896** | verified data rows in the latest corpus |
| **208 / 208** | known extraction issues resolved |
| **0** | silent failures — every problem is flagged |

---

## The problem we set out to solve

Every week, dozens of timesheets arrive as Excel workbooks. Each contains multiple employee sheets with hours worked, pay categories (regular time, overtime, double time, per diem, etc.), dates, and worker classifications. Pulling all this into one place — accurately, on time, every week — meant hours of copy-paste, error-prone manual entry, and no easy way to audit later. Multiply by 5,000 historical files, and manual entry simply isn't viable.

## What we built

A self-running tool that reads any number of timesheet Excel files in a folder and produces a **single verified payroll spreadsheet** — with every number cross-checked against the source workbook's own formulas. No human in the loop for the routine 99%. Humans only review the small handful of rows the tool itself flags.

```
Folder of workbooks   →   Pipeline runs   →   Verified output
(any number of .xlsx)    (reads, verifies,    (5-tab spreadsheet)
                          flags)
```

## How it works — in plain English

Think of it as a very careful intern who doesn't get tired, doesn't make typos, never skips a sheet, and finishes thousands of timesheets in the time it takes to make coffee. The intern follows four rules:

1. **Read every sheet.** Walk through each workbook and identify which sheets are real employee timesheets vs reference data.
2. **Find each field by its label, not by a fixed position.** Templates drift — a column gets inserted, a row shifts. The tool reads the header row to find each value, so small layout changes don't break it.
3. **Cross-check the math.** Add up Monday-to-Sunday hours per pay category, compare to Excel's own weekly total. A mismatch flags the row — it doesn't silently accept it.
4. **Surface everything.** Output has five clearly labeled tabs: the data itself, a per-employee accuracy summary, the run log, sheets that couldn't be safely extracted, and EE-ID conflicts. Nothing is hidden.

## Accuracy — what we can guarantee

- ✓ **Deterministic.** Same input always produces the same output, every run.
- ✓ **Self-validating.** Every numeric column is cross-checked against Excel's own totals.
- ✓ **Auditable.** Each output cell traces back to a specific source cell.
- ✓ **No silent drops.** Sheets that can't be safely extracted are visibly listed with a reason.

## Scale — how fast it runs

| Volume | Single-threaded | Parallelized |
|---|--:|--:|
| 1 workbook | ~10 sec | ~10 sec |
| 100 workbooks | ~17 min | ~2 min |
| 1,000 workbooks | ~3 hours | ~21 min |
| 5,000 workbooks | ~14 hours | ~1.75 hours |
| 10,000+ workbooks | ~28 hours | distribute via cloud |

## Where this can be used

- **Weekly payroll preparation** — consolidate timesheets from dozens of crews into one upload-ready file.
- **Year-end audit & compliance reporting** — produce a clean, traceable record of hours worked across the year.
- **Historical data migration** — moving years of legacy timesheets into a new HR/payroll system.
- **Cross-company reconciliation** — comparing timesheets across job sites, companies, or pay periods.
- **Any structured Excel data extraction** — the same approach works for invoices, expense reports, inspection logs.

## Why not AI or OCR?

We deliberately chose **not** to use an AI model (LLM) or OCR engine. The reasoning:

| | Our approach | AI / LLM approach |
|---|--:|--:|
| Cost per 5,000-file run | $0 | ~$1,500–$15,000 |
| Same input → same output? | Always | No (varies between runs) |
| Traceable to source cell? | Every cell | No |
| Can invent values? | Never | Yes (hallucination risk) |
| Speed (5,000 files) | ~2 hours | ~40+ hours |

OCR adds no value when the input is already `.xlsx` — every cell already has structured text and formulas. AI is the right tool for some problems; payroll arithmetic is not one of them. We may use AI later as a *fallback* for the few sheets the deterministic pipeline can't handle, but never for the main data path.

## What "automation" means here

> **Not magic. Not AI. Just precise rules executed reliably.**
>
> Write down the rules once. Write them precisely. Run them across thousands of files. Surface anything the rules can't safely handle. Boring on purpose — and boring is what makes it reliable enough for payroll.

## The benefits in one paragraph

What used to take a person **days** of manual copying now runs in **hours**, with every number cross-checked, every problem row flagged, and every output cell traceable to its source. There is no ongoing infrastructure cost. The pipeline is reproducible — running it again tomorrow on the same files produces the same answer. And it scales: the same code that handles 40 workbooks today will handle 5,000 next month with no rewrites.

## Current status

- Phase 1 (single-file extraction with QA) — **Done**
- Phase 2 (batch mode + layered QA flags) — **Done**
- Phase 3 (layout-drift resilience, header-verified field detection) — **Done**. All 208 reported feedback issues resolved.
- Phase 4 (parallelization for sub-1-hour runs on 5,000 files) — Next
- Phase 5 (scheduled / cloud pipeline for ongoing weekly ingestion) — Pending volume & cadence decision

---

**Source code:** [github.com/Zubin123/Extract-excel](https://github.com/Zubin123/Extract-excel)
**Stack:** Python 3.12, pandas, openpyxl — no cloud dependencies, runs on any laptop or server.
**Prepared by:** Mohammed Zubin · iBridge Global Services
