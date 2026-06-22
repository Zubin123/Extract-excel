# Excel Timesheet Extractor — A Plain-English White Paper

*Turning 5,000 timesheet workbooks into one verified payroll spreadsheet*

A reliable, auditable pipeline that extracts weekly payroll data with zero silent failures — and why this approach beats AI for payroll-grade reliability.

---

## At a glance

| | |
|---|---|
| **5,000+** | files per run, no manual touch |
| **3,896** | verified rows in the latest corpus |
| **208 / 208** | known extraction issues resolved |
| **0** | silent failures — every issue is flagged |

---

## The problem

Every week, dozens of timesheets arrive as Excel workbooks, each carrying multiple employee sheets with hours, pay categories, dates, and classifications. Pulling this into one place — accurately, weekly — used to mean hours of copy-paste. Across 5,000 historical files, manual entry stops being viable.

## What we built

A self-running tool that reads any folder of timesheet Excel files and produces a **single verified payroll spreadsheet** — with every number cross-checked against the source workbook's own formulas. Humans only review the small handful of rows the tool flags.

```
Folder of workbooks   →   Pipeline runs   →   Verified output
(any number of .xlsx)    (reads, verifies,    (5-tab spreadsheet)
                          flags)
```

## How it works — in plain English

Think of it as a very careful intern who never tires, never makes typos, never skips a sheet, and finishes thousands of timesheets in the time it takes to make coffee. The intern follows four rules:

1. **Read every sheet.** Identify which sheets are real employee timesheets vs reference data.
2. **Find each field by its label, not by a fixed position.** If a column shifts or a row gets inserted, the tool still finds the value by reading the header.
3. **Cross-check the math.** Sum Monday-to-Sunday hours per pay category, compare to Excel's own weekly total. A mismatch flags the row — it doesn't silently accept it.
4. **Surface everything.** Output has five clearly labeled tabs — data, accuracy summary, run log, sheets that couldn't be extracted, and ID conflicts. Nothing is hidden.

## Accuracy — what we can guarantee

- ✓ **Deterministic.** Same input → same output, every run.
- ✓ **Self-validating.** Every numeric column is cross-checked.
- ✓ **Auditable.** Each output cell traces back to a source cell.
- ✓ **No silent drops.** Sheets we can't extract are visibly listed with reasons.

## Scale

| Volume | Single-threaded | Parallelized |
|---|--:|--:|
| 100 workbooks | ~17 min | ~2 min |
| 1,000 workbooks | ~3 hours | ~21 min |
| 5,000 workbooks | ~14 hours | ~1.75 hours |
| 10,000+ | ~28 hours | distribute via cloud |

## Where this can be used

- **Weekly payroll preparation** — consolidate timesheets from many crews into one upload-ready file.
- **Year-end audit & compliance** — produce a clean, traceable record of hours worked.
- **Historical data migration** — move years of legacy timesheets into a new HR/payroll system.
- **Cross-site reconciliation** — compare timesheets across job sites or pay periods.
- **Other structured Excel extraction** — same approach works for invoices, expense reports, inspection logs.

## Why not AI or OCR?

We deliberately chose **not** to use an AI model or OCR engine for the main data path:

| | Our approach | AI / LLM approach |
|---|--:|--:|
| Cost per 5,000-file run | $0 | ~$1,500–$15,000 |
| Same input → same output? | Always | No, varies between runs |
| Traceable to source cell? | Every cell | No |
| Can invent values? | Never | Yes (hallucination risk) |
| Speed (5,000 files) | ~2 hours | ~40+ hours |

OCR adds no value when the input is already `.xlsx` — every cell already has structured text. AI is the right tool for some problems; payroll arithmetic is not one of them.

## What "automation" really means here

> **Not magic. Not AI. Just precise rules executed reliably.**
>
> Write down the rules once. Write them precisely. Run them across thousands of files. Surface anything the rules can't safely handle. Boring on purpose — and boring is what makes it reliable enough for payroll.

## The benefits

What used to take a person **days** of manual copying now runs in **hours**, with every number cross-checked, every problem row flagged, and every output cell traceable to its source. No ongoing infrastructure cost. Reproducible. And it scales — the same code that handles 40 workbooks today will handle 5,000 next month with no rewrites.
