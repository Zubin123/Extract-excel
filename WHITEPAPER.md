# Excel Timesheet Extractor — A Plain-English White Paper

A reliable, auditable pipeline that turns a folder of weekly timesheet workbooks into a single verified payroll spreadsheet, with zero silent failures.

---

## At a glance

| | |
|---|---|
| **5,000+** | workbooks per run |
| **3,896** | verified rows in the latest corpus |
| **208 / 208** | known extraction issues resolved |
| **0** | silent failures — every issue is flagged |

---

## The problem

Every week, dozens of timesheets arrive as Excel workbooks, each carrying multiple employee sheets with hours, pay categories, dates, and worker classifications. Templates vary slightly between weeks and offices. Pulling all this data into one place — accurately, with every number verified — is the job this tool does.

## What we built

A self-running tool that reads any folder of timesheet Excel files and produces a **single verified payroll spreadsheet**. Every number is cross-checked against the source workbook's own formulas. Humans only review the small handful of rows the tool flags for attention.

```
Folder of workbooks   →   Pipeline runs   →   Verified output
(any number of .xlsx)    (reads, verifies,    (5-tab spreadsheet)
                          flags)
```

## What the output looks like

One Excel file with five clearly labeled tabs — nothing is hidden:

| Tab | What's in it |
|---|---|
| **Data** | Every employee, every day, every pay category — one row per day plus a weekly TOTALS row. |
| **QA Summary** | Per-employee accuracy check, colour-coded green / yellow / red so you can see at a glance which rows need attention. |
| **Run Info** | When it ran, on which files, how many sheets passed, the overall result. |
| **Unmatched Sheets** | Sheets the pipeline refused to extract, with a clear reason for each. |
| **ID Conflicts** | The same employee ID appearing under different names in the same workbook. |

## How it works — in plain English

The pipeline follows four rules:

1. **Read every sheet.** Identify which sheets are real employee timesheets versus reference/lookup sheets.
2. **Find each field by its label, not by a fixed position.** If a column gets shifted or a row gets inserted into a future template, the tool still finds the value by reading the header text above it.
3. **Cross-check the math.** Sum Monday-to-Sunday hours per pay category and compare to Excel's own weekly total. A mismatch flags the row for review — the tool never silently accepts wrong numbers.
4. **Surface everything.** Anything the rules can't safely handle gets recorded on a visible tab with a reason. There are no silent drops, no quiet failures.

## Accuracy — what we can guarantee

- ✓ **Deterministic.** Same input always produces the same output.
- ✓ **Self-validating.** Every numeric column is cross-checked against Excel's own totals.
- ✓ **Auditable.** Each output cell traces back to a specific source cell.
- ✓ **No silent drops.** Any sheet the pipeline can't extract is listed with a reason.

## Where this can be used

- **Weekly payroll preparation** — consolidate timesheets from many crews into one upload-ready file.
- **Year-end audit & compliance** — produce a clean, traceable record of hours worked.
- **Historical data migration** — move years of legacy timesheets into a new HR or payroll system.
- **Cross-site reconciliation** — compare timesheets across job sites, companies, or pay periods.
- **Other structured Excel extraction** — the same approach works for invoices, expense reports, inspection logs, and similar tabular data.

## What "automation" really means here

> **Not magic. Just precise rules executed reliably.**
>
> Write down the rules once. Write them precisely. Run them across many files. Surface anything the rules can't safely handle. Boring on purpose — and boring is what makes it reliable enough for payroll.

## The benefits

Every number is cross-checked against the source. Every problem row is flagged for human review. Every output cell traces back to its source cell. The pipeline is reproducible — running it again on the same files produces the same answer. And the same logic that handles a small folder today works equally well on a much larger one tomorrow.
