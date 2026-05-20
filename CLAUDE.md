# CLAUDE.md — Instructions

Build a Python script that converts a weekly-timesheet Excel workbook into a flat
tabular output matching a given template.

---

## 1. Files you'll work with

Two Excel files (provided separately, place them in `data/` — see project layout below):

- **Input sample**: `WE_01_08_22_CA_GF_S_-_WA_4939-0899-6263_1.xlsx`
  → the source workbook to read from.
- **Output template**: `Header_Template_New_Rev_1.xlsx`
  → defines the exact column order and formatting of the output. Do not modify it; only read its header row to match the schema.

---

## 2. Project layout — set this up first

```
timesheet-poc/
├── CLAUDE.md                ← this file
├── pyproject.toml           ← created by `uv init`
├── uv.lock                  ← created by uv
├── .venv/                   ← gitignored
├── .gitignore               ← include .venv/, output/, __pycache__/
├── data/
│   ├── WE_01_08_22_CA_GF_S_-_WA_4939-0899-6263_1.xlsx   ← input
│   └── Header_Template_New_Rev_1.xlsx                   ← template
├── src/
│   └── extract.py           ← the script you'll build
├── output/
│   └── (generated files go here)
└── tests/
    └── test_extract.py
```

---

## 3. Environment setup — use `uv` + `.venv`

Always use `uv`. Never `pip install` into system Python. Never activate the venv manually — use `uv run` for every command.

```bash
# One-time setup
uv init --python 3.12
uv venv
uv add openpyxl pandas
uv add --dev pytest

# Running anything
uv run python src/extract.py data/WE_01_08_22_CA_GF_S_-_WA_4939-0899-6263_1.xlsx
uv run pytest tests/
```

Commit `pyproject.toml` and `uv.lock`. Gitignore `.venv/`.

---

## 4. Input file structure (already verified — do not re-derive)

The input workbook has 16 sheets. **Skip** these 6 reference sheets:

```python
REFERENCE_SHEETS = {
    "Master Template (9) JOB",
    "Certified Codes",
    "Employees",
    "Job Details",
    "Job List",
    "ColumnLists",
}
```

The remaining 10 sheets are one weekly timesheet per employee (e.g. `Case`, `Sanders`, `Raper ` — note: preserve trailing spaces in sheet names).

Every employee sheet uses the **same fixed grid**:

| Cell / range                       | Meaning                                           |
|-----------------------------------|---------------------------------------------------|
| `A1`                              | Employee Name (`"Last; First M"`)                 |
| `B2`                              | Employee ID (int)                                 |
| Row 3 at day-block start cols     | Date for that day                                 |
| Row 4 at day-block start cols     | Day label (`MON`, `TUE`, etc.)                    |
| Row 5                             | Start time per day                                |
| Row 6                             | Lunch Out                                          |
| Row 7                             | Lunch In                                           |
| Row 8                             | Stop                                               |
| Row 9                             | Total Hours per day (may be `#VALUE!` if OFF)     |
| Row 13, col G                     | WC State                                          |
| Row 13, col H                     | ST                                                |
| Row 24                            | Per-day totals for RT/OT/DT/PD/4D/4A              |
| `BB24, BC24, BD24, BE24, BF24, BG24` | Grand weekly totals (RT, OT, DT, PD, 4D, 4A)   |

Day blocks (6 columns each: RT, OT, DT, PD, 4D, 4A):

```python
DAY_BLOCKS = [
    ("MON",  "L",  "Q"),
    ("TUE",  "R",  "W"),
    ("WED",  "X",  "AC"),
    ("THUR", "AD", "AI"),
    ("FRI",  "AJ", "AO"),
    ("SAT",  "AP", "AU"),
    ("SUN",  "AV", "BA"),
]
```

Load with `openpyxl.load_workbook(path, data_only=True)` so formulas return computed values.

---

## 5. Output format (read headers from `Header_Template_New_Rev_1.xlsx`)

23 columns in this exact order:

```
Folder Name | Box Name | Excel Name | Sheet Name | Employee Name | EE ID |
WC State | ST | Date | Day | Start | Lunch Out | Lunch In | Stop |
Total Hours | RT | OT | DT | PD | 4D | 4A | PTO | HP
```

### Rows to emit per employee
- 7 day rows (MON → SUN, in that order — note SUN's date is the day *before* MON's)
- 1 TOTALS row (`Day = "TOTALS"`, no Date / no time fields, hour totals from `BB24:BG24`)

So 10 employees × 8 rows = **80 rows** for this sample.

### Formatting
- `Date` → `MM/DD/YY` (e.g. `01/03/22`)
- Times → `H:MM AM/PM` (e.g. `6:00 AM`); blank when the source cell isn't a time (e.g. `"OFF"`)
- Hour columns → numeric; treat `None` / blank as `0.0`; treat `#VALUE!` as `None`
- `PTO` and `HP` → leave blank (not present in sample)
- `Folder Name` → `"2022"` (CLI arg, default)
- `Box Name` → `"WE 01 08 22"` (inferred from filename, or CLI arg)
- `Excel Name` → basename of the input file
- `Sheet Name` → the employee sheet's name

---

## 6. Accuracy check (build this in)

For each employee, after extraction, verify:

```
sum(MON..SUN of column X)  ==  TOTALS-row column X
```

for each of: **RT, OT, DT, PD, 4D, 4A**. Tolerance 0.01. Print a per-employee pass/fail table to the console.

Expected on the sample input: **10/10 employees pass.**

---

## 7. CLI

```bash
uv run python src/extract.py data/WE_01_08_22_CA_GF_S_-_WA_4939-0899-6263_1.xlsx \
    --folder "2022" --box "WE 01 08 22" --out output/extracted.xlsx
```

Exit `0` on all-pass, `2` if any employee fails accuracy.

---

## 8. Test

Write `tests/test_extract.py` with at least:
- Output has exactly 80 rows.
- All 10 employees pass the accuracy check.
- Spot check: Case TUE RT == 8.0, Case TOTALS RT == 32.0.

Run with `uv run pytest`.

---

## 9. Rules

- **No LLM at runtime.** The grid is fixed, cell coordinates are known. Pure Python only.
- **Don't re-derive the layout.** Section 4 is the verified spec. If something seems wrong, ask the user.
- **Preserve trailing spaces** in sheet names.
- **Don't hand-edit `pyproject.toml`** — use `uv add`.
- **Don't modify the template file.** Only read its header row.

---

## 10. Done when

- `uv run python src/extract.py data/WE_01_08_22_CA_GF_S_-_WA_4939-0899-6263_1.xlsx` produces `output/extracted.xlsx`
- Console shows `10/10 employees passed`
- `uv run pytest` is green
- The output file's column order matches `Header_Template_New_Rev_1.xlsx` row 1 exactly
