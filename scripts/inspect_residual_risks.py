"""Inspect residual risks introduced by the recent fixes."""
import pandas as pd
from pathlib import Path

ROOT = Path(r"c:\Users\MohammedZubinEssudee\OneDrive - iBridge Global Services\Desktop\Extract Excel")
NEW = ROOT / "output" / "verify_run_v4.xlsx"

data = pd.read_excel(NEW, sheet_name="Data")
qa = pd.read_excel(NEW, sheet_name="QA_Summary")
id_conflicts = pd.read_excel(NEW, sheet_name="ID_Conflicts")
unmatched = pd.read_excel(NEW, sheet_name="Unmatched_Sheets")


def banner(t):
    print("\n" + "=" * 88)
    print(t)
    print("=" * 88)


banner("RISK 1 — ID_Conflicts false positives from sheet-name fallback")
print(f"{len(id_conflicts)} conflicts logged. Are they real or fallback-induced?\n")
print(id_conflicts.to_string(index=False))


banner("RISK 2 — Sheet names that aren't really the employee name")
placeholder_rows = qa[qa["Sheet Issues"].fillna("").str.contains("Employee Name placeholder")]
print(f"{len(placeholder_rows)} sheets use sheet-name fallback. Names include suffixes like ' T'?\n")
sample = placeholder_rows[["File", "Sheet Name", "Employee"]].head(20)
print(sample.to_string(index=False))


banner("RISK 3 — Unmatched 'unfilled_template' sheets that might actually contain data")
# Check if any 'unfilled template' sheets have a day grid + filled time rows
print(f"{len(unmatched)} entries on Unmatched_Sheets.")
print("Top reasons:")
print(unmatched["Reason"].value_counts().head().to_string())


banner("RISK 4 — Fields still NOT verified against header")
print("""
The following positions are still hardcoded and not header-checked:
  - Pay-type order within day block:  [RT, OT, DT, PD, 4D, 4A]  (assumed by config)
  - Time row positions:               Start=5, LO=6, LI=7, Stop=8  (assumed)
  - Date row:                         row 3                       (assumed)
  - Day label order:                  MON, TUE, WED, THUR, FRI, SAT, SUN  (assumed)
  - Employee Name (A1):               cell address only           (looks-like-name heuristic)

If any of these drift in a future template, the QA system won't catch the swap.
The grand-totals sum cross-check protects pay-type column reads but not column
*identity* — a template swapping RT and OT positions would balance arithmetically
while putting RT hours in the OT column.
""")
