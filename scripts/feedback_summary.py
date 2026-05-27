"""Aggregate the user's feedback into discrete issue counts and source samples."""
import pandas as pd
from collections import defaultdict
from openpyxl import load_workbook
from pathlib import Path
import warnings
warnings.filterwarnings("ignore")

ROOT = Path(r"c:\Users\MohammedZubinEssudee\OneDrive - iBridge Global Services\Desktop\Extract Excel")
FB = ROOT / "phase2_corpus_v3 - Feedback.xlsx"
DATA = ROOT / "data"

df = pd.read_excel(FB, sheet_name="Data")
commented = df[df["Comments 1"].notna()].copy()
print(f"Total rows with comments: {len(commented)}")
print()

# Group by exact comment text
counts = commented["Comments 1"].value_counts()
print("Comment type breakdown:")
for c, n in counts.items():
    print(f"  {n:4d}  {c!r}")

# For each comment category, show one example file/sheet
print()
print("Sample (file, sheet) for each comment category:")
for c in counts.index:
    sample = commented[commented["Comments 1"] == c].iloc[0]
    f = sample.get("Excel Name", "")
    s = sample.get("Sheet Name", "")
    extracted_wc = sample.get("WC State", "")
    extracted_st = sample.get("ST", "")
    extracted_emp = sample.get("Employee Name", "")
    exp_emp = sample.get("Employee Name.1", "")
    exp_wc = sample.get("WC STATE", "")
    exp_st = sample.get("ST.1", "")
    print(f"\n  [{c}]")
    print(f"    File:  {f}")
    print(f"    Sheet: {s}")
    print(f"    Extracted: emp={extracted_emp!r}, WC={extracted_wc!r}, ST={extracted_st!r}")
    print(f"    Expected:  emp={exp_emp!r}, WC={exp_wc!r}, ST={exp_st!r}")

# List files/sheets where WC State extraction is missing
print("\n\n=== WC State missing — unique (file, sheet) pairs ===")
wc_missing = commented[commented["Comments 1"].astype(str).str.contains("WC State", na=False)]
pairs = wc_missing[["Excel Name", "Sheet Name"]].drop_duplicates()
print(f"{len(pairs)} unique (file, sheet) with WC State missing:")
print(pairs.to_string(index=False))

print("\n\n=== ST missing — unique (file, sheet) pairs ===")
st_missing = commented[commented["Comments 1"].astype(str).str.contains(r"\bST\b", na=False, regex=True)]
pairs = st_missing[["Excel Name", "Sheet Name"]].drop_duplicates()
print(f"{len(pairs)} unique (file, sheet) with ST missing:")
print(pairs.to_string(index=False))
