"""Look at the Comments column and side-by-side expected data in feedback."""
import pandas as pd
from pathlib import Path
import warnings
warnings.filterwarnings("ignore")

ROOT = Path(r"c:\Users\MohammedZubinEssudee\OneDrive - iBridge Global Services\Desktop\Extract Excel")
FB = ROOT / "phase2_corpus_v3 - Feedback.xlsx"

df = pd.read_excel(FB, sheet_name="Data")

# Show rows with non-null Comments 1
commented = df[df["Comments 1"].notna()]
print(f"Rows with comments: {len(commented)}")
if len(commented) > 0:
    for _, row in commented.head(20).iterrows():
        print(f"\n--- {row.get('Excel Name', '')} | {row.get('Sheet Name', '')} | {row.get('Day', '')} ---")
        print(f"  Extracted Emp Name: {row.get('Employee Name')!r}")
        print(f"  Extracted WC State: {row.get('WC State')!r}   ST: {row.get('ST')!r}")
        print(f"  Comment: {row['Comments 1']!r}")
        # Look at expected side
        if pd.notna(row.get("Employee Name.1")):
            print(f"  Expected Emp Name:  {row.get('Employee Name.1')!r}")
        if pd.notna(row.get("WC STATE")):
            print(f"  Expected WC State:  {row.get('WC STATE')!r}   ST: {row.get('ST.1')!r}")

# Also look at Hall files specifically
print("\n\n=== HALL FILES IN FEEDBACK ===")
hall = df[df["Excel Name"].astype(str).str.contains("Hall", na=False)]
print(f"Hall rows: {len(hall)}")
for _, row in hall.head(10).iterrows():
    print(f"  {row.get('Sheet Name', '')!r:25} | EmpName={row.get('Employee Name')!r} | EE={row.get('EE ID')!r} | WC={row.get('WC State')!r} | ST={row.get('ST')!r}")

# Look at the "expected" Hall data on the right side
print("\n=== HALL FILES EXPECTED (right-hand columns) ===")
hall_expected = df[df["Excel Name.1"].astype(str).str.contains("Hall", na=False)]
print(f"Hall rows in expected block: {len(hall_expected)}")
for _, row in hall_expected.head(10).iterrows():
    print(f"  {row.get('Sheet Name.1', '')!r:25} | EmpName={row.get('Employee Name.1')!r} | EE={row.get('EE ID #')!r} | WC={row.get('WC STATE')!r}")

# Now look at Delgado etc — what's the EXPECTED employee name?
print("\n=== DELGADO / HANSON / etc EXPECTED NAMES (from feedback) ===")
target_sheets = ["Delgado", "Griego", "Nunez", "McDonald", "Benedict", "Rivera", "Hanson"]
for sn in target_sheets:
    match = df[df["Sheet Name"].astype(str).str.strip() == sn]
    if len(match) > 0:
        r = match.iloc[0]
        extracted_emp = r.get("Employee Name")
        expected_emp = r.get("Employee Name.1")
        print(f"  Sheet={sn!r:12} | Extracted={extracted_emp!r:25} | Expected={expected_emp!r}")
