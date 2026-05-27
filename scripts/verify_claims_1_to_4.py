"""Verify Claims 1-4 are now fixed (extracted) or visible (Unmatched_Sheets)."""
import pandas as pd
from pathlib import Path

ROOT = Path(r"c:\Users\MohammedZubinEssudee\OneDrive - iBridge Global Services\Desktop\Extract Excel")
NEW = ROOT / "output" / "verify_run_v4.xlsx"

data = pd.read_excel(NEW, sheet_name="Data")
qa = pd.read_excel(NEW, sheet_name="QA_Summary")
unmatched = pd.read_excel(NEW, sheet_name="Unmatched_Sheets")


def banner(t):
    print("\n" + "=" * 90)
    print(t)
    print("=" * 90)


banner("CLAIM 1 — WE 12 18 21 CA 6427: the 6 claimed missed sheets")
target = "WE 12 18 21 CA 6427"
claimed = ["Delgado", "Griego", "Nunez", "McDonald", "Benedict", "Rivera"]
file_rows = data[data["Excel Name"].str.contains(target, na=False)]
file_qa = qa[qa["File"].str.contains(target, na=False)]
print(f"Data tab unique sheet names for this file ({file_rows['Sheet Name'].nunique()} total):")
print(sorted(file_rows["Sheet Name"].dropna().unique().tolist()))
print()
for name in claimed:
    rows = file_rows[file_rows["Sheet Name"].str.fullmatch(rf"{name}\s*", na=False)
                     | file_rows["Sheet Name"].str.fullmatch(name, na=False)]
    if len(rows) == 0:
        # also check loose match
        rows = file_rows[file_rows["Sheet Name"].str.contains(name, na=False)]
    if len(rows) > 0:
        emp_name = rows["Employee Name"].iloc[0]
        ee_id = rows["EE ID"].iloc[0]
        wc = rows["WC State"].iloc[0]
        print(f"  OK '{name}' EXTRACTED: emp_name={emp_name!r}, EE ID={ee_id}, WC State={wc!r}, rows={len(rows)}")
    else:
        print(f"  XX '{name}' still missing from Data tab")


banner("CLAIM 4 — WE 12 25 21 6562 UPL: Hanson")
target = "WE 12 25 21 6562 UPL -WA 4937"
hanson_rows = data[data["Excel Name"].str.contains(target, na=False) & data["Sheet Name"].str.contains("Hanson", na=False)]
if len(hanson_rows) > 0:
    print(f"  OK Hanson EXTRACTED: emp_name={hanson_rows['Employee Name'].iloc[0]!r}, "
          f"EE ID={hanson_rows['EE ID'].iloc[0]}, rows={len(hanson_rows)}")
    print(f"  Sample rows:")
    print(hanson_rows[["Day", "Date", "Total Hours", "RT", "QA_Flag"]].to_string(index=False))
else:
    print("  XX Hanson still missing")


banner("CLAIM 2 & 3 — Hall files: OVERHEAD-CA")
for tag in ["WE 12 18 21 Hall", "WE 12 25 21 Hall"]:
    print(f"\nFile: {tag}*")
    um_rows = unmatched[unmatched["File"].str.contains(tag, na=False)]
    if len(um_rows) > 0:
        print(f"  Now visible on Unmatched_Sheets ({len(um_rows)} entries):")
        print(um_rows.to_string(index=False))
    else:
        print("  Nothing on Unmatched_Sheets")
    data_rows = data[data["Excel Name"].str.contains(tag, na=False)]
    if len(data_rows) > 0:
        print(f"  Data tab entries: {sorted(data_rows['Sheet Name'].dropna().unique().tolist())}")
    else:
        print("  Data tab: (none)")


banner("OVERALL OUTPUT STATUS")
print(f"Data rows:              {len(data)}")
print(f"QA_Summary entries:     {len(qa)}")
print(f"Unmatched_Sheets count: {len(unmatched)}")
print(f"\nQA Overall distribution:")
print(qa["Overall"].value_counts().to_string())
print(f"\nUnmatched reasons (top 5):")
print(unmatched["Reason"].value_counts().head(5).to_string())
