"""Check that single-employee files (Chris Coulson etc.) where row 12 has no
header labels at all are properly surfaced too.
"""
import pandas as pd
from pathlib import Path

ROOT = Path(r"c:\Users\MohammedZubinEssudee\OneDrive - iBridge Global Services\Desktop\Extract Excel")
NEW = ROOT / "output" / "verify_run_v2.xlsx"

data = pd.read_excel(NEW, sheet_name="Data")
qa = pd.read_excel(NEW, sheet_name="QA_Summary")

print("Chris Coulson rows — what does the totals row's QA_Flag look like?")
mask = data["Excel Name"].str.contains("Chris Coulson", na=False)
totals = data[mask & (data["Day"] == "TOTALS")]
print(totals[["Excel Name", "Sheet Name", "WC State", "ST", "QA_Flag"]].to_string(index=False))

print("\nQA_Summary entries for Chris Coulson files:")
qa_mask = qa["File"].str.contains("Chris Coulson", na=False)
print(qa[qa_mask][["File", "Sheet Name", "Overall", "Sheet Issues"]].to_string(index=False))

print("\nAll sheets where WC State is blank in the Data tab:")
blank_mask = data["WC State"].isna() | (data["WC State"].astype(str).str.strip() == "")
files_with_blanks = data[blank_mask].groupby(["Excel Name", "Sheet Name"]).size().reset_index(name="blank_rows")
print(files_with_blanks.head(30).to_string(index=False))
