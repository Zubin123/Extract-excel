"""Verify the header-anchor fix solves Claims 5 and 6 (and exposes the hidden ST bug)."""
import pandas as pd
from pathlib import Path

ROOT = Path(r"c:\Users\MohammedZubinEssudee\OneDrive - iBridge Global Services\Desktop\Extract Excel")
NEW = ROOT / "output" / "verify_run_v2.xlsx"

data = pd.read_excel(NEW, sheet_name="Data")
qa = pd.read_excel(NEW, sheet_name="QA_Summary")

print("=" * 90)
print("CLAIM 5 — WE 12 25 21 6818 UPL: WC State should now read 'CA' (from H13), not 'N' (G13)")
print("=" * 90)
target = "WE 12 25 21 6818"
rows = data[data["Excel Name"].str.contains(target, na=False)]
sample = rows[["Sheet Name", "WC State", "ST"]].drop_duplicates().head(15)
print(sample.to_string(index=False))
print(f"\nUnique WC State values for this file: {sorted(rows['WC State'].dropna().unique().tolist())}")
print(f"Unique ST values for this file:       {sorted(rows['ST'].dropna().unique().tolist())}")

print()
print("=" * 90)
print("CLAIM 6 — WC State coverage across all files")
print("=" * 90)
wc_counts = data["WC State"].value_counts(dropna=False)
print(wc_counts.to_string())

print()
print("=" * 90)
print("ANCHOR [CHECK] FLAGS — sheets whose WC State / ST anchor was relocated or missing")
print("=" * 90)
anchor_qa = qa[qa["Sheet Issues"].fillna("").str.contains("anchor")]
print(f"{len(anchor_qa)} sheet(s) with anchor [CHECK] flag\n")
print(anchor_qa[["File", "Sheet Name", "Overall", "Sheet Issues"]].head(20).to_string(index=False))

print()
print("=" * 90)
print("OVERALL STATUS DISTRIBUTION")
print("=" * 90)
print(qa["Overall"].value_counts().to_string())
