"""Show the current output's Total Hours column."""
import pandas as pd
from pathlib import Path

OUT = Path(__file__).parent.parent / "output" / "extracted.xlsx"
df = pd.read_excel(OUT, sheet_name="Data")
qa = pd.read_excel(OUT, sheet_name="QA_Summary")

print("=== Data sheet: Total Hours and QA_Flag ===")
print(df[["Employee Name", "Day", "Start", "Stop", "Total Hours", "QA_Flag"]].to_string())
print("\n=== QA_Summary: Total Hours columns ===")
print(qa[["Employee", "Total Hours_daily_sum", "Total Hours_grand_total", "Total Hours_match", "Overall"]].to_string())
