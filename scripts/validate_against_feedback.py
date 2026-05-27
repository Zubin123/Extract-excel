"""Validate the new output (v4) against the user's feedback expectations."""
import pandas as pd
from pathlib import Path
import warnings
warnings.filterwarnings("ignore")

ROOT = Path(r"c:\Users\MohammedZubinEssudee\OneDrive - iBridge Global Services\Desktop\Extract Excel")
NEW = ROOT / "output" / "phase2_corpus_v4.xlsx"
FB = ROOT / "phase2_corpus_v3 - Feedback.xlsx"

new = pd.read_excel(NEW, sheet_name="Data")
fb = pd.read_excel(FB, sheet_name="Data")

# Keep only feedback rows with comments
commented = fb[fb["Comments 1"].notna()].copy()
print(f"Feedback has {len(commented)} commented rows. Checking each...\n")

# Build lookup key (Excel Name, Sheet Name, Day) in new
def make_key(df, sheet_col="Sheet Name", file_col="Excel Name"):
    return list(zip(df[file_col].astype(str), df[sheet_col].astype(str), df["Day"].astype(str)))

new["key"] = make_key(new)
commented["key"] = make_key(commented)
new_lookup = new.set_index("key")[["Employee Name", "WC State", "ST"]].to_dict("index")

# For each comment category, count how many are now fixed
categories = commented["Comments 1"].unique()
for cat in categories:
    rows = commented[commented["Comments 1"] == cat]
    fixed = 0
    still_wrong = []
    for _, r in rows.iterrows():
        key = r["key"]
        if cat == "All Fields Data Entraction is Missing":
            f, s, _ = key
            matches = new[(new["Excel Name"] == f) & (new["Sheet Name"] == s)]
            if len(matches) == 0:
                still_wrong.append((key, "row missing from new output"))
                continue
            row0 = matches.iloc[0]
            expected_emp = r.get("Employee Name.1")
            expected_wc = r.get("WC STATE")
            expected_st = r.get("ST.1")
            details = []
            if pd.notna(expected_emp) and str(expected_emp) != str(row0["Employee Name"]):
                details.append(f"emp expected={expected_emp!r} got={row0['Employee Name']!r}")
            if pd.notna(expected_wc) and str(expected_wc) != str(row0["WC State"]):
                details.append(f"WC expected={expected_wc!r} got={row0['WC State']!r}")
            if pd.notna(expected_st) and str(expected_st) != str(row0["ST"]):
                details.append(f"ST expected={expected_st!r} got={row0['ST']!r}")
            if not details:
                fixed += 1
            else:
                still_wrong.append((key, "; ".join(details)))
            continue
        if key not in new_lookup:
            still_wrong.append((key, "row missing from new output"))
            continue
        new_row = new_lookup[key]
        if cat == "WC State field not Extracted":
            expected = r.get("WC STATE")
            actual = new_row["WC State"]
            ok = (pd.notna(expected) and str(expected) == str(actual)) or (pd.isna(expected) and (pd.isna(actual) or actual == ""))
            if ok:
                fixed += 1
            else:
                still_wrong.append((key, f"expected={expected!r}, got={actual!r}"))
        elif cat == "ST field not Extracted":
            expected = r.get("ST.1")
            actual = new_row["ST"]
            ok = (pd.notna(expected) and str(expected) == str(actual)) or (pd.isna(expected) and (pd.isna(actual) or actual == ""))
            if ok:
                fixed += 1
            else:
                still_wrong.append((key, f"expected={expected!r}, got={actual!r}"))
        elif cat == "Employee Name Populated from Sheet Name":
            # User expects nan (blank). Pass if blank.
            actual = new_row["Employee Name"]
            ok = pd.isna(actual) or actual == ""
            if ok:
                fixed += 1
            else:
                still_wrong.append((key, f"expected=blank, got={actual!r}"))
        elif cat == "All Fields Data Entraction is Missing":
            # Feedback rows for this category have Day=nan; match on file+sheet instead
            f, s, _ = key
            matches = new[(new["Excel Name"] == f) & (new["Sheet Name"] == s)]
            if len(matches) == 0:
                still_wrong.append((key, "row missing from new output"))
                continue
            row0 = matches.iloc[0]
            expected_emp = r.get("Employee Name.1")
            expected_wc = r.get("WC STATE")
            expected_st = r.get("ST.1")
            details = []
            if pd.notna(expected_emp) and str(expected_emp) != str(row0["Employee Name"]):
                details.append(f"emp expected={expected_emp!r} got={row0['Employee Name']!r}")
            if pd.notna(expected_wc) and str(expected_wc) != str(row0["WC State"]):
                details.append(f"WC expected={expected_wc!r} got={row0['WC State']!r}")
            if pd.notna(expected_st) and str(expected_st) != str(row0["ST"]):
                details.append(f"ST expected={expected_st!r} got={row0['ST']!r}")
            if not details:
                fixed += 1
            else:
                still_wrong.append((key, "; ".join(details)))
    print(f"\n[{cat}] — {fixed}/{len(rows)} fixed")
    for key, why in still_wrong[:6]:
        print(f"  STILL WRONG: {key}  ->  {why}")
    if len(still_wrong) > 6:
        print(f"  ... and {len(still_wrong) - 6} more")
