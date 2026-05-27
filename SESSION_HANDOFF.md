# Session Handoff — 2026-05-27

> Continuation notes for picking up this project in a new Claude session.
> Read this AFTER `CLAUDE.md` (the project source-of-truth) and alongside
> `NEXT_PHASE_PLAN.md` (the forward architecture plan).
>
> This file captures everything done and learned in the 2026-05-26 → 27
> working session, including corrections to earlier claims in CLAUDE.md.

---

## 1. TL;DR — current state

- Authoritative output: `output/phase2_corpus_v4.xlsx` (regenerated in-place
  multiple times this session; local-only / gitignored).
- The 271 employee sheets currently IN the output are **verified correct**
  against the source workbooks (independent audit, not self-QA).
- **29 "Field Mechanic" employee sheets are silently ABSENT** from the
  output. They contain real timesheets we cannot yet extract. This is the
  biggest open correctness gap. Fix = `NEXT_PHASE_PLAN.md` Phase 1.
- 41 sheets sit on the `Unmatched_Sheets` tab (unverified — could be empty
  templates or more unreadable layouts).

---

## 2. What was done this session (chronological)

1. **WC State / ST position census** — counted distinct on-sheet positions
   for these two fields across the corpus (initial question).

2. **Empty-employee removal** — added a drop filter to `src/extract.py`.
   Iterated the rule three times as the definition was refined:
   - v1: drop if all day rows zero AND no time entries AND all pay-types zero
   - v2: handle string `"0.00"` (not just numeric 0); initially exempted
     NO_GRID employees from dropping (per user request to keep Hall
     OVERHEAD-CA)
   - **v3 (current/final):** drop iff grand TOTALS `Total Hours` is
     0/None/blank/`'0.00'` AND no day row has any time-of-day entry
     (Start/Lunch Out/Lunch In/Stop). Pay-type cells intentionally NOT
     consulted (HP-PTO / OT-only sheets show pay-type totals without
     representing worked hours). NO_GRID exemption REMOVED.
   - Final result: **218 sheets dropped**, 271 kept, 0 NO_GRID in output.
   - Code: `_employee_has_no_activity()` + `_drop_empty_employees()` in
     `src/extract.py` (~line 609-675).

3. **WC State / ST empty-value flagging** — extended `_anchor_issue()` in
   `src/extract.py` (~line 127) to emit a `[CHECK]` flag when the resolved
   WC State or ST VALUE is empty/None (previously only flagged when the
   header was relocated/missing, not when found-but-empty). Three messages:
   - `[CHECK] WC State empty — no value in source cell`
   - `[CHECK] WC State empty — header relocated to <addr> but no value`
   - `[CHECK] WC State empty — header not found in header_row`

4. **NO_GRID investigation** — user correctly suspected NO_GRID sheets were
   a missed layout. Manual inspection (`scripts/inspect_nogrid.py`) proved
   ALL 29 NO_GRID sheets are real "Field Mechanic Time Sheet" timesheets
   with a different layout (see §4). They were being emitted as 8 blank
   rows, then dropped as empty. Their real data is NOT recovered.

5. **Independent audit** (`scripts/audit_v4.py`) — re-read every source
   workbook, did NOT trust our QA. Results in §3.

6. **Full structural census** (`scripts/structural_census.py`) — exhaustive
   count of every employee-like sheet's layout. Results in §4.

7. **Saved `NEXT_PHASE_PLAN.md`** — the architecture plan to move from
   hard-coded profiles to self-describing structural discovery.

---

## 3. Independent audit results (v4, 271 sheets) — VERIFIED, not self-QA

| Check | Result |
|---|---|
| Pay-type grand totals (RT/OT/DT/PD/4D/4A) match source cells | **271/271 correct** |
| WC State matches header-resolved source value | **271/271 correct** |
| ST matches header-resolved source value | **271/271 correct** |
| Total Hours grand total matches source | 207/271 direct match; **64 are `#VALUE!` in source, correctly recovered** by summing day rows → effectively 271/271 |
| Negative hours / >24h-day / impossible values | **0** |
| Listed 17 empty-employee sheets removed | **17/17 dropped** |

**Conclusion:** the data that IS in v4 is correct. The output is INCOMPLETE
(missing 29 Field Mechanic + unverified 41 Unmatched), not WRONG.

Audit script: `scripts/audit_v4.py` — re-run any time with
`.venv\Scripts\python.exe scripts\audit_v4.py`.

---

## 4. STRUCTURAL CENSUS — exact numbers (corrects CLAUDE.md's "2 profiles")

Scanned all 40 workbooks, 721 total sheets. Of these, **502 employee-like
sheets** (219 reference/template/training by name).

### THREE template families, SEVEN distinct full signatures:

| Family | Sheets | Day row | First day col | Stride | Token style | Time-label col | WC State | ST |
|---|---|---|---|---|---|---|---|---|
| **Standard** | 460 | 4 | L | 6 | MON (short) | K | G12 | H12 |
| **Field Mechanic** | 29 | 3 | N | 5 | MONDAY (full) | M | L11 | M11 |
| **Shifted** | 13 | 4 | M | 6 | MON (short) | L | H12 | I12 |

### Identity layout (where name + EE ID live) — 5 variants:
| Pattern | Name | EE ID | Count |
|---|---|---|---|
| standard | A1 | B2 | 435 |
| no-identity | (none) | (none) | 26 |
| placeholder | A1="Input Employee Number" | B2=int | 22 |
| template-empty | A1="Input..." | B2 empty/"EE: #" | 13 |
| alt (Hall) | B1 | C2 | 6 |

### Universal invariants (never vary across all 502 sheets):
- Date row is ALWAYS `day_label_row − 1`
- Time labels ALWAYS start at `day_label_row + 1`
- Pay-type columns are NEVER headered with RT/OT/DT text — order is purely
  positional. **This means our YAML assumption of [RT,OT,DT,PD,4D,4A] order
  cannot be cross-checked from sheet content.** It's an unverified assumption.

### Field Mechanic family detail (the unhandled one):
- Appears in **14 workbooks**: Coulson (×2 files), Pottage (×2), Beeler (×2),
  Medina (×2), Hall (×2), Smith Dustin, Girtman, + JN-Correction variants.
- Employees: Coulson 3047, Pottage 8540, Beeler 1586, Medina 7475,
  Hall 4876 (alt-layout B1/C2), Smith 9489 (alt), Girtman 4683.
- Title cell Z1 = "FIELD MECHANIC TIME SHEET".
- Day labels MONDAY..SUNDAY at row 3, cols N S X AC AH AM AR.
- Date row 2; time labels in col M, rows 4-7 (Start/Lunch Out/Lunch In/Stop).
- Sheet names: `Job`, `OVERHEAD- CA`, `Girtman - CA/JOBS/WA`, `WE 12 18 21`,
  `WE 12 18 Correction`.
- Pay-totals row + grand-total columns + pay_block_width NOT yet verified —
  must inspect a Field Mechanic sheet before adding a profile.

Census script: `scripts/structural_census.py`.

---

## 5. HONEST assessment of the QA system (user asked "is QA a comedy?")

The only HARD QA gate is the pay-type sum cross-check (day sums == grand
totals). That is real and load-bearing — `PASS` rows genuinely have
consistent pay-type math.

But QA had NO independent check for: Date, Employee Name, EE ID, WC State*,
ST*, Sheet Name, or extraction COMPLETENESS. (*WC/ST get header-label
verification but no value cross-check against a roster.)

The `PASS` badge OVERSTATES what was validated — it only speaks to the 6
pay-type columns, not the whole row. NO_GRID was treated as a benign status,
so 29 sheets of blank-but-real-data rows passed through looking like valid
records.

### What real QA needs (not built yet — `NEXT_PHASE_PLAN.md` Phase 3):
1. **Roster cross-check** — every workbook has an `Employees` sheet mapping
   EE ID → name / WC State / ST. Use it as independent ground truth. This
   alone would catch the identity-field mismatches.
2. **Date sequence as a HARD gate** — filename WE date implies MON..SUN
   dates; mismatch = FAIL, not [INFO].
3. **Completeness check** — a sheet producing all-blank identity/time fields
   = FAIL, not NO_GRID.

---

## 6. Files changed / created this session

Modified:
- `src/extract.py` — added `_employee_has_no_activity`, `_drop_empty_employees`
  (wired into `extract()` before Run_Info build); extended `_anchor_issue`
  for empty-value flagging; added "Sheets Dropped (empty)" to Run_Info.

Created (analysis scripts — all read-only, safe to re-run):
- `scripts/count_wcst_layouts.py` — WC/ST position census
- `scripts/inspect_nogrid.py` — manual NO_GRID sheet inspection
- `scripts/inspect_coulson.py`, `scripts/find_mon_coulson.py` — Field Mechanic probe
- `scripts/audit_v4.py` — independent source-vs-output audit
- `scripts/structural_census.py` — full layout census (the §4 numbers)
- `scripts/inspect_feedback_v4.py`, `scripts/analyze_feedback_v4.py`,
  `scripts/check_expected_fm.py`, `scripts/verify_v4_drops.py`,
  `scripts/why_kept.py` — feedback-file analysis + drop verification

Created (docs):
- `NEXT_PHASE_PLAN.md` — forward architecture (structural discovery)
- `SESSION_HANDOFF.md` — this file

Regenerated:
- `output/phase2_corpus_v4.xlsx` (in-place; gitignored)

Feedback file present at project root:
- `phase2_corpus_v4 - Feedback.xlsx` — user/manager annotations. Structure:
  cols 1-25 = our v4 output, col 26 blank, cols 27-48 = expected output,
  col 49 blank, cols 50-63 = per-field TRUE/FALSE match grid. The expected
  block has NO rows for the 6 Field Mechanic Job/OVERHEAD sheets (confirms
  they should be dropped from CURRENT output until extraction is fixed).

---

## 7. Corrections to CLAUDE.md (update CLAUDE.md when convenient)

- CLAUDE.md §5/§6 say "2 profiles (standard + shifted)". **WRONG** — there
  are 3 families; Field Mechanic (29 sheets) is undocumented and unhandled.
- CLAUDE.md §4 calls `Job` / `OVERHEAD- CA` "no-grid placeholder sheets".
  **WRONG for the 14 Field Mechanic workbooks** — there those ARE the real
  employee timesheets (Coulson, Pottage, etc.), just in a layout we can't read.
- CLAUDE.md §5 result counts (425 PASS / 35 REVIEW / 29 NO_GRID) are from a
  PRE-drop-filter run. Current v4 = 251 PASS / 20 REVIEW / 0 NO_GRID / 271
  total employee sheets / 218 dropped.

---

## 8. Open decisions awaiting user/manager (from earlier in session)

1. Pay-type / OT semantics per company (RT+DT vs RT+OT+DT) — blocks a true
   Total Hours cross-check. CLAUDE.md §7.
2. PTO / HP columns — always emitted blank; where do they live in source?
3. WC State vs ST business meaning — needed for domain validation.
4. Is the `Employees` roster sheet reliably present in every workbook?
   (It IS present in the Field Mechanic files — saw it during inspection.)
5. Folder Name hardcoded to "2022" — should come from path?
6. Are the 41 Unmatched_Sheets truly empty templates? Needs manual eyeball.
7. Operator failure-mode preference: refuse / best-effort / halt.
8. Future 5,000-file batches: same companies? same years? template revisions?

---

## 9. Recommended next action (resume point)

**Start `NEXT_PHASE_PLAN.md` Phase 1** — make the extractor read the Field
Mechanic family:

1. In `src/anchors.py`: broaden `DAY_TOKENS` to include full day names
   (MONDAY..SUNDAY). `find_day_label_row` already scans rows 1-20 so row 3
   will be found once tokens match.
2. Add a `field_mechanic` profile to `config/schema.yaml`:
   - day_label_row: 3, date_row: 2
   - day_start_cols: N S X AC AH AM AR (stride 5)
   - time_rows: {Start: 4, Lunch Out: 5, Lunch In: 6, Stop: 7}
   - anchors: wc_state near "WC STATE" (L11), st near "ST" (M11),
     name A1 (or B1 for alt), ee_id B2 (or C2 for alt)
   - **MUST FIRST inspect a Field Mechanic sheet to find: pay_block_width,
     pay-type column order, grand_totals columns, and the pay-totals row.**
     Do NOT assume width 6 / [RT,OT,DT,PD,4D,4A] — verify it.
3. Update `select_profile` detect rules for day_label_first_col = N.
4. Re-run extractor; the 29 sheets should now extract real hours and pass
   the sum cross-check. Verify with `scripts/audit_v4.py`.
5. The alt-layout (B1/C2) Hall + Smith sheets need the identity path that
   already exists (`employee_alt_layout`) combined with the new profile.

After Phase 1, the 26 "no-identity" Standard-family sheets also deserve a
look (could be unfilled templates or another identity variant).

---

## 10. Environment / run commands

- Working dir: `C:\dev\Extract-excel` (NOT OneDrive — corrupts git packs)
- Python: `.venv\Scripts\python.exe` (call directly; `uv run` may not resolve)
- Run extractor: `.venv\Scripts\python.exe src\extract.py data\ --out output\phase2_corpus_v4.xlsx`
- Run audit: `.venv\Scripts\python.exe scripts\audit_v4.py`
- Run census: `.venv\Scripts\python.exe scripts\structural_census.py`
- Data lives in `data\sample\` and `data\sample1\` (gitignored; copy from
  shared drive on a new machine).
- Git: branch `main`, remote https://github.com/Zubin123/Extract-excel.git
  Last committed HEAD before this session: 64367e3. The drop-filter and
  flagging changes from this session are NOT yet committed.
