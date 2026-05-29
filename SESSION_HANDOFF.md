# Session Handoff — 2026-05-29

> Continuation notes for picking up this project in a new Claude session.
> Read this AFTER `CLAUDE.md` (the project source-of-truth).
>
> Supersedes the prior 2026-05-27 handoff. The architectural work that
> handoff was setting up is DONE and manager-validated.

---

## 1. TL;DR — current state

- **Authoritative output:** `output/corpus_v5.xlsx` (gitignored). Manager has
  reviewed this output and accepted it.
- **Accuracy vs `phase2_corpus_v4 - Feedback.xlsx`** (the manager's annotated
  expected output): **15 of 16 fields at 100%**. The only 0.8% miss is one
  Wirth MON PTO cell (source = `19.98`, manager rounded expected to `20.00`).
  Not an extraction error; we report source faithfully.
- **Architecture:** every output field — RT/OT/DT/PD/4D/4A/PTO/HP, WC State,
  ST, dates, time cells, day columns, identity — is now resolved by **label
  discovery in the source**, not by fixed cell positions. No per-family
  branching in field resolution. The YAML profile's `pay_types` and
  `day_start_cols` are fallback hints only. See CLAUDE.md §6.4.
- **Run shape:** 290 employee sheets extracted (260 standard + 11 shifted +
  19 Field Mechanic). 281 PASS / 9 REVIEW / 0 FAIL. 199 dropped by
  empty-employee filter. 41 unmatched (unfilled templates, visible).
- **What's next:** new dataset is about to arrive (manager will drop files
  into `data/` or `data/sample*/`). The whole point of the label-driven
  architecture is to make this run end-to-end without code/YAML changes.
  The verification recipe is CLAUDE.md §13.

---

## 2. What was done this session (chronological)

1. **Empty-employee drop filter committed** (`a90cd9c`) — carry-over from
   prior session.
2. **Inspected Field Mechanic sheets** (Coulson, Pottage, Hall) cell-by-cell.
   Discovered: line-item table, not a day-block grid. Header at row 11
   (label-discovered). Pay-types `RT/OT/DT/PTO/HP` interleaved across 7
   days × 5 cols. Per-line item with JOB-column marker.
3. **Implemented Field Mechanic extract path** (`extract_field_mechanic_sheet`)
   — every field by label. Initial detector matched standard sheets too
   (they also carry `WC STATE` + `RT1/OT1` headers); fixed with the
   structural discriminator: **full day names (MONDAY) = Field Mechanic;
   short (MON) = standard/shifted.**
4. **Audit shows 290/290** pay-type totals + WC/ST match source. Committed
   as `00af7f1`.
5. **Manager's v4 feedback file (`phase2_corpus_v4 - Feedback.xlsx`)
   inspected.** Structure: cols 1–24 our old snapshot, col 25 Comments,
   cols 27–48 manager's expected, cols 50–63 match-grid. Built
   `scripts/compare_v4_feedback.py` to compare CURRENT output to expected
   block, treating expected=None as "not specified" (the manager leaves
   pay-type cells blank on TOTALS rows).
6. **First comparison** showed real mismatches:
   - Field Mechanic line-item DOUBLING (~99 cells, always 2×) —
     real bug. Cause: echo/import rows below real line items also being
     summed. Fixed with `find_fm_job_col` → only sum rows with a job code.
   - Standard sheets' zero-hour cells: manager expected `'0.00'`, we
     emitted `None`. Fixed via `_format_numeric_cells` — renders numeric
     output as 2-decimal strings on day rows.
   - 6 HP/PTO cells on standard sheets where the source's row-12 had
     `M12='PTO1'` instead of `M12='OT1'` (Wirth) or `P12='HP1', Q12='PTO1'`
     (Impastato). These were being routed to OT/4D/4A positionally — the
     same "silent position drift" class as the original v2 WC State bug.
7. **Made the standard path fully label-driven for pay-types**
   (commit `a6bae9a`):
   - Extended `PAYTYPE_DAY_RE` to include `PD|4D|4A` tokens (was missing).
   - New `anchors.discover_paytype_columns_by_label()`: for each day-block,
     reads row-12's labels and maps `(day_idx, pay_type) -> column`.
     Unknown tokens flag `[CHECK]`.
   - `extract_sheet` now reads each output pay-type from its label-
     discovered column, falling back to YAML positions only if no labels
     found. TOTALS row sums the day values (authoritative); source
     positional grand-totals are used as a cross-check that flags
     discrepancies (this surfaces the manager's hand-correction
     workarounds, which is informational).
8. **Final feedback comparison: 15/16 fields at 100%.** Renamed
   authoritative output to `corpus_v5.xlsx` (commits `8b67410`, `4c5245e`).

---

## 3. Files changed / created this session

Modified:
- `src/anchors.py` — added `DAY_TOKENS_FULL`, `ALL_DAY_TOKENS`, `_DAY_ALIASES`,
  `PAYTYPE_DAY_RE` (now includes PD/4D/4A), `day_label_row_style`,
  `find_field_mechanic_header_row`, `find_fm_job_col`,
  `discover_field_mechanic_paytype_cols`, `discover_paytype_columns_by_label`,
  `day_alias_index`, `day_index_to_label`, `find_label_in_grid`.
- `src/extract.py` — added `extract_field_mechanic_sheet`, FM dispatch in
  `process_workbook`, label-driven pay-type discovery throughout
  `extract_sheet`, `_format_numeric_cells` writer formatter.
- `scripts/audit_v4.py` — `to_num` for string-numeric output, FM branch
  (label-driven verification), FM WC/ST cross-check.
- `CLAUDE.md` — reframed §5/§6/§8/§9, added §6.4 (the architectural rule)
  and §13 (new-dataset verification recipe).

Created:
- `scripts/inspect_fm_full.py` — full raw-grid dumper for any sheet (used
  for FM discovery; reusable for any new layout)
- `scripts/probe_fm_paytypes.py` — read-only label-discovery probe
- `scripts/compare_v4_feedback.py` — cell-by-cell vs manager's expected

Memory written (in `~/.claude/projects/c--dev-Extract-excel/memory/`):
- `label-driven-extraction-goal.md` — user's directive
- `field-mechanic-structure.md` — the discovered structure
- `MEMORY.md` — index

Renamed:
- `output/corpus_v5_final.xlsx` → `output/corpus_v5.xlsx`

---

## 4. The verification recipe (what 100% looks like for a new dataset)

Full detail in CLAUDE.md §13. Quick form:

```powershell
# 1. Drop new files into data/sample*/
# 2. Extract:
.venv\Scripts\python.exe src\extract.py data\ --out output\corpus_v6.xlsx

# 3. Audit (edit OUT path in audit_v4.py first):
.venv\Scripts\python.exe scripts\audit_v4.py
#    100% looks like: pay-type 290/290, WC/ST 290/290, suspicious=0

# 4. Compare cell-by-cell vs manager's feedback (edit CURRENT + FEEDBACK in script):
.venv\Scripts\python.exe scripts\compare_v4_feedback.py
#    100% looks like: every field 100% (or 99+% with only rounding
#    discrepancies like 19.98 vs 20.00)
```

The cell-by-cell comparator is the **true accuracy metric**. The audit's
positional check is informational — it can show 7 false-positive mismatches
on sheets where the source itself has positional inconsistencies (manager
workarounds where holiday hours were hand-entered into a wrong grand-total
cell). The label-driven output is correct in those cases; the comparator
confirms it against the manager's truth.

---

## 5. Open decisions / handoff items

1. **No new feedback file yet for `corpus_v5.xlsx`.** The current 15/16 ·100%
   number is against the prior `phase2_corpus_v4 - Feedback.xlsx`. When the
   manager annotates `corpus_v5.xlsx` (or the next run on the new dataset),
   the comparator needs `EXPECTED_COLS` re-verified against the new
   feedback file's column layout (run the column-header probe at the top
   of `compare_v4_feedback.py`).
2. **`scripts/audit_v4.py` `OUT` path is hardcoded** to
   `output/phase2_corpus_v4.xlsx`. Either edit it per-run or expose a CLI
   arg. Same for `compare_v4_feedback.py`'s `CURRENT` and `FEEDBACK`.
3. **`output/phase2_corpus_v4.xlsx` is stale** — the prior authoritative
   file. Not deleted in case the manager's review notes reference it.
4. **Pay-type semantics unresolved** (CLAUDE.md §7): RT+OT+DT vs RT+DT for
   Total Hours varies by company. We don't validate Total Hours = sum of
   pay-types; that check would fail false positives on this dataset.
5. **`scripts/inspect_fm_full.py`** is named FM-specific but is a general-
   purpose raw-grid dumper. Useful first step when any new layout shows up.
6. **YAML drift toward irrelevance** — `pay_types`, `day_start_cols`,
   `grand_totals.cols` are now fallback hints only. Could be removed if
   label-driven path holds on the new dataset.

---

## 6. Environment / run commands

- Working dir: `C:\dev\Extract-excel` (NOT OneDrive — corrupts git packs)
- Python: `.venv\Scripts\python.exe`
- Data: `data\sample\` + `data\sample1\` (gitignored; copy from shared drive
  on a new machine)
- Authoritative output: `output\corpus_v5.xlsx`
- Manager feedback: `phase2_corpus_v4 - Feedback.xlsx` (gitignored — see
  `.gitignore` `*Feedback.xlsx` pattern)
- Git: branch `main`, remote `https://github.com/Zubin123/Extract-excel.git`
- HEAD at handoff: `4c5245e` "Fix typo: corpus_vs.xlsx -> corpus_v5.xlsx"
  (everything pushed)

---

## 7. Resume point for the next session

**Most likely scenario:** the user has dropped new workbooks into `data/`
and wants to verify the extractor handles them at the same accuracy.

Do this:
1. Read CLAUDE.md §5, §6.4, §8, §13.
2. Run the three-step verification from §4 above against the new data.
3. If the cell-by-cell comparator hits 100% (or 99+% with only rounding),
   the architecture held — ship the new output.
4. If not, the four mismatch root-cause buckets in CLAUDE.md §13 step 4
   are exhaustive: rounding, missing pay-type token, new identity layout,
   or new structural family. Diagnose, fix the smallest thing possible
   (vocabulary / regex / classifier branch), re-verify.

**Less likely but possible:** manager comes back with annotations on
`corpus_v5.xlsx` itself. Then `compare_v4_feedback.py` is the tool — point
it at the new feedback file, re-verify column layout, run, address any
mismatches the same way as for a new dataset.
