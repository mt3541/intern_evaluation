# GNC intern CV → workbook extraction pipeline

Three-stage, idempotent. All stages are repeatable; only stage 1 touches PDFs.

```
PDFs → [01] → cv_text_cache/*.txt → [02] → extracted/candidates_auto.json
                                                        ↓
                                      extracted/candidates_overrides.json (hand-edited)
                                                        ↓
                                                      [03] → gnc_intern_evaluation_v1.xlsx
```

## Stage 1 — `01_extract_text.py`
Reads every PDF in `gnc_june_2026/` with `pdfplumber` and caches the plain
text under `cv_text_cache/<stem>.txt`. Cached files are reused on rerun;
pass `--force` to re-extract. Pure mechanical, zero LLM tokens.

Image-only PDFs come out near-empty (e.g. `Simar_Bayramusta_CV…`,
`Samet_Atak_PORTFOLIO`, `Melike_Demirtaş_PORTFOLIO`). For those, the
candidate will need a manual review pass — flagged in the overrides file.

## Stage 2 — `02_parse_candidates.py`
Heuristic parser. For each unique candidate (grouped by the first two
deaccented filename tokens — Turkish chars normalized) merges CV +
portfolio text and emits one JSON record with:

- **Hard fields** — name, university, department, academic_status, gpa,
  teknofest level: regex / keyword / synonym lookup against the workbook's
  `Uni_Tiers`, `Dept_Fit`, `Status_Map`, `Teknofest_Map` reference sheets.
- **Soft 1-5 scores** — real_project, portfolio, coursework, tools, rc_uav,
  english: keyword-count / length heuristics. These are STARTING POINTS;
  reviewers should sanity-check against the CV before finalizing.

Output: `extracted/candidates_auto.json` (regenerated on every run).
Stdout: a `REVIEW <name>` line per candidate whose hard fields couldn't
be confidently extracted.

## Stage 3 — `03_fill_workbook.py`
Merges `candidates_auto.json` ← `candidates_overrides.json` and writes
into the `Scoring` sheet (rows 3-32). Any field set in the overrides file
wins over the auto-extracted value. The first run backs up the original
workbook to `gnc_intern_evaluation_v1.xlsx.bak`.

**Tech-interview columns (O Controls/GNC, P Math/Physics, Q Coding,
R Debugging, S Communication) and HR Feedback (N) are intentionally
left blank** — those must be filled after the interview / HR call.
The workbook's auto-computed weighted score / ranking will only become
meaningful once they're populated.

## Workflow

```bash
# First time:
python3 scripts/01_extract_text.py
python3 scripts/02_parse_candidates.py
# review extracted/candidates_auto.json + REVIEW warnings,
# patch extracted/candidates_overrides.json as needed
python3 scripts/03_fill_workbook.py
```

To add or remove candidates: drop / delete a PDF in `gnc_june_2026/`,
rerun stage 1 then stages 2 and 3. The overrides file is keyed by
`<firstname>_<lastname>` (deaccented lowercase) so existing edits survive
new additions.

## Candidate ID convention

```python
canonical_id = "_".join(deaccent(stem).lower().split("_")[:2])
```

Example: `Hüseyin_Kırca_PORTFOLIO.pdf` → `huseyin_kirca`.

## Stage 5 — `05_build_question_bank.py`

Reads `extracted/question_bank.json` (9 groups × 10 Q+A, bilingual EN/TR)
and (re)writes the **`QuestionBank` sheet** in the workbook. Backs up the
xlsx to `.bak2` on first run. The row order is the contract for the
per-student sheets — do NOT reorder questions without also rerunning
stage 6.

To edit a question or answer: open `extracted/question_bank.json`,
modify the relevant `q_en` / `q_tr` / `a_en` / `a_tr` in place, rerun
stage 5. Every student sheet's reference automatically picks up the new
text (it's a live formula link to the bank).

Group 7 (`company_platforms`) is intentionally placeholders — fill in
proprietary platform questions and rerun.

## Stage 6 — `06_build_student_sheets.py`

Two responsibilities (run after stage 5):

1. **Restructure the Scoring sheet.** Replaces the 5 fixed tech-interview
   columns (`O Controls/GNC` … `S Communication`) with **9 columns** —
   one per question group from the bank. The other columns shift right
   by 4 (T → X, U → Y, V → Z, W → AA, X..AD → AB..AH). All affected
   formulas are rewired:
   - `Y` (Weighted Score) uses new auto-component columns AB..AF + AG.
   - `Z` (Rank) uses sort-stable `ROW()` form, targeting Y.
   - `AG` (Tech Composite) = `AVERAGE(O:W)` skipping blanks.
   - `Ranking` sheet `INDEX/MATCH` formulas retargeted to Scoring!X/Y/Z.

2. **Generate one `Q_<canonical_id>` sheet per candidate** (30 sheets),
   using the same canonical_id from stage 2 so each sheet pairs with the
   Scoring row. Each sheet has 9 group blocks; each block is a banner row
   (with a live group average in col G) and 10 question rows whose
   #/Group/Q-EN/Q-TR/A-EN/A-TR are formula references to the
   QuestionBank. Cols G (Grade 1-5, yellow) and H (Notes) are interviewer
   inputs. Blank Grade = "not asked", excluded from group average.

   Top of each sheet: candidate name + live Tech Composite cell that
   averages the 9 group banners (groups with no graded questions are
   skipped so students asked different topics stay comparable). Row 2 is
   a merged free-form **General Notes** cell.

**Idempotent** — rerunning stage 6 drops and recreates each `Q_<id>`
sheet but PRESERVES whatever you had typed in cols G:H plus the general
notes cell B2. So you can iteratively edit `question_bank.json` + rerun
stage 5 alone (content-only, doesn't touch student sheets), OR rerun
both 5+6 (rebuilds structure; grades survive).

### Updating a candidate's interview score

1. Open the candidate's `Q_<id>` sheet during the interview.
2. Type a 1-5 grade in col G next to each question you actually ask.
   Leave un-asked rows blank — they don't count.
3. Per-question notes go in col H; broader observations in the merged
   `General Notes` cell at the top.
4. The group averages, the Tech Composite, and the Scoring/Ranking
   sheets all recalculate automatically.

### Adding / removing question groups

Edit `extracted/question_bank.json` (add or remove a `groups[]` entry),
update `N_GROUPS` in `06_build_student_sheets.py`, rerun stage 5 then 6.
Also update `update_rubric_sheet()` to add a rubric for the new group.

## Reference data

The workbook's lookup sheets (`Uni_Tiers`, `Dept_Fit`, `Status_Map`,
`Teknofest_Map`, `Weights`) are the source of truth for valid values.
The parser's `UNIVERSITIES` / `DEPARTMENTS` dicts mirror those keys; if
you add a row to a lookup sheet, also add a synonym in the parser so
new CVs map to it.
