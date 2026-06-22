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

## Reference data

The workbook's lookup sheets (`Uni_Tiers`, `Dept_Fit`, `Status_Map`,
`Teknofest_Map`, `Weights`) are the source of truth for valid values.
The parser's `UNIVERSITIES` / `DEPARTMENTS` dicts mirror those keys; if
you add a row to a lookup sheet, also add a synonym in the parser so
new CVs map to it.
