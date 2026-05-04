# Date Normalization, Export Cleanup, and Matching Performance Plan

## Summary
Implement three coordinated changes in `matching_nomes_gui_v2.py`:
1. Normalize match2 date columns to `dd/mm/yyyy` on both sides (Excel 1 and Excel 2) and compare by exact day.
2. Remove `resultados_match` from export output entirely.
3. Apply targeted, low-risk matching performance optimizations (balanced mode): safer candidate pruning, cache repeated comparisons, and reduce per-row fuzzy scoring volume.

## Current State Analysis
- Date handling today:
  - `prepare_input_frames()` reads sheets with `dtype=str`, then stores match2 raw values in `match2_t1_original` / `match2_t2_original`.
  - `match2_t1_norm` / `match2_t2_norm` are currently just uppercase string normalization, not date normalization.
  - `_parse_optional_datetime()` and pairing logic parse dates later during reconciliation pairing, not consistently at candidate scoring stage.
- Matching/scoring today:
  - For each source row, `choose_candidate_pool()` can fallback to all targets if prefix/first/last token lookup misses.
  - `score_candidate()` recalculates full fuzzy metrics every time for every pair.
  - Main scoring loop in `analyze_matching()` computes all metrics for every candidate in pool, then sorts and keeps top-N.
- Export today:
  - `export_analysis_result()` still writes `resultados_match` (summary + detailed results) and then formats workbook.
  - Output currently has `excel_1_original`, `excel_2_original`, `resultados_match`, `conciliacao_quantidades`.

## Proposed Changes

### 1. Normalize match2 dates to `dd/mm/yyyy` before matching
- File: `matching_nomes_gui_v2.py`
- What:
  - Add a dedicated date normalizer for match2 fields, e.g. `normalize_date_ddmmyyyy(value: Any) -> str`.
  - Normalize both sides (`match2_t1_norm`, `match2_t2_norm`) using this function in `prepare_input_frames()`.
- How:
  - Parse with `pd.to_datetime(..., errors="coerce", dayfirst=True, format="mixed")`.
  - On success, emit canonical `dd/mm/yyyy`.
  - On failure, keep existing fallback normalized text (uppercase/trim) to avoid data loss.
  - Ensure candidate logic uses normalized match2 values for exact-day comparison.
- Decision:
  - Matching rule is exact day equality on normalized `dd/mm/yyyy` when dates are parseable on both sides.

### 2. Enforce exact-day date comparison in candidate scoring
- File: `matching_nomes_gui_v2.py`
- What:
  - Update match2 scoring behavior in `score_candidate()` for date-normalized values.
- How:
  - If both normalized match2 values are valid `dd/mm/yyyy`, compute match2 equality as exact string equality.
  - Keep current non-date fallback behavior for non-parseable values.
  - Keep `match2_weight` pipeline intact so only match2 sub-score semantics change for date values.
- Why:
  - Requirement is to match with equal date format after normalization.

### 3. Remove `resultados_match` sheet from export
- File: `matching_nomes_gui_v2.py`
- What:
  - In `export_analysis_result()`, stop writing both summary and detailed rows to `resultados_match`.
  - Keep only:
    - `excel_1_original`
    - `excel_2_original`
    - `conciliacao_quantidades`
- How:
  - Remove `summary_rows_df.to_excel(... resultados_match ...)`.
  - Remove `results_export_df.to_excel(... resultados_match ...)`.
  - Update `results_startrow` usage and formatting function assumptions so export formatting does not expect `resultados_match`.
  - Update `format_output_workbook()` to skip any branch specific to `resultados_match`.
- Decision:
  - No replacement summary sheet; no summary block in final sheet.

### 4. Performance optimization #1: reduce obvious non-match candidate generation
- File: `matching_nomes_gui_v2.py`
- What:
  - Add conservative pre-filtering before full `score_candidate()` execution.
- How:
  - Keep existing pool builder, but apply fast checks first:
    - require same first token OR high prefix overlap OR token intersection above minimal threshold.
    - for date-enabled match2, when both dates parse and are different days, deprioritize/drop from scoring (balanced safe rule).
  - Only run expensive fuzzy metrics on filtered candidates.
- Why:
  - Cuts useless candidate evaluations with low behavior risk.

### 5. Performance optimization #2: cache repeated normalization/comparison work
- File: `matching_nomes_gui_v2.py`
- What:
  - Memoize repeated string-comparison calculations.
- How:
  - Add local dictionaries keyed by `(left_name, right_name)` for:
    - `fuzz.token_set_ratio`
    - `fuzz.partial_ratio`
    - `fuzz.token_sort_ratio`
    - prefix ratio and ordered/aligned ratios where applicable
  - Cache token sets / first / last token for repeated names if needed.
- Why:
  - Many source rows compare against repeated target names; caching avoids duplicate expensive calls.

### 6. Performance optimization #3: reduce per-row fuzzy scoring volume
- File: `matching_nomes_gui_v2.py`
- What:
  - Introduce staged scoring in `analyze_matching()`:
    - Stage A fast score/filters
    - Stage B full score only for best subset
- How:
  - Run quick gate (cheap metrics) for all pool candidates.
  - Keep top-K preselected candidates (internal threshold) for full `score_candidate()`.
  - Preserve `top_candidates_to_keep` semantics for final output.
- Why:
  - Reduces total full-fuzzy calls while maintaining balanced behavior.

### 7. Keep export/original-sheet guarantees from previous fixes
- File: `matching_nomes_gui_v2.py`
- What:
  - Preserve already-applied behavior:
    - original sheet value fidelity
    - targeted coloring only for selected match-related cells
    - reconciliation ordering by source order
- Why:
  - Avoid regressions while implementing new requirements.

## Assumptions & Decisions
- Confirmed decisions from user:
  - Date normalization/comparison: exact day (`dd/mm/yyyy`) on both sides.
  - `resultados_match`: remove completely; no replacement summary sheet.
  - Performance strategy: balanced safe gains (not maximum-risk pruning).
- Compatibility:
  - Keep existing GUI config fields and workflows.
  - Keep matching output semantics (`ACEITO/REVISAR/SEM_MATCH`) unless changes are direct consequence of exact-day rule.
- Out of scope:
  - No new UI controls added for this task.
  - No new dependencies.

## Verification Steps
1. Static checks:
  - Run `py_compile` on `matching_nomes_gui_v2.py`.
  - Run diagnostics and fix any introduced issues.
2. Functional date normalization checks:
  - Feed date-like values in match2 (`T`/`H`) with mixed source formats.
  - Confirm normalized internal values are `dd/mm/yyyy`.
  - Confirm exact-day matches behave as expected.
3. Export structure checks:
  - Run export and confirm output workbook contains only:
    - `excel_1_original`
    - `excel_2_original`
    - `conciliacao_quantidades`
  - Confirm `resultados_match` is absent.
4. Performance checks (same real dataset):
  - `tests/CAR_01_A_29_ABRIL_BPA.xlsx`
  - `tests/OCI_UNIFICADOS_ABRIL.xlsx`
  - Capture progress timestamps to compare:
    - candidate scoring phase duration
    - export phase duration
5. Regression checks:
  - Confirm reconciliation row counts remain coherent (left/right non-empty counts).
  - Confirm previous export fidelity fixes remain valid for source values and targeted coloring.

## Success Criteria
- Match2 dates are normalized to `dd/mm/yyyy` and compared by exact day when parseable.
- Export no longer generates `resultados_match`.
- Matching runtime improves materially through reduced candidate/scoring workload with balanced risk.
- No regression in final export integrity and reconciliation organization.
