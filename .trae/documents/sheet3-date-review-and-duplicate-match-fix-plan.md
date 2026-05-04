# Sheet3 Date Review and Duplicate Match Fix Plan

## Summary
Fix three connected problems in `matching_nomes_gui_v2.py`:
1. `conciliacao_quantidades` sometimes flips ambiguous dates such as `06/04/2026` into `04/06/2026`.
2. Exact/review matches in Sheet3 are missing the expected corresponding name on Column B for cases like `LUCINEIA PEREIRA DA SILVA`.
3. Repeated-name matching is incorrect for cases like `ROSANE BATISTI`, because the current matching stage effectively scores against only the first target-row date for a duplicated target name instead of evaluating the repeated target rows correctly.

## Current State Analysis

### 1. Ambiguous date parsing in Sheet3
- File: `matching_nomes_gui_v2.py`
- Current helpers:
  - `normalize_date_ddmmyyyy()`
  - `_parse_optional_datetime()`
  - `format_reconciliation_extra_value()`
- All three rely on `pd.to_datetime(..., dayfirst=True, format="mixed")`.
- Risk:
  - For ambiguous values like `06/04/2026`, mixed parser behavior can reinterpret the date instead of preserving the intended `dd/mm/yyyy`.
  - This affects both matching and Sheet3 display formatting.

### 2. Missing Column B pair in `conciliacao_quantidades`
- File: `matching_nomes_gui_v2.py`
- Current reconciliation logic:
  - left side grouped by `final_group_t2_norm`
  - right side grouped by match-normalized key
  - rows paired by `_pair_group_rows_by_match2()`
- A partial fix already aligned grouping keys better, but pairing quality still depends on what the matching phase produced.
- If the chosen target candidate is wrong or incomplete, Sheet3 can still show a valid left-side row with no corresponding right-side exact/review pair.

### 3. Confirmed root cause for repeated-name mismatch
- File: `matching_nomes_gui_v2.py`
- In `build_target_catalog()`:
  - target rows are grouped by `nome_t2_match_norm`
  - `match2_t2_original` and `match2_t2_norm` are aggregated using `"first"`
- Consequence:
  - For repeated names in Excel 2, only the first row’s date survives candidate scoring.
  - Example pattern: if a name appears 13 times in Excel 1 and 2 times in Excel 2, scoring does not truly compare each source row against the two distinct target-row date variants.
  - This explains cases where exact date matches exist but rows are not classified as exact.

## Proposed Changes

### A. Replace ambiguous date parsing with deterministic review-and-correction logic
- File: `matching_nomes_gui_v2.py`
- What:
  - Introduce a strict date parser/normalizer for values intended to be `dd/mm/yyyy`.
- How:
  - Add helper(s) such as:
    - `parse_date_prefer_ddmmyyyy(value)`
    - `normalize_date_ddmmyyyy_safe(value)`
  - Parsing strategy:
    - if already matches `dd/mm/yyyy`, parse strictly with that format and keep the same day/month order
    - if matches ISO-like patterns (`yyyy-mm-dd`, `yyyy-mm-dd HH:MM:SS`), parse explicitly as ISO then format to `dd/mm/yyyy`
    - only use generic fallback parsing as last resort
  - Add an auto-review/correction path:
    - when a value is ambiguous and cannot be trusted, preserve the original text for display and avoid silently swapping day/month
- Why:
  - Prevents `06/04/2026` from turning into `04/06/2026`.

### B. Use the same strict date logic everywhere dates matter
- File: `matching_nomes_gui_v2.py`
- What:
  - Apply the deterministic parser consistently in:
    - `normalize_date_ddmmyyyy()`
    - `_parse_optional_datetime()`
    - `format_reconciliation_extra_value()`
- Why:
  - Matching and reconciliation display must use the same date interpretation rules.

### C. Stop collapsing duplicate target rows to one date during candidate scoring
- File: `matching_nomes_gui_v2.py`
- What:
  - Rework `build_target_catalog()` so repeated target names preserve row-level date variants.
- How:
  - Keep a stable quota/match key for repeated names, but do not aggregate `match2_t2_*` using only `"first"`.
  - Introduce row-level candidate records for target rows, or grouped records split by meaningful date variants.
  - Ensure scoring sees distinct target-row dates for repeated names.
- Why:
  - This is the main correctness fix for examples like `ROSANE BATISTI`.

### D. Update assignment and quota handling to work with repeated target rows safely
- File: `matching_nomes_gui_v2.py`
- What:
  - Preserve current name-level quota behavior while allowing row-level candidate evaluation.
- How:
  - Keep quota accounting at the appropriate logical target group level.
  - Carry row-level identifiers (`target_row_id`, date variant) through scoring and assignment.
  - Ensure the final chosen candidate retains the actual target row/date that produced the best valid match.
- Why:
  - Needed so exact date-compatible repeated names can actually win the match.

### E. Improve Sheet3 pairing so exact/review rows bring the correct Column B counterpart
- File: `matching_nomes_gui_v2.py`
- What:
  - Refine `_pair_group_rows_by_match2()` after the repeated-name fix.
- How:
  - Pair priority should be:
    1. strict exact date equality using deterministic parser
    2. exact normalized match2 text equality
    3. stable closest-date fallback
    4. deterministic row-id fallback only when a true target association still exists
  - Avoid pairing a left-side row to an arbitrary right-side row just because it shares the same truncated name group.
- Why:
  - This should correct missing Column B counterparts for true exact/review rows like `LUCINEIA PEREIRA DA SILVA`.

### F. Add automatic review detection for suspicious repeated-name/date conflicts
- File: `matching_nomes_gui_v2.py`
- What:
  - Add a lightweight safety rule for repeated-name groups where available exact-date target rows exist but the selected candidate is inconsistent.
- How:
  - If multiple target rows exist for the same match-normalized name and a strict date-equal candidate exists, prefer it.
  - If ambiguity remains after deterministic ranking, force `REVISAR` instead of accepting a weaker mismatched date candidate.
- Why:
  - Prevents false “not exact” outcomes when strong exact-date evidence exists.

## Assumptions & Decisions
- The intended canonical date format is `dd/mm/yyyy`.
- Silent month/day inversion is unacceptable and should be prevented even if it means preserving original text when confidence is low.
- Matching correctness for repeated names takes priority over preserving the current simplified grouped-target scoring model.
- Existing performance improvements should be preserved as much as possible, but correctness comes first for repeated-name/date matching.

## Verification Steps
1. Static validation:
  - Run `py_compile` on `matching_nomes_gui_v2.py`
  - Run diagnostics and resolve any introduced issues
2. Date safety validation:
  - Feed representative values:
    - `06/04/2026`
    - `27/04/2026`
    - `2026-04-27 00:00:00`
  - Confirm output stays or normalizes to correct `dd/mm/yyyy` without month/day flip
3. Real workbook validation using the files in `tests`:
  - `CAR_01_A_29_ABRIL_BPA.xlsx`
  - `OCI_UNIFICADOS_ABRIL.xlsx`
  - compare against observed issues in `CAR_01_A_29_ABRIL_BPA__OCI_UNIFICADOS_ABRIL_resultado_matching.xlsx`
4. Specific case checks:
  - confirm `LUCINEIA PEREIRA DA SILVA` exact match in Column A brings the expected Sheet2 counterpart in Column B
  - inspect repeated-name case `ROSANE BATISTI` and confirm rows with equal dates are no longer downgraded incorrectly
5. Integrity checks:
  - verify Sheet3 still preserves all relevant left/right rows
  - verify exact/review counts remain coherent after row-level duplicate handling

## Success Criteria
- No ambiguous Sheet3 date is silently flipped from `dd/mm/yyyy` to `mm/dd/yyyy`.
- Exact/review rows in Sheet3 consistently bring the expected Column B counterpart when it exists.
- Repeated-name cases with valid equal dates are no longer scored only against the first target-row date.
- Examples like `ROSANE BATISTI` and `LUCINEIA PEREIRA DA SILVA` behave correctly in the real test workbook.
