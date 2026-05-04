# Conciliação Quantidades: Date Display and Alignment Fix Plan

## Summary
Fix three issues in `conciliacao_quantidades`:
1. Date values from Excel 1 extras are being shown as `yyyy-mm-dd HH:MM:SS` instead of the required `dd/mm/yyyy`.
2. Column 1 rows that are matched/reviewed are often not bringing the corresponding names to Column 2.
3. Pairing/alignment across match types (`ACEITO`, `REVISAR`, `SEM_MATCH`, excedentes) is unstable and visually misaligned.

The plan focuses on correcting reconciliation key consistency, date presentation formatting, and deterministic pairing behavior without changing unrelated UI or export sheets.

## Current State Analysis

### 1) Date display regression in `conciliacao_quantidades`
- File: `matching_nomes_gui_v2.py`
- In `build_grouped_reconciliation_df()`, extra values for `E1:*` and `E2:*` are written with `str(raw_value or "")`.
- Since source rows come from DataFrame-backed records, date-like values can appear as timestamp text (`yyyy-mm-dd HH:MM:SS`) in the final sheet.
- There is no dedicated formatter for non-monetary date-like extras in reconciliation output.

### 2) Confirmed key mismatch likely causing missing Column 2 pairing
- File: `matching_nomes_gui_v2.py`
- `build_target_catalog()` groups by `nome_t2_match_norm` (truncated by `max_external_chars`) and then renames that key to `nome_t2_norm`.
- Later, `final_group_t2_norm` is populated from chosen candidate `nome_t2_norm` (which is currently that truncated key).
- But in `build_grouped_reconciliation_df()`, right-side records are grouped by `target_df["nome_t2_norm"]` (full normalized name, not truncated).
- Result: left group key and right group key can differ for long names, so matched/review rows in Column 1 may not find corresponding Column 2 rows.

### 3) Misalignment across match types
- File: `matching_nomes_gui_v2.py`
- Pairing uses `_pair_group_rows_by_match2()` with prefix/date heuristics, but group composition can already be inconsistent due to key mismatch.
- `_bucket_right` is currently inherited from left bucket when both sides exist, which can visually blur right-side state when rows are forced/aligned.
- Current render behavior may appear unaligned for extra matches/revisions because pairing quality depends on group integrity first.

## Proposed Changes

### A. Unify reconciliation grouping keys to full normalized target key
- File: `matching_nomes_gui_v2.py`
- What:
  - Preserve both keys in target catalog:
    - stable matching key (truncated/match key)
    - full normalized key (for reconciliation grouping and display integrity)
- How:
  - In `build_target_catalog()`, keep grouped key as explicit `nome_t2_match_key` (or equivalent), and keep full key as `nome_t2_norm_full`.
  - In candidate generation and global assignment, continue using the matching key for performance/quota behavior.
  - In final-state fields (`analysis_*`, `final_group_t2_norm`), store and propagate the full normalized key for reconciliation grouping.
  - In `build_grouped_reconciliation_df()`, group left and right by the same full normalized key.
- Why:
  - This directly addresses the missing corresponding Column 2 names when Column 1 has match/review rows.

### B. Normalize reconciliation date display to `dd/mm/yyyy`
- File: `matching_nomes_gui_v2.py`
- What:
  - Add a reconciliation display formatter for extra columns that:
    - keeps monetary handling unchanged for configured monetary columns
    - formats date-like values as `dd/mm/yyyy`
    - leaves non-date text unchanged
- How:
  - Introduce helper (e.g., `format_reconciliation_extra_value(raw_value, is_money)`):
    - if money column: existing numeric parse behavior remains
    - else: try parse date with `pd.to_datetime(..., dayfirst=True, format="mixed")`
    - if parse succeeds: output `strftime("%d/%m/%Y")`
    - if parse fails: return original string value
  - Apply helper for both `E1:*` and `E2:*` extra columns in `build_grouped_reconciliation_df()`.
- Why:
  - Enforces consistent, human-auditable date format in the final reconciliation sheet.

### C. Stabilize pairing/alignment behavior for all match types
- File: `matching_nomes_gui_v2.py`
- What:
  - Improve deterministic row alignment after key fix so `ACEITO/REVISAR/SEM_MATCH` and excess rows render coherently.
- How:
  - Keep current source-order grouping, but refine pairing strategy:
    - prioritize exact match2 date equality when both sides have parseable dates
    - then exact/near match2 prefix rule
    - then stable row-id fallback
  - Ensure right-only leftovers are appended deterministically and not interleaved unpredictably.
  - Revisit `_bucket_right` assignment to avoid always inheriting left bucket in ambiguous cases; use right-side-informed bucket when available.
- Why:
  - Improves perceived and actual alignment for extra matches/revisions and no-match scenarios.

### D. Guard against regressions in existing export behavior
- File: `matching_nomes_gui_v2.py`
- What:
  - Keep previous export constraints intact:
    - no `resultados_match` sheet
    - original sheet fidelity behavior
    - existing performance improvements
- Why:
  - Avoid reintroducing recently fixed issues while correcting reconciliation.

## Assumptions & Decisions
- Scope is limited to reconciliation correctness and display formatting in `conciliacao_quantidades`.
- Match2 normalization introduced earlier (`dd/mm/yyyy`) remains valid and should be reused for pairing consistency.
- Date display requirement in reconciliation extras is explicit: output must be `dd/mm/yyyy` for date-like values.
- No new UI controls are required for this task.

## Verification Steps
1. Static validation:
  - Run `py_compile` for `matching_nomes_gui_v2.py`.
  - Run diagnostics and resolve introduced issues.
2. Real-file reconciliation validation:
  - Use:
    - `tests/CAR_01_A_29_ABRIL_BPA.xlsx`
    - `tests/OCI_UNIFICADOS_ABRIL.xlsx`
  - Generate export and inspect `conciliacao_quantidades`.
3. Date display checks:
  - Confirm `E1:*` / `E2:*` date-like extras appear as `dd/mm/yyyy`.
  - Confirm non-date text remains unchanged.
4. Alignment checks:
  - Sample matched/reviewed names in Column 1 and verify corresponding Column 2 rows appear in the same paired row when expected.
  - Validate deterministic ordering and pairing for leftovers.
5. Integrity checks:
  - Compare counts:
    - non-empty left rows vs relevant source result rows
    - non-empty right rows vs target rows represented
  - Ensure no drop introduced by key remap.
6. Regression checks:
  - Confirm export still contains expected sheets and respects prior behavior.

## Success Criteria
- `conciliacao_quantidades` shows date-like extras in `dd/mm/yyyy`.
- Matched/partial rows in Column 1 bring corresponding Column 2 names consistently when they belong to the same logical target group.
- Alignment is stable for `ACEITO`, `REVISAR`, `SEM_MATCH`, and excess cases without apparent row drift.
