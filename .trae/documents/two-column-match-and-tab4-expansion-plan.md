# Plan: Two-Column Match + Tab4 Expansion + Compact GUI

## Summary
- Goal: add an optional second match column (example: date) on top of name matching, expose it in GUI, and extend tab 4 with selectable extra columns from both Excel sides.
- Success criteria:
  - User can configure a second match column for Excel 1 and Excel 2.
  - Second-column comparison is configurable by character length (`N`) and contributes to a weighted auxiliary score.
  - Name matching remains primary and **never blocked** by second-column differences.
  - Final status logic remains name-driven; second-column weighted score is used for ranking/explanation/visibility, not as a hard blocker.
  - Tab 4 includes user-selected extra columns from both sides (default includes the second match column).
  - GUI remains functional and gets adaptive compact spacing in configuration areas.

## Current State Analysis
- Current matching pipeline reads only one match column per side:
  - Validation and preview rely on `name_col_t1` / `name_col_t2` in [validate_config](file:///c:/MyProjects/Matcher/matching_nomes_gui_v2.py#L518-L596) and [collect_workbook_preview](file:///c:/MyProjects/Matcher/matching_nomes_gui_v2.py#L470-L516).
  - Input preparation normalizes only name columns in [prepare_input_frames](file:///c:/MyProjects/Matcher/matching_nomes_gui_v2.py#L598-L642).
  - Scoring is name-only in [score_candidate](file:///c:/MyProjects/Matcher/matching_nomes_gui_v2.py#L221-L302).
- Tab 4 currently has exactly 2 visible columns (`Excel 1`, `Excel 2`) in [build_grouped_reconciliation_df](file:///c:/MyProjects/Matcher/matching_nomes_gui_v2.py#L1525-L1574).
- Config GUI currently exposes one match column per side and no selector for tab-4 extra columns in [_build_config_tab](file:///c:/MyProjects/Matcher/matching_nomes_gui_v2.py#L2068-L2207).

## Assumptions & Decisions
- Locked decisions from user:
  - Second-column rule: compare by configurable character length (`N`) and compute auxiliary similarity metadata.
  - Priority: name matching remains decisive; second column never blocks a match directly.
  - Weighted behavior: second-column signal is incorporated as an auxiliary weighted score used for ranking/explanation (not hard gating status).
  - Tab 4 extra columns source: both sides (Excel 1 and Excel 2).
  - GUI compaction: adaptive compact (focus compaction on config sections).
- Output compatibility:
  - Keep existing 4-tab export structure.
  - Preserve existing color buckets and quantity reconciliation logic.

## Proposed Changes

### 1. Extend config schema for second match column and tab4 extras
- File: `matching_nomes_gui_v2.py`
- Add new config variables:
  - `match2_col_t1`, `match2_col_t2` (optional Excel column letters)
  - `match2_prefix_chars` (int, minimum 1)
  - `match2_weight` (float, default low-medium)
  - `tab4_extra_cols_t1`, `tab4_extra_cols_t2` (serialized selection list)
- Validation updates in `validate_config()`:
  - Accept empty second-column config (feature optional).
  - If one side is filled, require the other side.
  - Validate `match2_prefix_chars` and `match2_weight`.
  - Validate selected extra columns exist after workbook header read.

### 2. Update workbook preview for second column visibility
- File: `matching_nomes_gui_v2.py`
- In `collect_workbook_preview()`:
  - Show configured second match column mapping for both files.
  - Show sample values for second column when configured.
  - Include a compact line indicating prefix-char rule and auxiliary weight.

### 3. Extend input preparation with second-column normalization
- File: `matching_nomes_gui_v2.py`
- In `prepare_input_frames()`:
  - Load and normalize optional second match columns into:
    - `match2_t1_original`, `match2_t1_norm`, `match2_t1_prefix`
    - `match2_t2_original`, `match2_t2_norm`, `match2_t2_prefix`
  - Keep existing name normalization unchanged.
  - Ensure missing/blank second-column values are handled safely (empty string).

### 4. Add auxiliary weighted second-column signal to candidate scoring
- File: `matching_nomes_gui_v2.py`
- In scoring flow (`score_candidate()` + caller in `analyze_matching()`):
  - Compute `score_match2_prefix` based on first `N` chars for configured second columns.
  - Add auxiliary metric fields to candidate records:
    - `match2_equal_prefix` (bool)
    - `match2_score` (0-100)
    - `score_composite` (name score + weighted match2 contribution)
  - Use `score_composite` for tie-breaking/ranking order where appropriate.
- Constraint:
  - Do not force SEM_MATCH/REVISAR solely because of second-column mismatch.

### 5. Add explainability fields into analysis/final outputs
- File: `matching_nomes_gui_v2.py`
- Add columns to initialized result defaults:
  - `analysis_match2_t1`, `analysis_match2_t2`, `analysis_match2_equal_prefix`, `analysis_match2_score`
  - `final_match2_t1`, `final_match2_t2`, `final_match2_equal_prefix`, `final_match2_score`
- In analysis assignment:
  - Persist chosen candidate’s second-column comparison evidence.
- In `recompute_final_state()`:
  - Keep status determination name-driven.
  - Preserve second-column evidence fields for export and review.

### 6. Expand tab 4 generator with user-selected extra columns
- File: `matching_nomes_gui_v2.py`
- In `build_grouped_reconciliation_df()`:
  - Keep core columns `Excel 1` and `Excel 2`.
  - Dynamically append selected extra columns from both sides:
    - Prefix naming convention: `E1:<column>` and `E2:<column>`.
  - Default behavior:
    - Include second match column as first extra column if configured.
  - Maintain:
    - alphabetical grouping,
    - blank separator rows between groups,
    - existing bucket metadata for painting.

### 7. Update export for tab4 expanded schema
- File: `matching_nomes_gui_v2.py`
- In `export_analysis_result()`:
  - Pass selected extra-column context to tab4 builder.
  - Export tab 4 with dynamic column set.
- In `format_output_workbook()`:
  - Ensure dynamic columns in tab4 are autosized/formatted consistently.
  - Keep existing bucket painting semantics unchanged.

### 8. GUI additions for second-column matching and tab4 selection
- File: `matching_nomes_gui_v2.py`
- In `_build_config_tab()`:
  - Add optional second match column fields:
    - `Coluna match 2 Excel 1`
    - `Coluna match 2 Excel 2`
    - `Caracteres para match 2`
    - `Peso match 2`
  - Add two multi-select style controls (or compact list+entry strategy) for tab4 extras:
    - extras from Excel 1
    - extras from Excel 2
  - Add helper text: name match remains primary; match2 is auxiliary.
- In sheet/column refresh logic:
  - Reuse detected headers to populate selectable extra columns.

### 9. Adaptive compact UI pass
- File: `matching_nomes_gui_v2.py`
- Apply compact spacing mainly in config tab:
  - reduce paddings and control widths where safe,
  - keep readable labels/tooltips,
  - avoid compacting review/analyze/export heavy tables.
- Keep behavior/functionality unchanged.

### 10. Persistence and backward compatibility
- File: `matching_nomes_gui_v2.py`
- In `save_ui_state()` / `load_ui_state()`:
  - persist new config vars.
  - keep compatibility with older saved state files missing new keys.

## Verification Steps
- Static:
  - Run `py_compile` for `matching_nomes_gui_v2.py`.
  - Verify no diagnostics errors.
- Functional GUI:
  - Open app and confirm new controls render and compact layout remains usable.
  - Load both files; ensure sheet/column options refresh correctly.
- Matching behavior:
  - Case A: same name + same date prefix => strong auxiliary score.
  - Case B: same name + different date prefix => still matchable by name, but auxiliary evidence marks difference.
  - Case C: second-column disabled => behavior matches previous baseline.
- Export:
  - Confirm 4 tabs still generated.
  - Confirm tab4 includes dynamic extra columns from both sides.
  - Confirm default extra includes second match column when configured.
  - Confirm separator blank rows between groups remain.
- Regression:
  - Validate existing quantity mismatch behavior (`EXCESS_LEFT`) still works.
  - Validate manual review/edit/export flow remains stable.

## Risks & Mitigations
- Risk: ambiguity between “weighted score” and “never blocks”.
  - Mitigation: enforce explicit precedence in code and GUI help text.
- Risk: dynamic tab4 columns can break painting/index assumptions.
  - Mitigation: keep bucket columns internal (`_bucket_*`) and paint by keys, not fixed indexes.
- Risk: compact GUI may reduce readability.
  - Mitigation: adaptive compaction limited to config sections only.
