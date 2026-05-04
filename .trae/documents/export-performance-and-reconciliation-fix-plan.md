# Export Performance and Reconciliation Fix Plan

## Summary
The export stage is still too slow on larger workbooks because the current formatting phase paints full rows in the two original export sheets and still performs a full OpenPyXL post-processing pass. There are also two likely correctness/UX issues in the current export logic: the last reconciliation sheet is explicitly reordered by a detected date column instead of preserving source order, and the exported "original" sheets are rebuilt from Pandas values instead of preserving the exact displayed source cell text, which can change date presentation and character formatting.

This plan keeps the current workbook structure, but changes export behavior to:
- paint only the match-related cells on original sheets
- preserve source order in the last sheet
- preserve exact displayed source values in exported original sheets without reformatting character/date content
- analyze and harden the reconciliation output so rows are not visually lost or misorganized

## Current State Analysis

### 1. Export slowness
- File: `matching_nomes_gui_v2.py`
- Current flow:
  - `export_analysis_result()` writes the workbook with Pandas.
  - `format_output_workbook()` reloads it with OpenPyXL and formats all sheets.
- Confirmed hotspots:
  - `_format_original_sheet()` paints the entire row for every matched/exported row in `excel_1_original` and `excel_2_original`.
  - `_fill_sheet_row()` loops through every cell in the row.
  - `_apply_sheet_autofit()` is now bounded, but formatting still remains expensive when original sheets are large because row coloring touches many cells.
- Impact:
  - Export time grows with sheet width and number of matched rows.
  - The biggest remaining performance cost is not matching, but export formatting.

### 2. Last-sheet ordering problem
- File: `matching_nomes_gui_v2.py`
- Confirmed behavior:
  - `build_grouped_reconciliation_df()` builds the last sheet from grouped rows.
  - `build_ordered_export_df()` sorts by a detected date column if one is found.
  - The reconciliation path also groups by `final_group_t2_norm` and then sorts group keys alphabetically.
- Risk:
  - The last sheet may look "wrong" to the user because it does not preserve the original Excel 1 order.
  - Group-based ordering and date-based reordering can make rows appear moved or "missing".

### 3. Possible result-loss perception
- File: `matching_nomes_gui_v2.py`
- Likely causes from current logic:
  - `build_grouped_reconciliation_df()` appends a blank separator row after every group.
  - grouping uses `final_group_t2_norm` or fallback `nome_t1_norm`
  - pairing is performed by `_pair_group_rows_by_match2()`
  - left rows are sorted by status/score before pairing, not by source row order
- Why this matters:
  - Even if rows are not actually dropped, the current ordering/pairing can make the final reconciliation sheet look incomplete or misorganized.
  - This is especially risky when there are repeated names, quota conflicts, or mixed accepted/review/no-match rows.

### 4. Original value formatting problem
- File: `matching_nomes_gui_v2.py`
- Confirmed behavior:
  - `prepare_input_frames()` reads Excel using `pd.read_excel(..., dtype=str)`.
  - `export_source_df` and `export_target_df` are exported from these Pandas DataFrames.
- Why dates/character types can appear changed:
  - Pandas reads cell values, not the exact Excel display formatting.
  - The export recreates cells from DataFrame content, so the original Excel display mask/style is not preserved automatically.
  - If a date column looked formatted in the source sheet, the exported "original" sheet can still differ because it is a reconstructed sheet, not a copy of the original worksheet.
- User decision:
  - Do not reformat original source character/date content.
  - Preserve the exact displayed source value in exported original sheets.

## Proposed Changes

### 1. Reduce original-sheet coloring to match-related cells only
- File: `matching_nomes_gui_v2.py`
- What:
  - Replace full-row coloring in `_format_original_sheet()` with targeted cell coloring.
- How:
  - Identify the selected match-driving columns from config:
    - `name_col_t1`, `name_col_t2`
    - optional `match2_col_t1`, `match2_col_t2`
  - Only color those corresponding cells in `excel_1_original` and `excel_2_original`.
  - Keep header styling and freeze panes unchanged.
- Why:
  - This directly addresses the main remaining export bottleneck.
- Decision:
  - Do not color full rows on original sheets.
  - Color only the actual cells relevant to the match.

### 2. Preserve exact displayed source values for original-sheet export
- File: `matching_nomes_gui_v2.py`
- What:
  - Stop exporting `excel_1_original` and `excel_2_original` only from Pandas-reconstructed values when exact display preservation is required.
- How:
  - Add a source workbook extraction path for the original sheets that reads the visible worksheet cell text/layout needed for export.
  - Export original sheets from a display-preserving representation instead of the normalized analysis DataFrames.
  - Keep the analysis pipeline (`prepare_input_frames()`) unchanged for matching logic.
- Why:
  - The user explicitly wants exact original displayed values and no character/date reformatting.
- Decision:
  - Preserve displayed cell value, not only the Pandas-interpreted value.
  - Do not coerce original-sheet date columns to new date/number types during export.

### 3. Keep the last reconciliation sheet in source order
- File: `matching_nomes_gui_v2.py`
- What:
  - Change reconciliation row ordering to preserve Excel 1 source order.
- How:
  - Use `source_row_id` / `excel_row_t1` as the primary order driver for left-side rows.
  - Remove or bypass the current date-based reorder behavior for this export context.
  - Keep group/pair logic, but make final rendered row order stable and source-driven.
- Why:
  - This matches user expectation and avoids apparent "loss" caused by reordering.
- Decision:
  - Source order wins over automatic date sorting.

### 4. Audit and correct reconciliation pairing/count behavior
- File: `matching_nomes_gui_v2.py`
- What:
  - Review the logic in:
    - `build_grouped_reconciliation_df()`
    - `_pair_group_rows_by_match2()`
    - `recompute_final_state()`
  - Focus on cases where accepted/review/no-match rows and quota conflicts may produce confusing or incomplete output.
- How:
  - Verify that every relevant left-side result row and every target row appears exactly once in the rendered reconciliation output for the intended grouping mode.
  - Check whether separator rows, grouping fallback (`nome_t1_norm`), and right-side bucket assignment produce misleading structure.
  - If necessary, adjust grouping and pairing to be stable, explicit, and source-order-safe.
- Why:
  - The user suspects some results are being lost and the last sheet is not organized correctly.
- Decision:
  - Treat this as a correctness review first, not just a visual tweak.

### 5. Keep date sorting out of export ordering unless explicitly requested
- File: `matching_nomes_gui_v2.py`
- What:
  - Remove implicit date-based ordering from export paths that are expected to preserve source layout.
- How:
  - Review `build_ordered_export_df()` and any place that calls `pick_primary_date_column()`.
  - Restrict date-based sorting to a future optional mode only if explicitly configured.
- Why:
  - Silent date-driven reorder is incompatible with the requested source-preserving export behavior.

### 6. Improve export-stage architecture for future speed
- File: `matching_nomes_gui_v2.py`
- What:
  - Analyze and structure a safer/faster export pipeline while keeping current behavior minimal for this fix.
- Improvements to include in the implementation:
  - split data-generation and formatting steps clearly
  - avoid re-walking wide sheets whenever possible
  - avoid converting original display content through Pandas when fidelity matters
  - keep targeted progress messages per stage
- Why:
  - The export stage is now the main operational bottleneck for large files.

## Assumptions & Decisions
- Keep the current workbook sheets:
  - `excel_1_original`
  - `excel_2_original`
  - `resultados_match`
  - `conciliacao_quantidades`
- Original-sheet formatting policy:
  - preserve exact displayed source values
  - do not reformat date/character content
  - do not color full rows
  - color only match-related cells
- Last-sheet ordering policy:
  - preserve source order from Excel 1
  - do not silently reorder by detected date
- Reconciliation correctness:
  - every relevant result row should remain visible
  - apparent losses must be investigated as a logic/output issue, not dismissed as only sorting

## Proposed File Changes

### `matching_nomes_gui_v2.py`
- Refactor `_format_original_sheet()` to paint only selected match cells.
- Add mapping from Excel column letters in config to worksheet column indexes for original-sheet formatting.
- Introduce a display-preserving export path for original worksheets.
- Remove implicit date-based ordering from source-preserving export output.
- Rework `build_grouped_reconciliation_df()` ordering so rendered rows follow source order.
- Review and correct pairing/grouping behavior where results may look lost or misplaced.
- Preserve current targeted `R$` formatting behavior only where explicitly configured for reconciliation extras.

## Verification Steps
- Export the provided real-world case:
  - `tests/CAR_01_A_29_ABRIL_BPA.xlsx`
  - `tests/OCI_UNIFICADOS_ABRIL.xlsx`
- Compare the generated workbook against:
  - `tests/CAR_01_A_29_ABRIL_BPA__OCI_UNIFICADOS_ABRIL_resultado_matching.xlsx`
- Verify:
  - export time is materially reduced
  - original sheets preserve visible source values exactly
  - source date columns are not reformatted
  - only match-related cells are colored on original sheets
  - the last sheet preserves source order
  - no rows appear lost in reconciliation output
  - reconciliation grouping/pairing is stable for repeated names and quota-conflict scenarios
- Run compile and diagnostics checks after implementation.

## Success Criteria
- Export no longer takes extreme time mainly due to original-sheet full-row painting.
- Original sheets keep the source display value exactly, especially for date-like columns.
- The last sheet is easier to audit because it preserves source order.
- Reconciliation output no longer appears to lose rows due to sorting/grouping side effects.
