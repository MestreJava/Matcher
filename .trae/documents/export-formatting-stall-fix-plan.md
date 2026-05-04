# Export Formatting Stall Fix Plan

## Summary
The export appears to stall at `Formatando arquivo...` because the workbook formatting phase in `matching_nomes_gui_v2.py` is doing an expensive full-workbook OpenPyXL pass after Pandas has already written all sheets. The most likely hotspot is the combination of per-row fill loops across entire worksheets plus `_autosize_columns()` scanning every column up to 3000 rows on every sheet. The fix should keep the same output structure, but make formatting bounded, targeted, and predictable so export completes reliably.

## Current State Analysis
- Export flow is:
  - `export_analysis_result()` writes sheets with Pandas/OpenPyXL.
  - Then it emits `Formatando arquivo...` and calls `format_output_workbook()`.
- Relevant code paths:
  - `export_analysis_result()` at `matching_nomes_gui_v2.py`
  - `format_output_workbook()` at `matching_nomes_gui_v2.py`
  - `_autosize_columns()` at `matching_nomes_gui_v2.py`
  - `build_grouped_reconciliation_df()` at `matching_nomes_gui_v2.py`
- The formatting stage currently does all of the following in one synchronous background export step:
  - `load_workbook(output_file)`
  - applies fills row-by-row for `resultados_match`
  - applies fills row-by-row for `excel_1_original`
  - applies fills row-by-row for `excel_2_original`
  - applies fills row-by-row for `conciliacao_quantidades`
  - applies per-column monetary format lookup on reconciliation sheet
  - applies `ws.auto_filter.ref = ws.dimensions` for every sheet
  - calls `_autosize_columns(ws)` for every sheet
  - saves workbook again with `wb.save(output_file)`
- `_autosize_columns()` is especially expensive because it iterates all worksheet columns and scans up to 3000 cells per column, converting each cell value to string.
- Since the user reports the progress stops specifically on `Formatando arquivo...`, the failure mode is most likely performance stall during workbook post-processing rather than data matching or DataFrame export.
- Recent changes added monetary formatting support on the reconciliation sheet, but that logic is narrow and header-targeted; it is unlikely to be the primary bottleneck compared with full-sheet autosizing and fill loops.

## Proposed Changes

### 1. Instrument and isolate the formatting hotspots
- File: `matching_nomes_gui_v2.py`
- What:
  - Add progress/log checkpoints inside `format_output_workbook()` for each major sheet/phase.
  - Split formatting into smaller helper functions so time sinks are visible.
- Why:
  - This confirms whether the slowdown is in autosizing, fill application, workbook save, or reconciliation formatting.
- How:
  - Extract helper boundaries such as:
    - `format_results_sheet(...)`
    - `format_original_sheet(...)`
    - `format_reconciliation_sheet(...)`
    - `apply_sheet_autofit(...)`
  - Emit progress messages before each section.

### 2. Bound or reduce autosize work
- File: `matching_nomes_gui_v2.py`
- What:
  - Replace the current blanket `_autosize_columns()` strategy with a bounded strategy.
- Why:
  - This is the strongest suspected cause of the stall.
- How:
  - Reduce scan depth from the current broad pass to a smaller sample window.
  - Skip autosize for very wide/very large sheets, or only autosize selected sheets/selected columns.
  - Prefer fixed widths for known heavy sheets where practical.
- Decision:
  - Keep autosize behavior where useful, but make it explicitly size-bounded so export time scales predictably.

### 3. Reduce per-cell formatting work on large sheets
- File: `matching_nomes_gui_v2.py`
- What:
  - Avoid unnecessary whole-row/per-cell fill loops where the value is not needed for user-visible output.
- Why:
  - OpenPyXL cell-by-cell formatting across large sheets is slow.
- How:
  - Limit color application to actual data region only.
  - Reuse resolved fill objects.
  - Avoid repeated dictionary conversions where possible.
  - Review whether `excel_1_original` and `excel_2_original` need full-row fill on every exported row or whether formatting can be narrowed.

### 4. Keep reconciliation-sheet monetary formatting targeted
- File: `matching_nomes_gui_v2.py`
- What:
  - Preserve the recent `R$` formatting feature, but ensure it remains header-targeted and does not trigger broad scanning.
- Why:
  - The fix should not regress the requested monetary export behavior.
- How:
  - Keep lookup restricted to the configured `E1:<col>` / `E2:<col>` headers only.
  - Continue applying currency format only when the cell value is numeric.

### 5. Protect the UI/background export flow
- File: `matching_nomes_gui_v2.py`
- What:
  - Improve background export observability and failure reporting.
- Why:
  - Even if formatting is slower than expected, the user should know which step is running instead of seeing a generic stuck state.
- How:
  - Emit more granular progress labels during formatting.
  - Ensure any exception in formatting/save reaches the existing background error handler with a clear stage label.

## Assumptions & Decisions
- The root cause is treated as a performance bottleneck in workbook post-formatting, not a matching-engine bug.
- The export structure and current sheets must remain unchanged:
  - `excel_1_original`
  - `excel_2_original`
  - `resultados_match`
  - `conciliacao_quantidades`
- The recent monetary-column behavior for `conciliacao_quantidades` stays in scope and must keep working.
- The fix should prioritize completion reliability over perfect column autofit precision.
- No change is planned to matching logic unless debugging proves the stall happens before formatting.

## Verification Steps
- Run a normal export with the sample project workbook and confirm the process moves past `Formatando arquivo...`.
- Add/export with a heavier workbook scenario if available, or simulate with the existing sample plus current features enabled.
- Verify the output file opens successfully and contains all expected sheets.
- Verify formatting still works for:
  - header styling
  - conflict highlighting
  - status bucket fills
  - monetary `R$` formatting on reconciliation sheet
- Confirm compile and diagnostics remain clean after changes.
- Confirm export progress messages clearly indicate the active formatting sub-step if formatting takes noticeable time.
