# Matcher Safe Cleanup And Refactor Order Spec

## Why
The current app works but has high maintenance risk due to a single large file and a few correctness/usability issues already identified in the audit. We need a safe, incremental cleanup path that improves reliability without rewriting the whole project.

## What Changes
- Fix correctness issues from the audit in a minimal, reversible way.
- Apply the approved safe cleanup items only where usage is confirmed.
- Execute refactor in the approved order, starting with low-risk structural extraction.
- Improve color configuration UX by using color names/labels in the app instead of raw ARGB literals like `FF4CAF50`.
- Preserve current matching behavior unless explicitly fixed by this change.

## Impact
- Affected specs: matching configuration, quota assignment, export consistency, GUI usability, maintainability.
- Affected code: `matching_nomes_gui_v2.py`, `README.md`, optional cleanup targets validated as unused.

## ADDED Requirements
### Requirement: Safe Refactor Execution Order
The system SHALL apply cleanup/refactor changes in guarded phases, with correctness fixes first and structural refactor second.

#### Scenario: Execute approved order
- **WHEN** the implementation starts
- **THEN** it MUST first fix confirmed correctness/runtime risks
- **AND** only then proceed to safe cleanup and modular refactor steps
- **AND** keep behavior stable for users.

### Requirement: Human-Friendly Color Configuration
The system SHALL allow users to configure match colors using human-friendly color names/labels in the GUI while preserving export formatting output.

#### Scenario: Select and store a named color
- **WHEN** the user configures a color in the app
- **THEN** the GUI shows a readable color name/label
- **AND** the application resolves it to the correct workbook fill code during export.

## MODIFIED Requirements
### Requirement: Reuse Limit Semantics
The existing T2 reuse setting SHALL behave as a true limit (cap), not as a minimum floor.

#### Scenario: Cap reuse by configured limit
- **WHEN** reuse is enabled and a max reuse value is set
- **THEN** each normalized T2 target cannot exceed that configured maximum during assignment.

### Requirement: Export Consistency Under Background Processing
The existing export flow SHALL prevent concurrent manual mutations that can alter export output mid-run.

#### Scenario: Lock mutable review actions during export
- **WHEN** export starts in background thread
- **THEN** all manual decision controls that mutate result state are disabled until export completes/fails.

## REMOVED Requirements
### Requirement: Undeclared Legacy Helpers (Conditional)
**Reason**: Functions with no confirmed in-repo usage increase maintenance surface.
**Migration**: Before removal, verify no external scripts import these helpers; if external usage exists, keep and document them as public API.
