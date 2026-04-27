# Tasks
- [x] Task 1: Fix correctness and runtime safety baseline.
  - [x] SubTask 1.1: Correct T2 reuse limit logic to enforce cap semantics.
  - [x] SubTask 1.2: Add robust boolean parsing for config values (avoid `bool("False") == True` pitfalls).
  - [x] SubTask 1.3: Prevent concurrent state mutation during export by disabling manual review mutation controls while busy.
  - [x] SubTask 1.4: Improve workbook metadata error surfacing (no silent failure without actionable feedback).

- [x] Task 2: Implement color-name UX for GUI configuration.
  - [x] SubTask 2.1: Introduce a color-name map and canonical normalization strategy (name/label -> fill code).
  - [x] SubTask 2.2: Update GUI color fields to show/select readable names while preserving existing color picker support.
  - [x] SubTask 2.3: Ensure export formatting resolves names to valid fill codes with backward compatibility for existing saved state values.
  - [x] SubTask 2.4: Validate saved UI state load/save compatibility for new color representation.

- [x] Task 3: Execute safe cleanup candidates with proof.
  - [x] SubTask 3.1: Confirm usage of `run_matching` and `build_export_catalog` before any removal/refactor.
  - [x] SubTask 3.2: If confirmed unused, remove or clearly mark as internal legacy helpers; otherwise document as supported entry points.
  - [x] SubTask 3.3: Keep cleanup scoped to validated low-risk targets only.

- [x] Task 4: Start phased refactor order without behavior rewrite.
  - [x] SubTask 4.1: Extract config normalization/validation helpers into a focused section/module boundary.
  - [x] SubTask 4.2: Extract workbook metadata read helpers into a focused section/module boundary.
  - [x] SubTask 4.3: Keep matching engine and export paths behavior-equivalent while reducing coupling.
  - [x] SubTask 4.4: Keep GUI class as orchestrator, moving logic only when low-risk and verifiable.

- [x] Task 5: Verify and harden.
  - [x] SubTask 5.1: Run compile/import/runtime smoke checks with sample workbook.
  - [x] SubTask 5.2: Validate no regressions in analysis summary counts and export generation.
  - [x] SubTask 5.3: Update README usage notes for color-name configuration and any changed behavior.

# Task Dependencies
- Task 2 depends on Task 1.
- Task 3 depends on Task 1.
- Task 4 depends on Task 1 and should be incremental after baseline fixes.
- Task 5 depends on Tasks 1-4.
