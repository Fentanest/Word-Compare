# Refactoring Log

## 2026-04-28

### Goal
- Keep the user-visible behavior from `GEMINI.md` unchanged.
- Move the codebase from a single-file controller toward layered modules.

### Completed
- Split the runtime into `app/ui`, `app/services`, and `app/reports`.
- Reduced [main.py](/home/better0101/projects/word-compare/main.py) to a thin entrypoint.
- Moved settings persistence into `SettingsService`.
- Moved Word COM lifecycle into `WordSession`.
- Moved comparison orchestration into `WordCompareService`.
- Moved hybrid extraction into `DocxExtractor`.
- Split Excel reporting into:
  - `ExcelDiffEngine`
  - `ExcelReportWriter`
  - `ExcelReportService`
- Kept [excel_generator.py](/home/better0101/projects/word-compare/excel_generator.py) as a compatibility wrapper.
- Started separating local debug scripts into `tools/` while leaving root wrappers temporarily.

### Batch 2 Details
- Added report input and diff-plan models in `app/reports/models.py`.
- Moved opcode calculation out of the writer into `app/reports/diff_engine.py`.
- Moved workbook rendering into `app/reports/excel_writer.py`.
- Updated `ExcelReportService` so the UI/service layer no longer depends on the old monolithic report function.
- Replaced the old `GEMINI.md` note with a fuller project guide tied to the current structure.
- Added transitional `tools/` modules and left root-level wrappers in place to avoid breaking existing habits immediately.

### Current Structure
- `app/ui/main_window.py`: UI event handling and model wiring
- `app/services/*`: Word automation, settings, document extraction
- `app/reports/*`: report input models, diff planning, workbook rendering
- `tools/*`: local diagnostics and simulation helpers

### Next Candidates
- Remove root wrapper scripts after confirming no one depends on them.
- Add lightweight smoke tests around report generation.
- Revisit `main.spec` and packaging once the folder layout stabilizes.
- Consider moving generated/test artifacts into `samples/` or a scratch directory.

### Notes
- `main_ui.py` remains generated code and should continue to be treated as read-only.
- The refactor is intentionally behavior-preserving first, cleanup second.
