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
- Expand the new lightweight report tests into broader service/UI smoke tests.
- Revisit `main.spec` and packaging once the folder layout stabilizes.
- Consider moving generated/test artifacts into `samples/` or a scratch directory.

### Notes
- `main_ui.py` remains generated code and should continue to be treated as read-only.
- The refactor is intentionally behavior-preserving first, cleanup second.

### Test Coverage Added
- `tests/test_diff_engine.py`
  verifies paragraph filtering and diff-plan generation.
- `tests/test_excel_report_pipeline.py`
  verifies workbook generation, sheet names, and key shared strings without depending on Excel or Word.
- `tests/test_file_list_manager.py`
  verifies list item creation, sorting, and before/after file pair assembly.
- GitHub Actions build now runs `python -m unittest discover -s tests -v` before packaging.

### Batch 3 Details
- Extracted file-list item creation, sorting, and pair assembly out of `main_window.py`.
- Added `app/ui/file_list_manager.py` so the main window only coordinates UI events.

### Batch 4 Details
- Improved table alignment keys from `first-cell / first-row` heuristics to full `row / column` signatures.
- Added regression tests for row insertion and column insertion where the old heuristic was prone to misalignment.

### Batch 5 Details
- Added `TableCellData` so extracted table cells can carry merge metadata together with text.
- Extended the extractor to capture `gridSpan` and `vMerge` from table cell XML properties.
- Updated table equality/signature logic so merge-structure changes are treated as real changes even when the visible text is the same.
- Added a regression test covering merge metadata differences.

### Batch 6 Details
- Extended `TableCellData` with layout metadata for `tcW`, `trHeight`, and `tblGrid` column widths.
- Updated the DOCX extractor so table-cell signatures now preserve width and row-height changes from Word XML.
- Added regression tests covering cell-width, row-height, and table-grid width changes.

### Batch 7 Details
- Added `ParagraphData` and `RunData` so body paragraphs can carry lightweight paragraph and run formatting metadata.
- Updated the diff engine to align body paragraphs by metadata-aware signatures instead of plain text only.
- Marked format-only body changes as `[서식 변경]` in the Excel report without adding extra Word COM layout calls.
- Kept page/line location expansion out of the default path because it would require expensive per-paragraph Word layout queries.

### Batch 8 Details
- Extended DOCX extraction to parse header/footer text, footnotes, endnotes, comments, revision tags, section settings, text boxes, drawing metadata, and table-level XML metadata from the saved `.docx` package.
- Added table cell formatting metadata to comparison signatures, including borders, shading, vertical alignment, text direction, and nested-table count.
- Routed extracted metadata blocks through the existing main diff/report pipeline so structure-only changes show up in Excel as `[구조/메타데이터 변경]`.
- Added XML-package regression tests to lock in header/footer, notes, comments, revision, section, shape, and table-style coverage.
- Updated the compare service to reopen source documents specifically for Excel extraction, so tracked changes and comments can still be observed even though the compare pipeline accepts revisions on the comparison documents.

### Batch 9 Details
- Added `[성능]` timing logs around Word session startup, document open, revision normalization, Word compare, result save, Excel source reopen, and whole-file processing.
- Added `[성능]` timing logs around Excel extraction, diff-plan calculation, workbook writing, and whole-report generation.
- Added extractor-internal timing logs for number conversion, temporary DOCX save, `python-docx` load, body/table parsing, XML metadata parsing, and whole extraction.
- Kept the instrumentation in the existing UI log callback path so real-user runs can be profiled without a separate profiler setup.

### Batch 10 Details
- Trimmed `main.spec` so only `logo.png` is bundled as runtime data; `logo.ico` remains embedded only as the executable icon.
- Expanded `excludes` in `main.spec` to drop additional unused PySide6 modules and development-only packages from the PyInstaller bundle.
- Kept the exclusions aligned with the actual runtime imports, which currently only rely on `QtCore`, `QtGui`, and `QtWidgets`.

### Batch 11 Details
- Added a Rust native extractor crate under `native/docx-structure-extractor/` that parses DOCX package XML and emits structured JSON for body paragraphs, tables, and metadata blocks.
- Added `app/services/native_docx_extractor.py` so Python can discover the native binary, execute it, and hydrate the result back into the existing dataclasses.
- Updated `DocxExtractor` to try the native path first and fall back to the existing `python-docx` path automatically when the binary is missing or fails.
- Added build scripts for Windows and POSIX shells, plus PyInstaller wiring to bundle `build/native/word_compare_native_extractor.exe` when present.
