# Apply Progress: MVVM Custom Preview

## Completed Tasks
### PR 1: Dependency Cleanup and Extraction Service Update
- [x] **Update Dependencies**: Removed `pywin32`, `win32com`, and `python-docx` dependencies. Added `openpyxl>=3.1.2` to `requirements.txt`.
- [x] **Remove Deprecated Word Service**: Deleted `services/word_service.py`, `tests/test_word_service_no_com.py`, `tests/test_process_service_word_com.py`, `tests/test_excel_service_hidden.py`, `word_smoke_test.py`, and `utils/win32_bootstrap.py`.
- [x] **Rewrite Excel Extraction Service**:
  - Replaced COM/interop usage in `services/excel_service.py` with native `openpyxl` parsing.
  - Implemented iteration through rows in `Hoja1` to extract text data based on headers.
  - Implemented extraction of anchored floating images (`ws._images`), mapped them to respective data rows using anchor coordinates.
  - Converted image bytes into `QImage.fromData(bytes)` format directly. Added `image` property to `AssetRecord`.
  - Wrote TDD tests in `tests/test_excel_service.py` to verify the `openpyxl` extraction.

## TDD Cycle Evidence
| Task | Target File | RED (Failing Test) | GREEN (Passing Test) | REFACTOR |
|---|---|---|---|---|
| Excel Extraction Rewrite | `services/excel_service.py` | `tests/test_excel_service.py` failing on `extract_data` | Implemented `openpyxl` parser & image mapper | `AssetRecord` modified to use `Any \| None` |

## Deviations from Design
- `AssetRecord` model was updated to include an `image` field to carry the `QImage` as specified by the prompt.

## Remaining Tasks
- Implement `LabelItemViewModel`
- Implement `MainViewModel`
- Implement `PreviewViewModel`
- Create Preview Subwindow
- Update Main Window View
- Implement Direct Print Engine
- Finalize End-to-End Integration

## Workload / PR Boundary
- PR 1 is ready for review and integration. The next step is PR 2 (ViewModels & Data Binding).
