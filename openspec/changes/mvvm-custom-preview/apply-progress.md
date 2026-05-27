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

### PR 2: ViewModels & Data Binding (MVVM Core)
- [x] **Implement LabelItemViewModel**: Created `view_models/label_item_view_model.py` and its tests. Added layout/scale properties.
- [x] **Implement MainViewModel**: Updated `view_models/main_view_model.py` to use `ExtractionWorker` internally, mapped properties correctly, fixed QThread crash by avoiding passing objects between threads directly, and refactored tests in `tests/test_main_view_model.py` to correctly initialize `QApplication` and await signals properly.
- [x] **Implement PreviewViewModel**: Created `view_models/preview_view_model.py` with state and navigation for pages. Implemented `set_items`, `next_page`, `previous_page`, `confirm`, and `redo` functions with respective signals (`previewReady`, `printStarted`, `printCompleted`, `redoRequested`). Wrote tests in `tests/test_preview_view_model.py`.

### PR 3: Custom PySide6 Preview Subwindow UI
- [x] **Create Preview Subwindow**: Created `ui/preview_subwindow.py` as a `QDialog` with a `QGraphicsScene`/`QGraphicsView` using a logical 48x23mm preview scale. Added `Confirmar` and `Rehacer` buttons connected to `PreviewViewModel.confirm()` and `PreviewViewModel.redo()`.
- [x] **Update Main Window View**: Updated `ui/main_window.py` to accept a `MainViewModel`, bind Excel/progress/processing signals, remove Word UI controls from the new MVVM path, and keep the main UI responsive while processing runs through the ViewModel worker.

## TDD Cycle Evidence
| Task | Target File | RED (Failing Test) | GREEN (Passing Test) | REFACTOR |
|---|---|---|---|---|
| Excel Extraction Rewrite | `services/excel_service.py` | `tests/test_excel_service.py` failing on `extract_data` | Implemented `openpyxl` parser & image mapper | `AssetRecord` modified to use `Any \| None` |
| ViewModels Setup | `view_models/main_view_model.py` | Tests crashing on Qt `QEventLoop` | Fixed `QThread` lifecycle & added `QGuiApplication` to test context | Instantiated ViewModels in main thread instead of worker |
| PreviewViewModel | `view_models/preview_view_model.py` | `tests/test_preview_view_model.py` failing on missing attributes | Implemented `confirm`, `redo`, navigation, & signals | Organized state into properties |
| Preview Subwindow | `ui/preview_subwindow.py` | `tests/test_preview_subwindow.py` failing on missing UI | Implemented modal preview dialog and connected action buttons | Added 48x23 logical preview scene |
| Main Window MVVM Binding | `ui/main_window.py` | `tests/test_main_window.py` failing on legacy Word UI and constructor | Added optional `MainViewModel` binding and removed Word controls from the MVVM UI path | Fixed test app initialization to use `QApplication` instead of `QGuiApplication` |

## Deviations from Design
- `AssetRecord` model was updated to include an `image` field to carry the `QImage` as specified by the prompt.
- `MainViewModel` worker logic was updated to emit data models (`records`) instead of view models, and the main thread instantiates the `LabelItemViewModel` to prevent cross-thread object passing which caused a C++ segfault.

## Remaining Tasks
- Implement Direct Print Engine
- Finalize End-to-End Integration
- Add SATO WS408 driver detection and installer/launcher robustness improvements

## Workload / PR Boundary
- PR 3 is complete. The next step is PR 4 (Direct Print Engine and Integration), followed by PR 5 (UI/UX hardening, driver detection, installer/launcher robustness).
