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

### PR 4: Direct Print Engine and Integration
- [x] **Implement Direct Print Engine**: Created `services/print_service.py` with `LabelPrintConfig`, `LabelRenderer`, and `PrintService`. The printer is configured as `SATO WS408`, with exact `QPageSize(QSizeF(48.0, 23.0), QPageSize.Unit.Millimeter)` and zero margins.
- [x] **Finalize End-to-End Integration**: Updated `PreviewViewModel.confirm()` to call an injected print callback and wired `MainWindow` preview creation to `PrintService.print_labels()`. Updated `main.py` to start the MVVM path with `MainViewModel()` instead of the legacy controller.

### PR 5: UI/UX, Driver Detection, Installer Robustness, and Legacy Cleanup
- [x] **Driver Detection**: Added `services/driver_check.py` using `QPrinterInfo.availablePrinterNames()` and friendly guidance for missing SATO WS408 driver.
- [x] **Print Error UX**: `PrintService` refuses to print if the driver is missing; `PreviewViewModel` emits `printFailed`; `PreviewSubwindow` shows a clear modal instead of crashing.
- [x] **Installer/Launcher Hardening**: Added installer preflight via PowerShell `Get-Printer`, removed pywin32 shortcut dependency, improved launcher process errors, and updated packaging specs.
- [x] **Legacy Cleanup**: Deleted stale controller/process/worker modules that imported the removed Word service and updated README to the Excel-only MVVM flow.

## TDD Cycle Evidence
| Task | Target File | RED (Failing Test) | GREEN (Passing Test) | REFACTOR |
|---|---|---|---|---|
| Excel Extraction Rewrite | `services/excel_service.py` | `tests/test_excel_service.py` failing on `extract_data` | Implemented `openpyxl` parser & image mapper | `AssetRecord` modified to use `Any \| None` |
| ViewModels Setup | `view_models/main_view_model.py` | Tests crashing on Qt `QEventLoop` | Fixed `QThread` lifecycle & added `QGuiApplication` to test context | Instantiated ViewModels in main thread instead of worker |
| PreviewViewModel | `view_models/preview_view_model.py` | `tests/test_preview_view_model.py` failing on missing attributes | Implemented `confirm`, `redo`, navigation, & signals | Organized state into properties |
| Preview Subwindow | `ui/preview_subwindow.py` | `tests/test_preview_subwindow.py` failing on missing UI | Implemented modal preview dialog and connected action buttons | Added 48x23 logical preview scene |
| Main Window MVVM Binding | `ui/main_window.py` | `tests/test_main_window.py` failing on legacy Word UI and constructor | Added optional `MainViewModel` binding and removed Word controls from the MVVM UI path | Fixed test app initialization to use `QApplication` instead of `QGuiApplication` |
| Direct Print Service | `services/print_service.py` | `tests/test_print_service.py` failing before service existed | Implemented `QPrinter` configuration, print callback integration, and `QPainter` label rendering | Added fake printer/painter tests so no real printer is required during CI |
| Driver Detection | `services/driver_check.py` | `tests/test_driver_check.py` failing before service existed | Implemented QPrinterInfo-backed detection and friendly IT guidance | Reused in `PrintService` and `ValidationService` |
| Installer Preflight | `deploy/printer_driver_preflight.py` | `tests/test_printer_driver_preflight.py` failing before helper existed | Implemented PowerShell Get-Printer preflight and installer guidance | Removed pywin32 shortcut dependency from installer |

## Deviations from Design
- `AssetRecord` model was updated to include an `image` field to carry the `QImage` as specified by the prompt.
- `MainViewModel` worker logic was updated to emit data models (`records`) instead of view models, and the main thread instantiates the `LabelItemViewModel` to prevent cross-thread object passing which caused a C++ segfault.

## Remaining Tasks
- Perform real hardware smoke test on a Windows machine with the SATO WS408 driver installed.
- Optionally add the real SATO WS4 driver installer artifact under a `drivers/` release folder if the organization approves redistribution.

## Workload / PR Boundary
- PR 5 is complete. The SDD apply phase is ready for fresh review and verify.
