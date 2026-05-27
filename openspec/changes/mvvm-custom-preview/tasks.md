## Review Workload Forecast

| Field | Value |
|-------|-------|
| Estimated changed lines | 600 - 800 |
| 400-line budget risk | High |
| Chained PRs recommended | Yes |
| Suggested split | PR 1 (Data Extraction) → PR 2 (ViewModels) → PR 3 (Preview UI) → PR 4 (Print Service & Integration) |
| Delivery strategy | auto-chain |
| Chain strategy | feature-branch-chain |

Decision needed before apply: No
Chained PRs recommended: Yes
Chain strategy: feature-branch-chain
400-line budget risk: High

## Implementation Tasks

### PR 1: Dependency Cleanup and Extraction Service Update
1. **Update Dependencies**: 
   - Target: `requirements.txt`
   - Action: Remove `pywin32`, `win32com`, and `python-docx` dependencies. Add `openpyxl>=3.1.2`.
2. **Remove Deprecated Word Service**: 
   - Target: `services/word_service.py` (or equivalent file)
   - Action: Delete the file completely to ensure no COM logic remains in the codebase.
3. **Rewrite Excel Extraction Service**: 
   - Target: `services/excel_service.py`
   - Action: Replace COM/interop usage with native `openpyxl` parsing. 
   - Ensure text data is extracted by iterating through rows in `Hoja1`. 
   - Extract anchored floating images (`ws._images`), map them to their respective data rows using anchor coordinates (row/col).
   - Convert image bytes into `QImage.fromData(bytes)` format directly.

### PR 2: ViewModels & Data Binding (MVVM Core)
4. **Implement LabelItemViewModel**: 
   - Target: `view_models/label_item_view_model.py`
   - Action: Create class inheriting from `QObject`. Define state properties (`asset_id`, `asset_name`, `section`, `image_data`, and offsets: `image_offset_x`, `image_offset_y`, `image_scale`). Add a `layoutChanged()` signal.
5. **Implement MainViewModel**: 
   - Target: `view_models/main_view_model.py`
   - Action: Create class to replace the existing main controller. Manage state (`selected_file_path`, `is_processing`, `progress_value`). Add signals (`fileSelected`, `processingStarted`, `processingFinished`, etc.). Wire it to trigger `excel_service.py` and emit the list of `LabelItemViewModel` objects on completion.
6. **Implement PreviewViewModel**: 
   - Target: `view_models/preview_view_model.py`
   - Action: Manage state for previewing generated labels (`label_items`, `current_page_index`). Add signals for `previewReady()`, `printStarted()`, and `printCompleted()`. Implement logic to handle `Confirmar` (forward to print service) and `Rehacer` actions.

### PR 3: Custom PySide6 Preview Subwindow UI
7. **Create Preview Subwindow**: 
   - Target: `views/preview_subwindow.py`
   - Action: Implement a modal PySide6 `QDialog` or `QMainWindow`. 
   - Incorporate `QGraphicsScene` and `QGraphicsView` to render the label preview. Enforce an exact 48mm x 23mm visual representation area for the current `LabelItemViewModel`.
   - Add interactive "Confirmar" and "Rehacer" buttons, connected to `PreviewViewModel` slots.
8. **Update Main Window View**: 
   - Target: `views/main_window.py`
   - Action: Remove existing business/controller logic. Bind the file selection and progress UI components to `MainViewModel` signals. When the `processingFinished` signal is received, trigger the instantiation and execution of `preview_subwindow.py`.

### PR 4: Direct Print Engine and Integration
9. **Implement Direct Print Engine**: 
   - Target: `services/print_service.py`
   - Action: Create a new native printing service.
   - Use `QPrinter(QPrinter.PrinterMode.HighResolution)` configured for the "SATO WS408" driver.
   - Set hardware dimensions strictly using `QPageSize(QSizeF(48.0, 23.0), QPageSize.Unit.Millimeter)`. Disable margins.
   - Use `QPainter` to directly render text and `QImage` representations onto the print spooler for each `LabelItemViewModel`.
10. **Finalize End-to-End Integration**: 
    - Target: Application Entry Point (`main.py`) and `PreviewViewModel`.
    - Action: Wire the "Confirmar" action directly to invoke the `print_labels` function from `print_service.py`. Ensure all pieces (extraction -> view model -> preview UI -> print) function sequentially without errors.

### PR 5: UI/UX, Driver Detection, and Installer Robustness
11. **UI/UX Polishing and Screen Freeze Prevention**:
    - Target: `views/main_window.py` and `view_models/main_view_model.py`.
    - Action: Ensure all heavy operations run on `QThread` (already initiated in PR 2) to prevent the main UI from freezing. Improve visual feedback (spinners, clear status messages) during the extraction.
12. **Driver Detection (SATO WS408)**:
    - Target: `services/print_service.py` or new `utils/driver_check.py`.
    - Action: Check available OS printers (`QPrinterInfo.availablePrinterNames()`). If "SATO WS408" is missing, throw a custom error to the ViewModel.
    - The UI must display an error modal telling the user: "Controlador SATO WS408 no detectado. Por favor instálelo o contacte a TI."
13. **Installer and Launcher Improvements**:
    - Target: `.spec` files (`GeneradorEtiquetasSAP.spec`, etc.) and deployment scripts.
    - Action: Ensure the installer packages correctly, cleans up logs gracefully, and has better unhandled exception hooks (already partially in `logger.py` but we need to ensure the user gets a friendly dialog instead of a silent crash). Add a README note about the driver prerequisite for the installer.