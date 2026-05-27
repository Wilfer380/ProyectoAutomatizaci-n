# Technical Design: MVVM Custom Preview

## 1. Overview
This document outlines the technical design for migrating the SAP Label Generator application to an MVVM architecture, removing Microsoft Word and COM dependencies (`pywin32`), and introducing a native PySide6 Label Preview Subwindow. The application will natively render and print 48x23mm labels directly to a SATO WS408 printer.

## 2. PySide6 ViewModels & Data Binding (MVVM)

### Data Binding Strategy
We will implement an observable pattern using PySide6's `QObject` and `Signal` / `Slot` mechanisms to ensure strict separation of concerns. Views will react to ViewModel changes without holding any business logic themselves.

- **Signals**: ViewModels emit signals (e.g., `labelsGenerated`, `printConfirmed`) to communicate state changes to the Views.
- **Slots**: Views connect these signals to UI updating methods to maintain state parity automatically.

### Specific ViewModels

**1. `MainViewModel` (Replaces `MainController`)**
- **State Properties**: `selected_file_path`, `is_processing`, `progress_value`, `status_message`.
- **Signals**: 
  - `fileSelected(str)`
  - `processingStarted()`
  - `progressUpdated(int)`
  - `processingFinished(list[LabelItemViewModel])`
  - `errorOccurred(str)`
- **Responsibilities**: Triggers the extraction service, tracks progress, and instantiates `LabelItemViewModel` objects when data extraction completes.

**2. `PreviewViewModel`**
- **State Properties**: `label_items` (list of `LabelItemViewModel`), `current_page_index`, `is_printing`.
- **Signals**:
  - `previewReady()`
  - `labelModified(int index)`
  - `printStarted()`
  - `printCompleted()`
- **Responsibilities**: 
  - Receives the list of parsed label items.
  - Manages "Confirmar" action which forwards the list to the `PrintService`.
  - Manages "Rehacer" action, allowing adjustments to the currently viewed label.

**3. `LabelItemViewModel`**
- **State Properties**: 
  - `asset_id` (Activo fijo)
  - `asset_name` (Denominación del activo fijo)
  - `section` (Seccion)
  - `image_data` (`QImage` representing the extracted asset photo)
  - `image_offset_x`, `image_offset_y`, `image_scale` (For layout fine-tuning when the user clicks "Rehacer")
- **Signals**: `layoutChanged()`

## 3. Data and Image Extraction Logic (Without COM/Word)

To completely drop `pywin32` and avoid MS Excel application instantiation:

- **Dependency Addition**: Introduce `openpyxl` to parse the source `.xlsx` file directly and natively.
- **Data Extraction**: Iterate over the rows of the configured source sheet (`Hoja1`). Extract text using existing constants defined in `SOURCE_HEADERS`.
- **Image Extraction**: 
  - Floating images in Excel sheets are parsed natively using `ws._images` from `openpyxl`.
  - Each image object from `openpyxl` contains an `anchor` (e.g., `_from.row` and `_from.col`) indicating the top-left cell the image is anchored to.
  - Map each extracted image to its corresponding `AssetRecord` by matching the image's row anchor to the data row index.
  - Read the binary image data and convert it immediately into a UI-ready structure via `QImage.fromData(bytes)`.

## 4. QPrinter Configuration and Layout Mapping

The requirement states the physical label is 48mm width x 23mm height, printed directly to the SATO WS408.

### Native Layout Mapping (48x23mm)
Instead of relying on Excel shape groups ("Group 18", etc.) or Word COM templates:
- A PySide6 `QGraphicsScene` combined with a `QGraphicsView` will act as the visual layout engine for the Custom Preview Subwindow.
- Dimensions are rigidly enforced: `48mm x 23mm`. 
- The internal drawing area logic:
  - **Left Block (Text)**: Render text fields (`asset_id`, `asset_name`, `section`) sequentially using standard Qt font sizing metrics.
  - **Right Block (Photo)**: The asset photo `QImage`, scaled dynamically while preserving aspect ratio, mapping firmly to the right side of the label bounds.

### Direct Printing Implementation (`QPrinter`)
```python
from PySide6.QtPrintSupport import QPrinter
from PySide6.QtGui import QPainter, QPageSize
from PySide6.QtCore import QSizeF, QMarginsF

def print_labels(view_models: list[LabelItemViewModel]):
    # Setup High Resolution printing for the SATO target
    printer = QPrinter(QPrinter.PrinterMode.HighResolution)
    printer.setPrinterName("SATO WS408") # Can be dynamic via Settings
    
    # Configure precise hardware dimensions: 48mm x 23mm
    page_size = QPageSize(QSizeF(48.0, 23.0), QPageSize.Unit.Millimeter)
    printer.setPageSize(page_size)
    
    # Remove standard margins since the layout utilizes the full 48x23mm area
    layout = printer.pageLayout()
    layout.setMargins(QMarginsF(0, 0, 0, 0))
    printer.setPageLayout(layout)
    
    # Paint directly to the print spooler
    painter = QPainter()
    painter.begin(printer)
    
    for i, label_vm in enumerate(view_models):
        if i > 0:
            printer.newPage()
            
        # Draw native text elements with explicit millimeter/point coordinate mapping
        painter.drawText(...) 
        
        # Draw scaled QImage representation of the asset
        painter.drawImage(...) 
        
    painter.end()
```

## 5. Rollout and Migration Plan
1. **Extraction Service Update (`excel_service.py`)**: Add `openpyxl`. Rewrite logic to extract text and anchored images directly without starting Excel COM objects. Remove `win32com` code.
2. **ViewModel Implementation**: Create `MainViewModel`, `PreviewViewModel`, and `LabelItemViewModel` strictly separating UI signals from extraction processes.
3. **Custom PySide6 Subwindow**: Build the modal visual preview incorporating the "Confirmar" and "Rehacer" logic connected to the `PreviewViewModel`.
4. **Direct Print Engine**: Implement a robust `PrintService` using the outlined `QPrinter` configuration over a native `QPainter` canvas.
5. **Dependency Cleanup**: Remove `pywin32` and `python-docx` from `requirements.txt` and delete `word_service.py` entirely, ensuring the system operates purely on PySide6 and openpyxl natively.