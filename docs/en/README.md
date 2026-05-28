# SATO WS408 Label Generator — Technical and Operational Documentation

This desktop application loads an inventory Excel file, lets the user select filters and individual labels, previews every selected label exactly through the print renderer, and sends the approved batch directly to a **SATO WS408** printer using **48 mm x 23 mm** physical labels.

## Quick path

1. Install dependencies: `pip install -r requirements.txt`.
2. Run: `python main.py`.
3. Select the inventory Excel file.
4. Open **Select filters**.
5. Check full filters or double-click a filter to choose individual labels.
6. Press **Generate labels**.
7. Review the full scrollable preview.
8. Confirm with **Confirm and print** or cancel with **Reject**.

## What the application does

| Stage | Behavior |
|---|---|
| Input | Reads Excel data from `Hoja1`; the `Seccion` column powers the filters. |
| Selection | Supports full-filter selection and individual per-label selection. |
| Preview | Renders all selected labels in a white scrollable preview window. |
| Approval | The user either prints the whole approved batch or rejects/cancels it. |
| Printing | Sends each label to `SATO WS408` through `QPrinter` and `QPainter`. |
| Errors | Shows friendly guidance when the driver is missing or Excel cannot be read. |

## What it intentionally does not do

- It does not use Microsoft Word.
- It does not use COM automation for generation, preview, or printing.
- It does not print without visual approval.
- It does not install the SATO driver automatically; it guides users to IT/support.

## Architecture

The app uses a simple **MVVM** structure:

| Layer | Responsibility | Main files |
|---|---|---|
| UI / Views | Widgets, dialogs, and visual events. | `ui/main_window.py`, `ui/filter_selection_dialog.py`, `ui/label_selection_dialog.py`, `ui/preview_subwindow.py` |
| ViewModels | Screen state, signals, and conversion into label presentation models. | `view_models/main_view_model.py`, `view_models/preview_view_model.py`, `view_models/label_item_view_model.py` |
| Services | Excel reading, rendering, printing, and printer-driver validation. | `services/excel_service.py`, `services/print_service.py`, `services/driver_check.py` |
| Models | Plain business/configuration data. | `models/asset_record.py`, `models/app_settings.py` |
| Deploy | Installer, launcher, and printer preflight. | `deploy/` |

Diagram: [`../shared/diagrams/c4-container.mmd`](../shared/diagrams/c4-container.mmd)

## MVVM flow

1. The view emits events: select Excel, select filters, generate labels.
2. `MainViewModel` loads records, groups filters, and stores the current selection.
3. The view reacts to signals (`filtersLoaded`, `recordCountChanged`, `processingFinished`).
4. `PreviewViewModel` receives the labels that are ready for review.
5. The preview dialog invokes injected callbacks for rendering and printing.

Diagram: [`../shared/diagrams/mvvm-signal-flow.mmd`](../shared/diagrams/mvvm-signal-flow.mmd)

## Excel extraction

Excel is the only supported source. Expected contract:

| Data | Source |
|---|---|
| Asset ID | `Activo fijo` column |
| Description | `Denominación del activo fijo` column |
| Filter/section | `Seccion` column from `Hoja1` |
| Image/logo | Image anchored to the same row as the record |

`ExcelService` uses `openpyxl` so the app does not need Excel to be running. Images are mapped by their anchor row and converted to `QImage`, which can then be used by both preview and printing.

Diagram: [`../shared/diagrams/excel-extraction-flow.mmd`](../shared/diagrams/excel-extraction-flow.mmd)

## Filter and label selection

The selection workflow is designed to prevent accidental printing:

1. The user opens **Select filters**.
2. The app shows each `Seccion` value with its label count.
3. If the user checks a filter without drilling down, all labels in that filter are printed.
4. If the user double-clicks a filter, a per-label checklist opens.
5. The app prints only the explicitly selected labels.

Diagram: [`../shared/diagrams/filter-label-checklist-flow.mmd`](../shared/diagrams/filter-label-checklist-flow.mmd)

## Preview and approval

The preview window is a safety gate before printing:

- White background to resemble the physical label.
- Vertical scroll for large batches, for example 179 labels.
- Shows several labels at once.
- Uses the same `LabelRenderer` as real printing.
- **Confirm and print** sends the full batch to the printer.
- **Reject** cancels without printing.

Diagram: [`../shared/diagrams/preview-approval-state.mmd`](../shared/diagrams/preview-approval-state.mmd)

## SATO WS408 printing

Approved configuration:

| Parameter | Value |
|---|---|
| Printer | `SATO WS408` |
| Physical size | `48 mm x 23 mm` |
| Resolution | `203 DPI` |
| Rendered size | approximately `384 x 184 px` |
| Margins | `0 mm` |
| Batch strategy | one independent print job per label (`separate_jobs=True`) |

Printing uses `QPrinter` and `QPainter`, with no Word or external application. Each label is sent as an independent job because real SATO testing showed this strategy keeps placement stable and avoids accumulated driver drift.

Diagram: [`../shared/diagrams/sato-print-pipeline.mmd`](../shared/diagrams/sato-print-pipeline.mmd)

## Decisions and rationale

| Decision | Rationale |
|---|---|
| Remove Word/COM | Word failed across machines, added external dependencies, and made preview/printing harder to control. |
| Use `openpyxl` | Reads Excel rows and images without launching Excel; improves testability. |
| Use native PySide6 | Keeps UI, preview, and printing inside one runtime. |
| Adopt MVVM | Keeps widgets lighter and separates state, events, and services. |
| Share renderer between preview and printing | Prevents users from approving an image that differs from the printed output. |
| Validate driver before printing | Fails early with actionable guidance for users and IT. |
| Print as separate jobs | Stabilizes SATO WS408 batches and avoids page-origin drift. |
| Use filter and per-label checklists | Reduces human error when filters contain many labels. |

## Best practices applied

- Clear separation between UI, ViewModels, services, and models.
- Callback injection for preview/printing, making tests easier.
- PySide6 signals/slots instead of direct widget-to-service coupling.
- User-facing, actionable error messages.
- Unit tests for extraction, ViewModels, preview, driver failure, and printing service.
- Mandatory visual preview before printing.
- Explicit printer name, physical size, DPI, and margins.

## Error handling

| Error | Expected response |
|---|---|
| Missing or invalid Excel | Show a friendly error and stop. |
| No records | Do not print; inform that no labels are available. |
| Missing SATO driver | Ask the user to install/configure `SATO WS408` or contact IT. |
| Print startup failure | Show the error and allow retry. |
| Unexpected exception | Save technical detail to logs and show a safe message. |

Diagram: [`../shared/diagrams/error-handling-boundaries.mmd`](../shared/diagrams/error-handling-boundaries.mmd)

## Installation and deployment

The deployment flow includes a printer-driver preflight:

1. The app is installed or copied.
2. The launcher/preflight checks whether Windows knows the `SATO WS408` printer.
3. If the driver is missing, installation may still complete, but real printing requires IT/support.
4. End users should not manually handle printer drivers if they lack permissions.

Diagram: [`../shared/diagrams/deployment-driver-preflight.mmd`](../shared/diagrams/deployment-driver-preflight.mmd)

## Maintainer validation

Before opening or merging changes:

```powershell
python -m unittest discover -s tests
```

Review checklist:

- [ ] Word/COM was not reintroduced.
- [ ] Preview and printing still share the same renderer.
- [ ] Label size remains 48 mm x 23 mm.
- [ ] Target printer remains `SATO WS408`.
- [ ] Filter/label selection does not print unselected records.
- [ ] Error messages remain clear for non-technical users.

## Troubleshooting

| Symptom | Likely cause | Action |
|---|---|---|
| Printer is not detected | Driver missing or printer name differs. | Install/configure `SATO WS408` through IT/support. |
| Excel loads no filters | `Seccion` column is missing or empty. | Check `Hoja1` and headers. |
| Label preview is empty | Missing asset data or incorrectly anchored image. | Check the Excel row. |
| Batch output drifts | Driver/configuration changed. | Keep `separate_jobs=True` and validate on the real SATO printer. |
| Preview differs from print | Renderer or DPI changed. | Check `LabelRenderer` and `LabelPrintConfig`. |
