# Label Printing Specification

## Purpose

To manage the extraction of asset data and images from an Excel source, generate an interactive PySide6 preview of the resulting physical labels, and print them directly to a SATO WS408 printer without relying on external software like Microsoft Word. This adheres strictly to an MVVM architectural pattern.

## Requirements

### Requirement: Excel Sole Source Input

The system MUST accept an Excel file (e.g., `Inventario de activos fijos 2026.xlsx`) as the sole input for label data and images. The system MUST NOT require, request, or load Microsoft Word templates.

#### Scenario: User selects input file

- GIVEN the user launches the application
- WHEN the initial file loading window is displayed
- THEN the UI only prompts the user to select an Excel file
- AND the UI does not show any file selector for a Word template

### Requirement: MVVM Architecture Adherence

The system MUST isolate data logic (Model), presentation state and logic (ViewModel), and the user interface (View) using PySide6 signals and slots for data binding.

#### Scenario: Label state update triggers UI refresh

- GIVEN the application is running
- WHEN the ViewModel updates the state of the loaded data or preview adjustments
- THEN the View (UI) automatically reflects the new state via signal/slot connections without containing business logic itself

### Requirement: Native Label Mapping and Rendering

The system MUST extract data and images from the Excel file and natively map them to a strict 48mm width x 23mm height physical dimension suitable for the SATO WS408 printer.

#### Scenario: Extracting and mapping data

- GIVEN an Excel file containing asset details and images has been loaded
- WHEN the system processes the file
- THEN the application internally maps the text and images to a 48mm x 23mm coordinate space
- AND the images are scaled correctly to fit within the physical dimensions

### Requirement: PySide6 Custom Preview Subwindow

The system MUST launch a native PySide6 modal subwindow to display the visual preview of the generated labels immediately after processing the source file.

#### Scenario: Launching the preview modal

- GIVEN the system has finished generating the native label layouts from the Excel data
- WHEN the label generation is complete
- THEN a modal subwindow automatically opens displaying a visual preview of the labels
- AND the preview accurately reflects the physical 48mm x 23mm dimensions

### Requirement: Interactive Preview Controls

The preview modal MUST provide "Confirmar" and "Rehacer" controls to allow the user to either approve the print job or adjust the layout.

#### Scenario: Adjusting the layout

- GIVEN the preview modal is open
- WHEN the user clicks the "Rehacer" (Redo/Adjust) button
- THEN the user is permitted to tweak image positions or label layouts

#### Scenario: Confirming the print job

- GIVEN the preview modal is open and the layout is satisfactory
- WHEN the user clicks the "Confirmar" (Confirm) button
- THEN the system initiates the direct print process

### Requirement: Direct Printer Integration

The system MUST print the confirmed labels directly to the target SATO WS408 printer using `QPrinter` and PySide6 printing APIs, bypassing MS Word and external COM dependencies.

#### Scenario: Executing direct print

- GIVEN the user has clicked "Confirmar" in the preview modal
- WHEN the direct print process is initiated
- THEN the system constructs a Qt print job and sends it directly to the OS spooler for the SATO WS408
- AND no MS Word process or external application is spawned or required