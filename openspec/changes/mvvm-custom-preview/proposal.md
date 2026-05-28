# Proposal: MVVM Architecture and Custom PySide6 Label Preview

## 1. Intent
The goal of this major architectural refactor is to ensure long-term scalability, maintainability, and reliability of the application. This will be achieved by migrating from the current MVC/Layered architecture to the Model-View-ViewModel (MVVM) pattern. Furthermore, the refactor addresses cross-machine stability issues by completely dropping the Microsoft Word COM dependency used for label previews and printing. Instead, a native Custom Label Preview Subwindow will be built directly in PySide6, enabling direct printing to the target SATO WS408 printer.

## 2. Scope
- **Architecture Migration (MVVM):** 
  - Restructure the codebase to strictly separate Models (data/logic), Views (PySide6 UI), and ViewModels (presentation logic and state management).
  - Implement data binding mechanisms (using PySide6 Signals and Slots) between Views and ViewModels.
- **Dependency Removal:** 
  - Eliminate all interactions with `win32com.client` or MS Word COM automation for generating, previewing, and printing labels.
- **Custom PySide6 Label Preview Subwindow:**
  - Retain the initial file loading window.
  - Upon generating labels, launch a new PySide6 modal/subwindow to display the visual preview of the generated labels.
  - Render embedded images correctly with precise dimensions suitable for printing (Label dimensions: 48mm width x 23mm height).
  - Add interactive buttons: "Confirmar" (Confirm) to proceed with printing, and "Rehacer" (Redo/Adjust) to allow users to adjust image positions if necessary.
- **Direct Printing Integration:** 
  - Implement direct printing to the SATO WS408 printer using `QPrinter` and related PySide6 printing APIs, bypassing external applications.

## 3. Affected Areas
- **UI Layer (Views):** Existing UI classes will be stripped of business logic. A new Preview Subwindow component will be added.
- **Presentation Logic (ViewModels):** A new layer will be created to manage state (e.g., loaded files, generated label data, preview adjustments).
- **Core Logic (Models):** Label generation logic will be isolated and refined.
- **Printing Module:** The entire printing mechanism will be rewritten to interface with `QPrinter` instead of MS Word.
- **Dependencies:** Removal of MS Word/COM related libraries (e.g., `pywin32`) from the project requirements, assuming they are not used elsewhere.

## 4. Risks & Mitigations
- **Risk:** Implementing MVVM in PySide6 requires robust signal/slot management to emulate data binding, which can introduce complexity.
  - **Mitigation:** Define a clear, standard base class for ViewModels to handle property changes and signal emissions consistently.
- **Risk:** Rendering exact physical dimensions (48mm x 23mm) on-screen and ensuring they translate perfectly to the SATO WS408 printer.
  - **Mitigation:** Utilize physical units (millimeters) explicitly in PySide6's `QPrinter` and `QGraphicsScene` / `QPainter` configurations. The "Rehacer" feature provides an operational safety net for the user to tweak positioning.
- **Risk:** Direct printing requires the SATO driver to correctly interpret the print job from the OS spooler via Qt.
  - **Mitigation:** Early prototyping of a minimal `QPrinter` job sent directly to the SATO WS408 driver to verify paper size, orientation, and resolution compatibility.

## 5. Rollback Plan
This major refactor will be executed in a dedicated feature branch (`feature/mvvm-custom-preview`). Rollback consists of reverting to the main branch (containing the current MVC/Word-based implementation) if insurmountable blockers arise during the direct printing implementation or if the MVVM complexity jeopardizes the project timeline. No destructive database or permanent environment changes are expected.

## 6. Success Criteria
1. Codebase strictly adheres to the MVVM architectural pattern.
2. MS Word COM dependencies are entirely removed from the application.
3. The custom PySide6 modal subwindow successfully displays post-generation previews.
4. Previews accurately reflect the 48mm x 23mm dimensions with correctly embedded images.
5. The UI includes functional "Confirmar" and "Rehacer" buttons for layout adjustment.
6. The application successfully and reliably prints directly to the SATO WS408 printer without opening external applications.