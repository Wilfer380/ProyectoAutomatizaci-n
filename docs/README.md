# Documentation / Documentación

This folder explains how the SATO WS408 label generator works, why it was built this way, and how to operate, maintain, and review it.

Esta carpeta explica cómo funciona el generador de etiquetas para SATO WS408, por qué se construyó así y cómo operarlo, mantenerlo y revisarlo.

## Languages / Idiomas

| Language | Document |
|---|---|
| Español | [`docs/es/README.md`](es/README.md) |
| English | [`docs/en/README.md`](en/README.md) |

## Shared diagrams / Diagramas compartidos

The diagrams are stored as Mermaid source files so GitHub can render or preview them with Mermaid-compatible tooling.

Los diagramas están guardados como archivos Mermaid para que GitHub o herramientas compatibles puedan renderizarlos.

| Diagram | Purpose |
|---|---|
| [`c4-container.mmd`](shared/diagrams/c4-container.mmd) | Main containers and external dependencies. |
| [`mvvm-signal-flow.mmd`](shared/diagrams/mvvm-signal-flow.mmd) | PySide6 MVVM signals/slots flow. |
| [`excel-extraction-flow.mmd`](shared/diagrams/excel-extraction-flow.mmd) | Excel row/image extraction. |
| [`filter-label-checklist-flow.mmd`](shared/diagrams/filter-label-checklist-flow.mmd) | Filter and per-label checklist selection. |
| [`preview-approval-state.mmd`](shared/diagrams/preview-approval-state.mmd) | Preview acceptance/rejection state machine. |
| [`sato-print-pipeline.mmd`](shared/diagrams/sato-print-pipeline.mmd) | SATO WS408 print pipeline. |
| [`error-handling-boundaries.mmd`](shared/diagrams/error-handling-boundaries.mmd) | User-facing error boundaries. |
| [`deployment-driver-preflight.mmd`](shared/diagrams/deployment-driver-preflight.mmd) | Installer and printer-driver preflight. |
