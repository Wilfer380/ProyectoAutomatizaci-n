# Design: Manual Filter Selection for Pasted IDs

## Technical Approach

Add a second, explicit processing mode that resolves pasted asset IDs across all filters, lets the user confirm the resolved selection, and then feeds the resulting `list[AssetRecord]` into the existing block preview/manual-print path. The current filter-based full-order mode remains the default and continues calling `ProcessService.run(...)` with `selected_filter` unchanged.

`main.py` remains only the app/worker bootstrap. `ui/main_window.py` adds the pasted-ID entry point and selection review UI. Controller/worker plumbing should pass the chosen mode and pasted IDs to `ProcessService`; service changes should be additive helpers, not rewrites of the current run path.

## Architecture Decisions

| Decision | Alternatives considered | Rationale |
|---|---|---|
| Add a manual-selection mode instead of changing the filter flow | Overload `selected_filter` or make pasted IDs replace the filter | Keeps full-order behavior untouched and rollback simple. |
| Resolve pasted IDs to `AssetRecord` before printing | Pass raw IDs into Word/Excel generation | Existing generation already expects records and validates generated assets from `AssetRecord.asset_id`. |
| Preserve pasted order with source indexes | Sort by Excel row/filter or use sets | Users explicitly need output order to match paste order; sets would lose duplicates/order. |
| Reuse `_print_blocks`, `_simulate_blocks`, and manual review callbacks | Build a separate print pipeline | Reduces COM/Word risk and preserves the current preview/manual-print behavior. |
| Surface errors at UI/controller and worker boundaries | Catch broadly inside every low-level method | Services should raise `ValidationError`/`RuntimeError` with context; UI and worker convert failures into visible messages and logs. |

## Data Flow

```text
Pasted IDs -> ValidationService.parse/validate_pasted_asset_ids
           -> ExcelService.resolve_records_by_asset_ids(ids)
           -> UI selection review, order preserved
           -> ProcessService.run_manual_selection(records, label="seleccion_manual")
           -> existing block split -> Excel label sheet -> Word preview/manual review -> print
```

Resolution strategy: normalize pasted lines with `normalize_excel_scalar`, keep `(source_index, asset_id)`, reject blanks after normalization, detect duplicate pasted IDs explicitly, scan `Hoja1` once into an asset-id lookup, then emit records by pasted source index. Missing IDs MUST be reported with the pasted position and ID. If duplicates are allowed later, each occurrence still keeps its own source index.

## File Changes

| File | Action | Description |
|---|---|---|
| `main.py` | Modify | Add worker CLI args only if needed, e.g. `--mode` and a safe pasted-ID payload path; keep existing args/defaults. |
| `ui/main_window.py` | Modify | Add manual mode input/review controls and signals; keep filter controls and default start button behavior unchanged. |
| `controllers/main_controller.py` | Modify | Orchestrate mode selection, validation, worker payload creation, and UI error surfacing. |
| `services/worker_client.py` / `services/worker_process.py` | Modify | Pass manual mode and pasted IDs to worker without changing current run-process contract defaults. |
| `services/excel_service.py` | Modify | Add `get_records_by_asset_ids(pasted_ids: list[str]) -> list[AssetRecord]` scanning all filters and preserving caller order. |
| `services/validation_service.py` | Modify | Add pasted-ID parsing/validation; raise `ValidationError` with user-readable missing/duplicate/empty-input messages. |
| `services/process_service.py` | Modify | Add `run_manual_selection(...)` or internal `run_with_records(...)` that reuses existing temp copy, validation, block, simulation, and print methods. |
| `services/word_service.py` | No functional change | Reused through existing block generation. |
| `services/print_service.py` | No functional change | Reused through existing print boundary. |
| `models/asset_record.py` | No change preferred | Current fields are sufficient; order should be carried by list ordering, not model mutation. |
| `tests/` | Modify/Add | Unit tests for parsing, resolution order, missing IDs, and process-service delegation seams. |

## Interfaces / Contracts

- `ValidationService.parse_pasted_asset_ids(text: str) -> list[str]`: returns normalized IDs in pasted order; raises `ValidationError` for empty input and malformed entries.
- `ExcelService.get_records_by_asset_ids(asset_ids: list[str]) -> list[AssetRecord]`: returns records in the same order as `asset_ids`; raises `ValidationError` for missing IDs or ambiguous duplicates in Excel.
- `ProcessService.run_manual_selection(..., selected_records: list[AssetRecord], selection_label: str = "seleccion_manual", ...) -> None`: shares the current temp-copy, COM, block, simulation, and print code path after record resolution.

## Error Handling

| Boundary | Catch | Surface |
|---|---|---|
| UI slots/controller actions | `ValidationError`, then unexpected `Exception` | `MainWindow.show_error(...)`, append log, reset controls; never crash Qt. |
| Worker process entry | `ValidationError`, `ManualAdjustmentCancelled`, unexpected `Exception` | Emit JSON `error`/`cancelled` with `ErrorGuard.friendly_message`, details only in logs. |
| Excel resolution | COM/open/read failures, missing/duplicate IDs | Raise `ValidationError` with pasted positions; close Excel in `finally`. |
| Process execution | File copy, COM, block/image errors | Keep existing try/finally cleanup; wrap unexpected errors as `ValidationError`. |
| Word/print | Word COM failures, printer failures | Preserve current Word fallback where available; otherwise raise friendly process error. |

## Testing Strategy

| Layer | What to Test | Approach |
|---|---|---|
| Unit | pasted-ID normalization, empty input, duplicates/missing IDs | Strict TDD with `unittest` before implementation. |
| Unit | `ExcelService` order-preserving resolver | Fake worksheet/service data; no COM required. |
| Unit | `ProcessService` accepts preselected records and still chunks by `BLOCK_SIZE` | Mock Excel/Word/Print services. |
| Controller | manual mode sends payload and validation failures keep UI alive | Mock window/worker client. |
| Regression | existing full-order path still calls current `run(... selected_filter ...)` | Existing tests plus a new guard test. |

## Migration / Rollout

No migration required. Rollback is hiding/removing the manual mode entry point; the full-order flow remains intact.

## Open Questions

- [ ] The OpenSpec delta spec was not present on disk during design; align tasks with it if it exists elsewhere.
