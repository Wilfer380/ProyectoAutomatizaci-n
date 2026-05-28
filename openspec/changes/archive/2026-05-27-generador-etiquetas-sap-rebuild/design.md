# Design: GeneradorEtiquetasSAP Document-Generation Rebuild

## Technical Approach

Rebuild generation internals behind `services/process_service.py` while preserving the UI/controller/worker contract. `ProcessService.run()` remains the entry point used by `services/worker_process.py`, `services/worker_client.py`, and `controllers/main_controller.py`; the change replaces monolithic per-block logic with boundaries for data loading, block planning, image export, Word composition, review handoff, logging, and recoverable failures.

Excel remains a private hidden COM instance (`DispatchEx`, `Visible=False`, alerts/events/screen updates disabled). Word generation stays COM-first for fidelity, with `python-docx` fallback when COM insertion fails. Blocks stay capped by `utils.constants.BLOCK_SIZE == 27`; unused `<imgN>` placeholders are cleared through `<img27>`.

## Architecture Decisions

| Topic | Decision | Rationale |
|---|---|---|
| Public contract | Keep worker events and callback signatures unchanged: `log`, `status`, `progress`, `record_count`, `manual_review_requested`, `result`, `error`, `cancelled`. | Avoid UI/controller churn and preserve the existing pause/resume manual-review flow. |
| Orchestration | Preserve `ProcessService` as the pipeline owner, but split internals into small private steps or helper services. | Proposal identifies service-layer rebuild as safest path; entry points remain rollback-friendly. |
| Excel boundary | `ExcelService` owns hidden Excel open/read/write/export only; callers never set visibility true. | Satisfies headless Excel requirement and isolates Windows COM behavior. |
| Block semantics | Use a `BlockPlan` concept: index, total, records, output path, image directory, progress base. | Makes 27-record splitting testable and failures isolated. |
| Image handling | Image export returns ordered slot results and per-slot failures; missing images become block warnings or block skips. | Allows later blocks to continue without crashing. |
| Word ownership | Generation saves/closes owned hidden Word documents before UI opens the review document; review opening is controller-owned. | Avoids duplicate COM ownership/opening conflicts caused by both service and controller owning the same file. |
| Messages | Logger receives traceback, paths, COM details, block/slot numbers; UI callbacks receive concise Spanish without stack traces. | Keeps diagnostics useful without technical noise. |

## Data Flow

```text
MainController -> WorkerClient -> worker_process -> ProcessService.run
    -> ValidationService: paths, filter, printer, sheets, headers, records
    -> ExcelService: hidden open, filter records, write block, export PNGs
    -> BlockPlanner: split records into <=27 item blocks
    -> WordService: copy template, replace <img1>...<img27>, save document
    -> ReviewCoordinator: emit manual_review_requested and wait response
    -> ProcessService: continue next block or cancel
```

Recoverable block failures are logged as block results, then the loop advances when later blocks can still be generated. Fatal preflight failures stop before document creation.

## File Changes

| File | Action | Description |
|---|---|---|
| `services/process_service.py` | Modify | Keep `run()`/`load_filters()` APIs; extract validation, block planning, image preparation, document generation, review wait, and recovery. |
| `services/excel_service.py` | Modify | Keep hidden COM ownership; expose clearer read/export methods and typed export failures. |
| `services/word_service.py` | Modify | Keep COM-first replacement and fallback; ensure placeholders are cleared and owned documents are saved/closed before review. |
| `services/validation_service.py` | Modify | Add template placeholder and image/block validation helpers if needed. |
| `services/worker_process.py` | Minimal/Unchanged | Preserve JSON event shapes; only map any new summary warnings through existing `log`/`status` events. |
| `controllers/main_controller.py` | Minimal/Unchanged | Keep manual review UI behavior; remain owner of opening the generated document for review. |
| `tests/` | Modify/Create | Add strict TDD coverage for block planning, hidden Excel, recoverable failures, placeholder handling, and worker contract stability. |

## Interfaces / Contracts

- `ProcessService.run(...) -> None` remains unchanged.
- `ManualAdjustCallback(int block_index, int total_blocks, str document_path, int baseline_mtime_ns) -> str` remains unchanged.
- Internal block result shape should capture `block_index`, `record_count`, `document_path`, `warnings`, and `failed` without changing worker JSON.
- Recoverable errors emit `log_callback` warnings and continue; fatal validation raises `ValidationError` with a friendly message.

## Testing Strategy

| Layer | What to Test | Approach |
|---|---|---|
| Unit | 27-item splitting, output naming, warning/error mapping, user vs logger messages. | `unittest` with fake services; write failing tests first. |
| Unit | Excel hidden open remains private and non-visible. | Extend `tests/test_excel_service_hidden.py`. |
| Unit | Word placeholder replacement clears unused `<imgN>` and fallback embeds images. | Extend `tests/test_word_service_no_com.py`. |
| Service | Missing image/export failure skips or warns at block level and continues later blocks. | Fake `ExcelService`/`WordService` in `ProcessService` tests. |
| Contract | Worker event names and manual-review request payload stay compatible. | Unit test `worker_process`/`worker_client` parsing with representative JSON. |
| Smoke | Real Excel/Word COM path remains hidden and review document opens once. | Manual smoke scripts after unit suite. |

## Migration / Rollout

No data migration required. Roll out as a service-layer refactor with unchanged UI and worker protocol, keeping rollback limited to `services/` and tests.

## Open Questions

- [ ] Should a block with one missing image produce a partial review document, or should the whole block be skipped while later blocks continue?
