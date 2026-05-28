# Proposal: GeneradorEtiquetasSAP Document-Generation Rebuild

## Intent
- Rebuild the unstable Excel/Word generation pipeline while preserving the existing UI, buttons, filters, window behavior, and manual-review flow.
- Remove visible Excel-window dependence and harden Word placeholder replacement so the app can continue on recoverable errors instead of crashing.

## Scope

### In Scope
- Isolate/rework the service-layer pipeline centered in `services/process_service.py` and related services.
- Read Excel in the background with no visible Excel window.
- Replace `<img1>...<img27>` reliably in Word, preserve formatting, and keep manual review working.
- Add strong logging and continue-on-error behavior for non-fatal document-generation failures.
- Keep the worker/controller contract stable if possible.

### Out of Scope
- UI redesign, button changes, filter behavior changes, or window-flow changes.
- Template redesign, business-rule changes, or record-selection changes.
- Changing the print/review user journey beyond making generation safer.

## Capabilities

### New Capabilities
- `document-generation-pipeline`: resilient Excel-to-Word document generation with hidden Excel access, placeholder image replacement, and manual-review-safe continuation.

### Modified Capabilities
- None.

## Approach
- Treat `services/process_service.py` as the orchestration boundary and rebuild only the generation internals behind its existing flow.
- Keep controller/worker message shapes stable; only extend them if a compatibility-preserving change is unavoidable.
- Move fragile Excel/Word COM steps behind explicit logging, validation, and per-block recovery paths.

## Affected Areas

| Area | Impact | Description |
|---|---|---|
| `services/process_service.py` | Modified | Rebuild orchestration and error-recovery around document generation |
| `services/excel_service.py` | Modified | Headless Excel reading/export path |
| `services/word_service.py` | Modified | Robust placeholder replacement and formatting-preserving insertion |
| `services/worker_process.py` | Minimal/Unchanged | Keep worker protocol stable |
| `controllers/main_controller.py` | Minimal/Unchanged | Preserve UI/controller flow |

## Risks

| Risk | Likelihood | Mitigation |
|---|---|---|
| COM automation regressions on Windows | High | Keep changes isolated, validate each block, preserve fallback paths |
| Placeholder/image alignment breaks formatting | Medium | Validate each slot and keep formatting-preserving insertion logic |
| Manual review flow becomes brittle | Medium | Preserve pause/resume contract and avoid UI changes |

## Rollback Plan
- Revert the service-layer rebuild and restore the current pipeline entry points.
- Keep the UI and worker/controller contract intact so rollback does not require user-facing changes.

## Dependencies
- Existing Windows COM automation stack (`pywin32`, Excel, Word) and current label templates.

## Success Criteria
- [ ] Excel is read without showing a visible Excel window.
- [ ] Word generation replaces `<img1>...<img27>` reliably and preserves formatting.
- [ ] Manual review continues to work without crashes.
- [ ] Recoverable errors continue with clear logs instead of stopping the whole run.
- [ ] UI, buttons, filters, and window behavior remain unchanged.
