# Tasks: GeneradorEtiquetasSAP Document-Generation Rebuild

## Review Workload Forecast

| Field | Value |
|-------|-------|
| Estimated changed lines | 500-750 |
| 400-line budget risk | High |
| Chained PRs recommended | Yes |
| Suggested split | PR1 Excel ingest + validation -> PR2 block/image/Word -> PR3 resilience + verification |
| Delivery strategy | ask-on-risk |
| Chain strategy | pending |

Decision needed before apply: Yes
Chained PRs recommended: Yes
Chain strategy: pending
400-line budget risk: High

### Suggested Work Units

| Unit | Goal | Likely PR | Notes |
|------|------|-----------|-------|
| 1 | Hidden Excel ingest, validation, block plan | PR 1 | Base = feature/tracker; no UI/controller changes |
| 2 | Image export + Word composition/ownership | PR 2 | Base = PR 1; keep manual review flow intact |
| 3 | Continue-on-error logs + tests/smoke verification | PR 3 | Base = PR 2; finish contract coverage |

## Phase 1: RED tests and contract safety

- [ ] 1.1 Add failing tests in `tests/test_excel_service_hidden.py` for hidden Excel open/read behavior and no visible window.
- [x] 1.2 Add failing tests in `tests/test_process_service_word_com.py` for 27-record block planning and stable worker/controller contract.
- [x] 1.3 Add failing tests in `tests/test_word_service_no_com.py` for placeholder clearing, image insertion, and no duplicate open/ownership.

## Phase 2: Excel ingest, filtering, block splitting

- [ ] 2.1 Rebuild `services/excel_service.py` to open Excel via private `DispatchEx`, keep `Visible=False`, and disable alerts/events/screen updating.
- [x] 2.2 Add/adjust `services/validation_service.py` helpers to validate Excel path, template path, filter, required columns, and placeholder readiness.
- [ ] 2.3 Implement `BlockPlan` and `<=27` record splitting in `services/process_service.py` without changing UI flow or record selection rules.

## Phase 3: Image export, Word composition, review ownership

- [x] 3.1 Harden Excel image export in `services/excel_service.py` to return ordered slot results, log per-slot failures, and keep later slots/blocks usable.
- [x] 3.2 Rebuild `services/word_service.py` so `<img1>...<img27>` are replaced in order, unused placeholders are cleared, and formatting is preserved.
- [x] 3.3 Ensure `services/process_service.py` saves/closes owned block documents before manual review so the controller opens the review copy once.

## Phase 4: Continue-on-error logging and verification

- [x] 4.1 Thread block/slot/file/COM traceback details through `services/process_service.py` logs while keeping UI messages concise and non-technical.
- [x] 4.2 Preserve continue-on-error behavior for recoverable missing-image/block failures so later blocks still run when safe.
- [ ] 4.3 Run `python -m unittest discover -s tests`, then `python excel_smoke_test.py` and `python word_smoke_test.py` if COM is available.
