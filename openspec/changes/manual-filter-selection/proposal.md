# Proposal: Manual Filter Selection for Pasted IDs

## Intent
- Add an additive mode for selecting pasted IDs across different filters without changing the current full-order flow.
- Preserve pasted order in the new selection list so users print in the same sequence they entered.

## Scope

### In Scope
- Manual selection of pasted IDs after record resolution, regardless of which filter each ID came from.
- Preserve pasted order in the resulting selection list.
- Reuse the existing preview/manual-print pipeline after resolution.
- Keep the current complete-order flow unchanged.

### Out of Scope
- Changes to Word layout, print formatting, or label templates.
- Reworking Excel import parsing or validation rules beyond what selection needs.
- Altering the existing complete-order UX or its defaults.

## Capabilities

### New Capabilities
- `manual-filter-selection`: resolve pasted IDs from mixed filters into a manually curated, order-preserving selection.

### Modified Capabilities
- None.

## Approach
- Add a new selection path in the UI/controller layer (`main.py`, `ui/main_window.py`) that branches only after IDs are resolved.
- Keep record lookup/validation in `services/excel_service.py` and `services/validation_service.py` unchanged unless needed to support mixed-filter resolution.
- Build the new selection list in pasted order, then pass it into the existing preview/manual-print flow used by `services/process_service.py`, `services/word_service.py`, and `services/print_service.py`.
- Preserve the current full-order path as the default behavior.

## Affected Areas

| Area | Impact | Description |
|---|---|---|
| `main.py` | Modified | Route the new selection mode without changing default flow |
| `ui/main_window.py` | Modified | Add/manual-select UX for pasted mixed-filter IDs |
| `services/process_service.py` | Modified | Accept order-preserving selection list |
| `services/excel_service.py` | Possibly modified | Resolve records for pasted IDs across filters |
| `services/validation_service.py` | Possibly modified | Validate mixed-filter selections |
| `services/word_service.py` / `services/print_service.py` | Unchanged or minimal | Reused downstream pipeline |

## Risks

| Risk | Likelihood | Mitigation |
|---|---|---|
| New mode accidentally alters the current full-order path | Low | Keep the branch additive and default to existing flow |
| Pasted order gets lost during record resolution | Medium | Carry source index through selection building |
| Mixed-filter validation surfaces unexpected edge cases | Medium | Limit changes to selection resolution, not print generation |

## Rollback Plan
- Disable the new selection entry point and retain the existing complete-order flow only.
- Revert UI/controller changes if order preservation or mixed-filter resolution causes regressions.

## Dependencies
- Existing preview/manual-print pipeline and record model resolution logic.

## Success Criteria
- [ ] Users can manually select pasted IDs coming from different filters.
- [ ] The new selection list preserves pasted order exactly.
- [ ] The current complete-order flow behaves exactly as before.
- [ ] The same preview/manual-print pipeline is reused after resolution.
