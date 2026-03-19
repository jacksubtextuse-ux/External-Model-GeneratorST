# VERVE Proforma Cleaner — Canonical Workflow Spec
Version: 1.1
Date: 2026-03-05
Status: Canonical merge of Set 1, Set 2, Set 3 + validated workbook diff evidence

## Scope
Transforms a raw VERVE `.xlsm` proforma workbook into a clean, investor-ready workbook via a deterministic 42-step pipeline.

## Source Precedence
1. Set 2 is the base executable structure (full 42-step JSON semantics).
2. Set 1 adds phase/dependency constraints and utility function details.
3. Set 3 adds execution-proven outcomes and edge cases (real workbook behavior).
4. Real workbook diff is the acceptance baseline for this test case.

## Non-Negotiable Rules
1. Never hardcode business values in code; always read live runtime values and write them back.
2. For hardcoding external references, load `formulas` and `values` together before overwriting.
3. For deletions, enforce: hardcode refs -> verify zero refs -> delete -> verify deletion.
4. Use `.text` for formatted display checks (`' - '`, `'$-'`) where specified.
5. Use bulk usedRange operations and batched sync patterns; avoid cell-by-cell API sync loops.
6. White (`#FFFFFF`) is the required blank-state fill after clearing ranges.
7. Missing optional sheets must be logged and skipped gracefully.
8. Row-shift-sensitive logic must re-evaluate positions dynamically after deletions.

## Canonical Utilities
- `safeDeleteSheet(context, sheetName)`
- `hardcodeRefsToSheet(context, targetSheetName)`
- `clearToWhite(context, sheetName, rangeAddress)`

## Canonical Phase Order
- Phase 1: Steps 1-8 (Executive Summary)
- Phase 2: Steps 9-10 (Cash Flow)
- Phase 3: Steps 11-15 (Assumptions pre-delete)
- Phase 4: Steps 16-19 (Development)
- Phase 5: Steps 20-21 (Sale Proceeds)
- Phase 6: Steps 22-36 (hardcode/delete sequence; strict)
- Phase 7: Steps 37-42 (final cleanup)

## Step Map (Canonical)
1. Rename file from `Executive Summary!E6` city parse to `VERVE_{city}_{MM-DD-YYYY}.xlsm`.
2. Executive Summary blue font -> black (exclude white).
3. Set `Executive Summary!E5` to `VERVE {city}`.
4. Hardcode `Executive Summary!J15: L15` values.
5. DSCR check on `Executive Summary!E60`; warn+yellow if `< 1.25`.
6. Conditionally delete zero OpEx rows (Ground Lease / Tax Abatement) with dynamic row handling.
7. Clear `Executive Summary!N1:U12`; preserve left border on `N1:N12`.
8. Conditional ROC row collapse logic for `Executive Summary` when `L6 == L7`; shift/update formulas/formats. Formatting lock: `K5` must be bold + underlined + right-aligned, `G7` bold + Cambria + left-aligned, `L7` bold + Cambria + centered + `0 "bps"`.
9. Cash Flow: group/hide candidate zero rows (14, 30, 41) via numeric checks `F:Q`.
10. Cash Flow: delete ROC net tax-abatement row when equality condition is met.
11. Assumptions blue font -> black (full scan; verification pass).
12. Remove all notes/comments on Assumptions.
13. Group/hide Assumptions dash rows in `B49:V285` using `.text` on `J:V` and label-presence guard.
14. Assumptions: locate Retail/TBD/formula pattern in `G40:I50`, hardcode Commercial SF cell.
15. Assumptions: hardcode refs to Building Program/Construction Pricing; verify zero remain.
16. Development: remove notes/comments.
17. Development: non-black/non-white text -> black.
18. Development: hardcode refs to Building Program/Construction Pricing; verify zero remain.
19. Development: group/hide unused line items using label + `H==0` + `K.text` dash condition.
20. Sale Proceeds: group/hide commercial block rows `24:30` on zero/`$-` condition.
21. Sale Proceeds: clear `B75:J112` to white and borderless state.
22. Hardcode workbook refs to Building Program.
23. Delete Building Program with zero-ref safety precheck.
24. Hardcode workbook refs to Construction Pricing.
25. Delete Construction Pricing.
26. Hardcode workbook refs to `3-Yr Co-GP Waterfall`.
27. Delete `3-Yr Co-GP Waterfall`.
28. Clear `Returns Exhibit!J3:O37` to white and borderless state.
29. Delete `Reassessed 1-Yr Waterfall`.
30. Delete `Reassessed 3-Yr Waterfall`.
31. Delete `Reassessed 4-Yr Waterfall`.
32. Delete `Historic OpEx Comparison`.
33. Delete `Change Log`.
34. Delete `TermSheet`.
35. Delete `TermSheet-Ignore`.
36. Delete `Proforma Comparison`.
37. Clear `Development!L21:O31` to white and borderless state.
38. In Development/Assumptions, clear any cell whose formula references the `Cash Flow` tab (any cell); remove borders; if fill is yellow then reset it to white.
39. Assumptions: reset non-approved fill colors to white (approved: `#002060`, `#DCE6F1`, `#FFFFFF`).
40. Assumptions: find `Residential Parking Spaces`, hardcode adjacent right cell value.
41. Assumptions: copy `W22:X26` -> `AC22:AD26` with full formatting (`copyFrom all`).
42. Assumptions: clear `W4:AA44` to white and borderless state.

## Canonical Error Handling
- Missing sheet: warn + skip.
- Ref(s) remain before deletion: abort with sheet/cell evidence.
- Hardcode verification fails: abort with remaining count.
- DSCR threshold fail: continue with warning/highlight.
- Condition not met: no action, log.
- Timeout risk: reduce batch size and retry.

## Test-Case Evidence Notes (from reviewed workbook)
- All 11 target deletion sheets are absent in reviewed output.
- Zero formula references remain to deleted sheets.
- Spot hardcodes align for major cells (J15/K15/L15, I43, X22, H11..H78, Q46).
- Observed exception for this test case: `Assumptions!D21` remains formula (`=+$X$26`) despite Step 40 intent.

## Output Acceptance
- Filename pattern satisfied.
- Required sheets retained and optional deletion targets removed.
- No deleted-sheet references remain.
- Cleared ranges are white and borderless (except explicit preserved border case).
- Group/hide outcomes applied where conditions met.
- DSCR check result logged.

