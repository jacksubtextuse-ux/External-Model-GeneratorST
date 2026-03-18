# VERVE Proforma Cleaner Workflow
Version: 1.0
Status: Parsed from user-provided Set 1 (Claude Excel)
Date captured: 2026-03-05

## Goal
Transform a raw VERVE `.xlsm` proforma workbook into an investor-ready deliverable through 42 ordered steps with dependency-safe hardcoding/deletion behavior.

## Input Contract
- Workbook format: `.xlsm`
- Required sheets: Executive Summary, Development Summary, Cash Flow, Assumptions, Development, Sale Proceeds, Returns Exhibit
- Optional sheets may exist and are eventually deleted.
- `Executive Summary!E6` contains full address; city parsed from second comma-separated segment.

## Output Contract
- Filename: `VERVE_{city}_{MM-DD-YYYY}.xlsm`
- Optional tabs removed per deletion sequence
- Blue text normalized to black (white preserved)
- References to deleted sheets hardcoded before deletion
- No broken references

## Critical Rules (captured)
1. Never hardcode fixed values; always runtime read/write.
2. Load formulas + values in same sync before overwriting formulas.
3. Never delete a sheet before zero-reference safety check.
4. Use `.text` for formatted dash checks.
5. Use bulk usedRange scans, not cell-by-cell.
6. Cleared ranges must end with white fill `#FFFFFF`.
7. Respect strict order for Steps 22–36.
8. Missing sheets must be skipped gracefully.
9. Re-evaluate rows dynamically after row deletions.
10. Formula-reference updates occur after shifts.

## Utility Functions Required
- `safeDeleteSheet(context, sheetName)`
- `hardcodeRefsToSheet(context, targetSheetName)`
- `clearToWhite(context, sheetName, rangeAddress)`

## Phase Plan
- Phase 1: Executive Summary (1–8)
- Phase 2: Cash Flow (9–10)
- Phase 3: Assumptions (11–15, 39–42)
- Phase 4: Development (16–19)
- Phase 5: Sale Proceeds (20–21)
- Phase 6: Deletions with hardcode safety (22–36)
- Phase 7: Final cleanup (37–42)

## Full Step List Reference
- Steps 1–42 captured from user-provided Set 1 and treated as authoritative for v1.0 implementation.

## Error Handling Contract
- Missing sheet: log + skip
- Remaining refs before deletion: abort with descriptive details
- Hardcode verification fails: abort with count
- DSCR < 1.25: highlight + notify, continue
- Condition not met: skip silently + log
- Timeout: reduce batch size and retry

## Post-Execution Validation
- Filename correct
- Target sheets color/formats normalized
- Deleted sheets absent
- No formulas referencing removed tabs
- White-cleared ranges with no borders
- Group/hidden rows collapsed
- DSCR result logged
