# LP Variant Canonical Spec (v1)

## Scope
This LP set is applied **after Base + Front Range** logic. It modifies waterfall presentation and linkages, then performs safe sheet cleanup.

Target waterfall sheets before rename:
- `1-Yr Waterfall`
- `3-Yr Waterfall`
- `4-Yr Waterfall`

## Reconciled Rules
1. Use `moveTo()` for true cut/paste moves (Step 7). Do not use `copyFrom()` there.
2. `copyFrom()` is allowed for Step 2 where we intentionally copy style+formula row content.
3. Step 9 must use `clearContents` (not `clear(all)`), and must never clear below row 340.
4. Preserve `_WF_Period_lookup` at `B341:E343` on each waterfall sheet.
5. All destructive operations require pre-checks and abort/skip behavior.
6. Use dynamic finding where possible (`findAllOrNullObject` / dynamic scans), avoid hardcoded row assumptions.

## Resolved Conflicts Across The 3 Source Files
- Sheet naming conflict (`Waterfall` vs `Sale`):
  - Steps 1-10 operate on `*Waterfall` names.
  - Step 11 performs rename to `*Sale`.
  - Steps 12-13 are workbook-wide safe deletions.
- Step 1 range conflict on 3-Yr/4-Yr:
  - Canonical behavior: clear `L:R` for 5 rows below the located `Project Level` row (`+1` to `+5`), then restore top border on first cleared row.
- Step 9 duplicated variants:
  - Canonical behavior: clear from `SUMMARY DISTRIBUTIONS` row through row `340` only, then set white fill and remove borders.

## Canonical Step Order
### Step 1
Find `Project Level` in each waterfall sheet (column L context), clear sub-block `L:R` from `row+1` through `row+5`, set white fill, remove borders, restore thin top border on first cleared row.

### Step 2
Copy `B:D` row for `Equity Multiple` from unleveraged block (sequence test: `Unleveraged Cash Flow -> IRR -> Equity Multiple`) into leveraged blank slot (sequence test: `IRR -> blank -> DSCR`).
Then rewrite destination `D` formula to leverage `Net Cash Flow` row dynamically.

### Step 3
Set return-summary Profit cell (`P`, 2 rows below `Profit` header in col P) to linked source `=E{NetCashFlowRow}`.

### Step 4
Set return-summary IRR cell (`Q`, 2 rows below `IRR` header in col Q) to `=D{LeveragedIRRRow}` found from col B `IRR` row.

### Step 5
Set return-summary Equity Multiple cell (`R`, 2 rows below `Equity Multiple` header in col R) to `=D{LeveragedEMRow}` found from col B `Equity Multiple` row.

### Step 6
Clear waterfall terms block `B7:J18` after safety check `B19` contains `CASH FLOW SUMMARY`. Set white fill, remove borders.

### Step 7
Move `B19:J39` to `B7` using `moveTo()` after checks:
- `B19` contains `CASH FLOW SUMMARY`
- `B7` is blank

### Step 8
Delete empty rows `29:45` with row-by-row emptiness verification across used columns.
If mixed, delete only confirmed-empty rows (bottom-up).

### Step 9
Find `SUMMARY DISTRIBUTIONS` in column A and clear contents from that row to row `340` (inclusive). Then fill white and remove borders.
Do not touch row `341+`.

### Step 10
Link Executive Summary to waterfall return-summary cells:
- `J16='1-Yr Waterfall'!Q22`, `J17='1-Yr Waterfall'!R22`
- `K16='3-Yr Waterfall'!Q22`, `K17='3-Yr Waterfall'!R22`
- `L16='4-Yr Waterfall'!Q22`, `L17='4-Yr Waterfall'!R22`

### Step 11
Rename tabs:
- `1-Yr Waterfall -> 1-Yr Sale`
- `3-Yr Waterfall -> 3-Yr Sale`
- `4-Yr Waterfall -> 4-Yr Sale`

### Step 12
Safely delete `FR Waterfall Analysis` only if zero formula references remain.

### Step 13
Safely delete `Returns Exhibit` only if zero formula references remain.

## Safety Gates (Mandatory)
- Any failed pre-check should skip/abort that specific step with explicit logging.
- Never perform sheet deletion when references remain.
- Never clear below row 340 on waterfall sheets in Step 9.

## Notes For Implementation
- Maintain detailed run logging per step (status + affected ranges/cells).
- Keep this LP set isolated as its own selectable option layered on top of Front Range.
