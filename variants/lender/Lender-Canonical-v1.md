# Lender Variant Canonical Spec (v1)

## Scope
This Lender set is applied **after Base + Front Range + LP**.

Input state expected from LP output:
- `Development Summary`
- `Cash Flow`
- `Assumptions`
- plus non-keeper tabs (to be removed in Step 2)

Target output state:
- Only `Development Summary`, `Cash Flow`, `Assumptions`
- Fully static (no live formulas)
- No references to deleted tabs
- Cash Flow and Assumptions delivery cleanup applied

## Reconciled Rules
1. Step 1 hardcode must use **matrix write** (`usedRange.values = newValues`) and never per-cell writes.
2. Load `formulas` and `values` together in one sync before replacement.
3. Clear operations must use **`clear(contents)`**, not clear-all.
4. Row deletions must run in **descending order**.
5. Column deletions must use **`DeleteShiftDirection.left`**.
6. Fill-color detection must use `getCellProperties` scans, not hardcoded row addresses.
7. `LIGHT_BLUE_OCCURRENCES_NEEDED` is required for table-depth boundaries:
   - `2` for unit mix table operations (Steps 7 and 9)
   - `1` for commercial table operation (Step 8)
8. Verify after each destructive operation.

## Canonical Step Order
### Step 1
Hardcode all formulas in `Development Summary`, `Cash Flow`, `Assumptions` via matrix write.
- Replace formula cells with computed values.
- If computed value is an Excel error string (`#...`), write `0` and log.

### Step 2
Safety test then delete all non-keeper tabs.
- Keep only: `Development Summary`, `Cash Flow`, `Assumptions`.
- Test A: zero live formulas in keeper sheets.
- Test B: zero references from keeper sheets to tabs being deleted.
- Abort if either test fails.

### Step 3
Cash Flow: conditionally delete rows by labels `Commercial Parking`, `Ground Lease`, `Tax Abatement`.
- Check column `F` on each matched row.
- Delete when value is `0`, blank, null, or `-`; keep otherwise.
- Delete in descending row order.

### Step 4
Cash Flow: find `NET OPERATING INCOME (LESS RESERVES)` and clear all rows below through last used row.
- Preserve the anchor row itself.
- Clear contents + white fill + remove borders.

### Step 5
Assumptions: find first `#002060` fill in column `B` below row 45.
- Clear from `B{foundRow}` through `{lastCol}{lastRow}`.
- Clear contents + white fill + remove borders.

### Step 6
Assumptions: find `ASSUMPTIONS` in column `C`.
- Measure contiguous `#002060` span to the right on that row.
- Clear from anchor row through last row, only across that measured span.
- Clear includes the blue header row itself.

### Step 7
Assumptions: delete columns for labels `Current Year Per Bed Input` and `Current Year Rent PSF`.
- Both labels must be in the same row.
- Top boundary: nearest `#002060` row above.
- Bottom boundary: second `#DCE6F1` occurrence (`LIGHT_BLUE_OCCURRENCES_NEEDED = 2`).
- Additional safety checks: two blank rows after boundary and commercial block integrity below.
- Delete range with shift-left.

### Step 8
Assumptions: delete `Current Rent/Yr` column block.
- Same boundary logic as Step 7, but single label and single column.
- Bottom boundary: first `#DCE6F1` occurrence (`LIGHT_BLUE_OCCURRENCES_NEEDED = 1`).
- Delete range with shift-left.

### Step 9
Assumptions: clear everything right of `Total SF` within unit mix table boundary.
- Preserve `Total SF` column and all columns left of it.
- Row boundary uses unit mix depth (`LIGHT_BLUE_OCCURRENCES_NEEDED = 2`).
- Clear contents + white fill + remove borders.

## Reconciliation Notes Across `Code.txt`, `Complete.txt`, `Full.txt`
- All three sources align on 9 steps and sequencing.
- No substantive rule conflicts found.
- Context files add the same safety rules seen in code:
  - matrix hardcode
  - clear-contents only
  - descending deletes
  - color-driven dynamic boundaries

## Implementation Notes
- Keep this Lender set isolated as its own selectable option layered on top of LP.
- Do not run Lender steps directly on Base/Front Range-only outputs.
