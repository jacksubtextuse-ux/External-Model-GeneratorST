# Front Range Variant — Complete Consolidated Set (Authoritative)
Date captured: 2026-03-06
Status: Latest authoritative Front Range instruction set

## Source
User-provided complete instruction set titled:
"VERVE Chapel Hill — Waterfall Tab Conversion: Complete Instruction Set for Claude Code"

## Workflow Overview
Transforms waterfall tabs with 5 sequential steps and sheet-level safety checks.

Target sheets:
- 1-Yr Waterfall
- 3-Yr Waterfall
- 4-Yr Waterfall

---

## Step 1 — Change Non-White Text to Black (Preserve White)
For each target sheet:
- Read used-range font colors.
- Preserve white (`#FFFFFF`).
- Preserve already black (`#000000`).
- Convert all other colors (including blue `#0070C0`) to black (`#000000`).
- Prefer batched `getCellProperties` read and batched write updates.

Expected example output:
- ~18 changed cells per target sheet, with many white cells skipped.

---

## Step 2 — Remove GP Fees Block (All 3 Sheets)
Per sheet:
1. Find exact label `GP Fees` dynamically.
2. Build block: `L{gpFeesRow}:R{gpFeesRow+3}`.
3. Verify labels in column L (+0..+3):
   - GP Fees
   - Development Fee
   - Construction Management Fee
   - GP Total Return
4. Scan workbook formulas for external references to block cells.
5. Hardcode external ref cells before clearing.
6. Clear block contents, remove all 6 borders, set fill white.

Expected example clear ranges:
- 1-Yr: `L30:R33`
- 3-Yr: `L28:R31`
- 4-Yr: `L28:R31`

---

## Step 3 — Remove GP Return/Co-GP Block (1-Yr)
- Find `GP Return on Equity` dynamically.
- Block: `S{startRow}:T{startRow+8}` (9 rows).
- Verify label offsets in column T:
  - 0 GP Return on Equity
  - 1 GP Promote
  - 2 Total GP Return
  - 4 Co-GP Split
  - 6 Co-GP Total Return
  - 7 Subtext Total Return
  - 8 Total
- Scan workbook for external refs to S/T block.
- Hardcode external refs.
- Clear contents, remove all 6 borders, white fill.
- Restore thick black left border at `S{startRow}` and `S{startRow+1}`.

Expected: `S26:T34` on 1-Yr.

---

## Step 4 — Remove GP Return/Co-GP Block (3-Yr)
- Use shared helper `removeGPReturnBlock`.
- Block: `S{startRow}:T{startRow+9}` (10 rows).
- Label offsets:
  - 0 GP Return on Equity
  - 1 GP Promote
  - 2 Total GP Return
  - 5 Co-GP Split
  - 7 Co-GP Total Return
  - 8 Subtext Total Return
  - 9 Total
- Same ref hardcode, clear, and border restore behavior.

Expected: `S26:T35` on 3-Yr.

---

## Step 5 — Remove GP Return/Co-GP Block (4-Yr)
- Same logic as Step 4, sheet=`4-Yr Waterfall`.
- Same 10-row block and label offsets.

Expected: `S26:T35` on 4-Yr.

---

## Safety Rules (Locked)
- Dynamic find required; do not hardcode source row numbers.
- Label verification must pass before clearing each block.
- External references must be hardcoded before block clear.
- For Steps 3–5, left border restoration on S top two rows must be `thick` black.

---

## Notes
- Border style telemetry can sometimes report medium while visually thick; enforce `Excel.BorderLineStyle.thick` for Steps 3–5 restore.
- Steps 3–5 are implemented with one helper function; only sheet, row-count, and label map differ.

## Isolation Contract
- Stored under `variants/front-range` only.
- Does not modify base model instructions.
- Used for Front Range option integration.
