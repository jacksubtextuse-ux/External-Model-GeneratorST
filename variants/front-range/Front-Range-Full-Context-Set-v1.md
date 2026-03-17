# Front Range Variant — Full Context Set (Isolated)
Date captured: 2026-03-06
Status: Saved separately from base model workflow and from Front Range base-only set

## Source
User-provided full/context instruction set for Front Range add-on variant.

## Step Package Captured

### Step 1: Change Non-White Text To Black (Preserve White)
Target sheets:
- 1-Yr Waterfall
- 3-Yr Waterfall
- 4-Yr Waterfall

Rules:
- Keep white text (#FFFFFF) unchanged.
- Convert blue (#0070C0) and any non-white text to black (#000000).
- Preserve intentional white header text in dark-blue sections.
- Prefer `getCellProperties` batch read + batched writes for performance.

Known white-text preserve regions called out:
- B7:F7, L7, P7:R7, L14, B19:I19, L19:R19, J9:J10, J19, P5 (all 3 waterfall sheets)

---

### Step 2: Remove GP Fees Section (All 3 Waterfall Sheets)
Per sheet:
- Find exact `GP Fees` dynamically (no hardcoded row).
- Build block: `L{gpFeesRow}:R{gpFeesRow+3}`.
- Verify labels in column L rows +0..+3:
  - GP Fees
  - Development Fee
  - Construction Management Fee
  - GP Total Return
- Scan workbook for external refs to block cells and hardcode those refs first.
- Clear block contents.
- Remove all 6 border sides.
- Set fill to white.

Expected example results noted:
- 1-Yr: L30:R33
- 3-Yr: L28:R31
- 4-Yr: L28:R31

---

### Step 3: Remove GP Return on Equity / Co-GP Block (1-Yr Only)
- Find exact `GP Return on Equity` dynamically.
- Block: `S{startRow}:T{startRow+8}` (9 rows).
- Verify label offsets in column T:
  - +0 GP Return on Equity
  - +1 GP Promote
  - +2 Total GP Return
  - +4 Co-GP Split
  - +6 Co-GP Total Return
  - +7 Subtext Total Return
  - +8 Total
- Scan workbook for external refs to S/T block and hardcode ref cells first.
- Clear block contents.
- Remove all 6 border sides.
- Set fill white.
- Restore thick black left border on S{startRow} and S{startRow+1}.

---

### Step 4: Remove GP Return on Equity / Co-GP Block (3-Yr Only)
- Same as Step 3 but block is 10 rows:
  - `S{startRow}:T{startRow+9}`
- Label offsets in column T:
  - +0 GP Return on Equity
  - +1 GP Promote
  - +2 Total GP Return
  - +5 Co-GP Split
  - +7 Co-GP Total Return
  - +8 Subtext Total Return
  - +9 Total
- Same reference hardcode + clear + border restore actions.

---

### Step 5: Remove GP Return on Equity / Co-GP Block (4-Yr Only)
- Identical to Step 4, sheet name is `4-Yr Waterfall`.
- Same 10-row block, same label offsets, same safety checks/actions.

---

## Consolidation Notes (as-provided context)
- Steps 4 and 5 can be implemented as one loop over:
  - ["3-Yr Waterfall", "4-Yr Waterfall"]
- Preserve dynamic search behavior; avoid fixed row assumptions.
- Always hardcode external references before clearing any target block.

## Isolation Contract
- Stored under variants/front-range only.
- Not merged into base model output logic.
- Intended to be wired later as the Front Range selectable option in the 4-option system.
