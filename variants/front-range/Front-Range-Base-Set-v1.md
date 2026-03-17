# Front Range Variant — Base Code Set (Isolated)
Date captured: 2026-03-06
Status: Saved separately from base model workflow

## Source
User-provided base instruction/code set for Front Range add-on variant.

## Captured Actions (from provided set)
1. Change non-white text to black on these sheets:
- 1-Yr Waterfall
- 3-Yr Waterfall
- 4-Yr Waterfall

2. For each of the 3 waterfall sheets, find GP Fees block dynamically:
- Locate exact label: "GP Fees"
- Target block: L{row}:R{row+3}
- Verify labels in column L rows:
  - GP Fees
  - Development Fee
  - Construction Management Fee
  - GP Total Return
- Scan workbook for formulas referencing this block.
- Hardcode those external reference cells before clearing.
- Clear block contents, remove all 6 border sides, fill white.

3. 1-Yr Waterfall — remove GP Return on Equity / Co-GP block:
- Find label: "GP Return on Equity"
- Target block: S{startRow}:T{startRow+8}
- Verify expected labels at offsets:
  - 0 GP Return on Equity
  - 1 GP Promote
  - 2 Total GP Return
  - 4 Co-GP Split
  - 6 Co-GP Total Return
  - 7 Subtext Total Return
  - 8 Total
- Scan workbook for external refs to S/T cells in the block.
- Hardcode those ref cells.
- Clear block contents, remove all 6 borders, set fill white.
- Restore thick black left border on S{startRow} and S{startRow+1}.

4. 3-Yr Waterfall — same as #3 with 10-row block and shifted label offsets:
- Target block: S{startRow}:T{startRow+9}
- Label offsets:
  - 0 GP Return on Equity
  - 1 GP Promote
  - 2 Total GP Return
  - 5 Co-GP Split
  - 7 Co-GP Total Return
  - 8 Subtext Total Return
  - 9 Total
- Same reference hardcode + clear + border restore behavior.

5. 4-Yr Waterfall — same as #4:
- Target block: S{startRow}:T{startRow+9}
- Same label checks and offsets as 3-Yr.
- Same reference hardcode + clear + border restore behavior.

## Isolation Contract
- This file is stored under variants/front-range and is not merged into base model instructions.
- Integration into the 4-option UI will be done later via explicit option selection.
