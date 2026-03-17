# Front Range Before/After Validation

- Before: `C:\Users\JackBranding\OneDrive - Subtext\Desktop\Codex\workflows\verve-proforma-cleaner\variants\front-range\samples\base-before.xlsx`
- After: `C:\Users\JackBranding\OneDrive - Subtext\Desktop\Codex\workflows\verve-proforma-cleaner\variants\front-range\samples\front-range-after.xlsm`

## Derived Block Addresses (from before file)
- 1-Yr Waterfall: GP Fees `L30:R33`, GP Return `S26:T34`
- 3-Yr Waterfall: GP Fees `L28:R31`, GP Return `S26:T35`
- 4-Yr Waterfall: GP Fees `L28:R31`, GP Return `S26:T35`

## Removal Checks (after file)
- 1-Yr Waterfall: gp_fees_removed=True, gp_return_removed=True
- 3-Yr Waterfall: gp_fees_removed=True, gp_return_removed=True
- 4-Yr Waterfall: gp_fees_removed=True, gp_return_removed=True

## Post-Reference Checks (after file)
- 1-Yr Waterfall: remaining_external_refs=0 (tracked_cells=46)
- 3-Yr Waterfall: remaining_external_refs=0 (tracked_cells=48)
- 4-Yr Waterfall: remaining_external_refs=0 (tracked_cells=48)

## Notes
- Font-color normalization is validated functionally (spec behavior) but exact hex counting is unreliable in openpyxl because theme/indexed colors are not always materialized as RGB.