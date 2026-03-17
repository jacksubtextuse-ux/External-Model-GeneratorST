# Workbook Diff Summary

- Original: `C:\Users\JackBranding\OneDrive - Subtext\Desktop\Codex\workflows\verve-proforma-cleaner\samples\original.xlsm`
- Reviewed: `C:\Users\JackBranding\OneDrive - Subtext\Desktop\Codex\workflows\verve-proforma-cleaner\samples\reviewed.xlsm`

## Sheet-level
- Removed sheets: 11
  - 3-Yr Co-GP Waterfall
  - Building Program
  - Change Log
  - Construction Pricing
  - Historic OpEx Comparison
  - Proforma Comparison
  - Reassessed 1-Yr Waterfall
  - Reassessed 3-Yr Waterfall
  - Reassessed 4-Yr Waterfall
  - TermSheet
  - TermSheet-Ignore
- Added sheets: 0
- Expected deleted tabs still present: none
- Unexpected removed tabs: none

## Remaining Formula Refs To Deleted Tabs
- none

## Spot Checks
- Executive Summary!E5: `VERVE Chapel Hill` -> `VERVE Chapel Hill` (formula: False -> False)
- Executive Summary!J15: `=Q5` -> `0.0515` (formula: True -> False)
- Executive Summary!K15: `=J15+Q6+Q6` -> `0.0535` (formula: True -> False)
- Executive Summary!L15: `=K15+Q6` -> `0.0545` (formula: True -> False)
- Executive Summary!E60: `=+'3-Yr Waterfall'!D64` -> `=+'3-Yr Waterfall'!D64` (formula: True -> True)
- Cash Flow!Q46: `=(Q$45+Q$30)/('Development Summary'!$I$36-'Construction Pricing'!$H$27)` -> `0.0757556875351677` (formula: True -> False)
- Assumptions!I43: `=+'Building Program'!F31` -> `1824` (formula: True -> False)
- Assumptions!X22: `=ROUND('Building Program'!M27,0)` -> `629` (formula: True -> False)
- Assumptions!D21: `=+$X$26` -> `=+$X$26` (formula: True -> True)
- Development!H11: `=+'Construction Pricing'!C16` -> `502832` (formula: True -> False)
- Development!H78: `='Construction Pricing'!I26` -> `3680270` (formula: True -> False)
- Sale Proceeds!B75: `Sale Analysis with Tax Reassessment` -> `None` (formula: False -> False)
- Returns Exhibit!J3: `Return Summary with Tax Reassessment` -> `None` (formula: False -> False)

## Sheet Change Counts (formula/literal)
- 1-Yr Waterfall: formulas=3, literals=1
- 3-Yr Waterfall: formulas=3, literals=1
- 4-Yr Waterfall: formulas=3, literals=1
- Assumptions: formulas=2, literals=12
- Cash Flow: formulas=16, literals=6
- Development: formulas=28, literals=6
- Executive Summary: formulas=52, literals=21
- Returns Exhibit: formulas=42, literals=37
- Sale Proceeds: formulas=62, literals=32