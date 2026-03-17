# LP Before vs After Validation

Compared:
- Before: `LP VERVE_Chapel Hill_20260304 JB Review.xlsm`
- After: `Copy of LP VERVE_Chapel Hill_20260304 JB Review Finished.xlsm`

## Passes
1. Waterfall tabs renamed to Sale tabs:
- `1-Yr Waterfall -> 1-Yr Sale`
- `3-Yr Waterfall -> 3-Yr Sale`
- `4-Yr Waterfall -> 4-Yr Sale`

2. Expected sheets removed:
- `FR Waterfall Analysis`
- `Returns Exhibit`

3. Zero formulas in output reference deleted tabs or old waterfall tab names.

4. Executive Summary links updated correctly:
- `J16='1-Yr Sale'!Q22`
- `J17='1-Yr Sale'!R22`
- `K16='3-Yr Sale'!Q22`
- `K17='3-Yr Sale'!R22`
- `L16='4-Yr Sale'!Q22`
- `L17='4-Yr Sale'!R22`

5. GP Fees and GP Return blocks are removed (labels no longer present on all 3 Sale tabs).

6. Structural move/clear behavior is present:
- `B7 = CASH FLOW SUMMARY`
- `B19 = Net Equity`

7. Row-340 safety preserved (`_WF_Period_lookup` still present):
- `B341 = Monthly`
- `C341 = 12`
- `A340` blank after clear-through-row-340.

8. Blue RGB text scan (`#0070C0` / `#0066FF`) found zero remaining cells on all 3 Sale tabs.

## Notable Difference
- Additional sheet exists in the finished file: `Claude Log`.
- This is not in the LP canonical step list, but it does not create formula-reference issues.

## Artifacts
- Detailed JSON: `variants/lp/lp-before-vs-after-diff.json`
