# Front Range Variant — Canonical Reconciliation
Version: 1.0
Date: 2026-03-06
Status: Linked + reconciled from base-only and full-context sets

## Linked Sources
- Base-only set:
  - `Front-Range-Base-Set-v1.md`
- Full-context set:
  - `Front-Range-Full-Context-Set-v1.md`

These two files describe the same Front Range instruction family.

## Reconciliation Result
- ✅ Step structure matches across both sets (5 steps).
- ✅ Dynamic-search and safety-check logic matches.
- ✅ Block ranges and label offsets match.
- ✅ External-reference hardcode-before-clear behavior matches.
- ✅ Border/fill cleanup behavior matches.

## Precedence Rules
1. Use the **base-only set** as concise workflow skeleton.
2. Use the **full-context set** for implementation details (performance notes, explicit preserve lists, expected row examples).
3. If wording differs, prefer the stricter/clearer rule from full-context.

## Canonical Front Range Steps
1. Change non-white text to black (`1-Yr`, `3-Yr`, `4-Yr` Waterfall), preserving white text.
2. Remove GP Fees block (`L:R`, 4 rows) on all 3 waterfall sheets with:
- dynamic label find,
- label verification,
- external reference hardcoding,
- clear contents, remove all borders, white fill.
3. Remove GP Return on Equity block on `1-Yr Waterfall` (`S:T`, 9 rows) with offset labels + reference hardcode + clear + thick left border restore on top two rows.
4. Remove GP Return on Equity block on `3-Yr Waterfall` (`S:T`, 10 rows, shifted offsets) with same safety/clear/border behavior.
5. Remove GP Return on Equity block on `4-Yr Waterfall` (same as step 4).

## Canonical Details Locked
- White text (`#FFFFFF`) is preserved in Step 1.
- Border reset includes all 6 sides (`edgeTop`, `edgeBottom`, `edgeLeft`, `edgeRight`, `insideHorizontal`, `insideVertical`).
- Left-border restore for Steps 3/4/5 is **thick solid black** on `S{startRow}` and `S{startRow+1}`.
- Steps 4 and 5 may be implemented as one loop over `['3-Yr Waterfall','4-Yr Waterfall']`.

## Integration Guardrail
This variant remains isolated under `variants/front-range` and must only run when user selects the Front Range option.
