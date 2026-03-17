# VERVE Proforma Cleaner Workflow
Version: 1.0
Status: Parsed from user-provided Set 2 (Complete instruction set)
Date captured: 2026-03-05

## Goal
Automate full transformation of a raw VERVE proforma workbook into a clean deliverable through 42 ordered operations.

## Core Rules
- Runtime reads only (no hardcoded values)
- Deletion requires pre-check for zero references
- Conditional actions only when condition evaluates true
- White fill `#FFFFFF` as default blank state

## Steps
This file stores the complete 42-step Set 2 specification exactly as provided by the user in structured form (rename, text normalization, conditional row deletes/collapses, notes/comments removal, grouped row hiding, external-reference hardcoding, safe sheet deletions, clear-to-white operations, and final assumptions cleanup).

### Key reusable patterns
- `safeDeleteSheet(sheetName)` with pre-scan abort + post-delete verify
- `hardcodeRefsToSheet(targetSheetName)` with formulas+values co-load and zero-ref verification
- `clearToWhite(sheetName, range)` clear+white fill+all border removal

## Implementation Notes Captured
- Step order is authoritative (1..42)
- Build for dynamic layouts where row positions can shift
- Use text/display checks where required (`' - '`, `'$-'`)
- Handle sheets missing in some workbooks gracefully

## Set 2 source snapshot
- Source stored as user prompt on 2026-03-05
- Treated as authoritative instruction variant #2 for same workflow family
