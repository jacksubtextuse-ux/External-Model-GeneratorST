from __future__ import annotations

import datetime as dt
import os
import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

from app.engine import WorkflowError


@dataclass
class RunLog:
    messages: list[str] = field(default_factory=list)

    def add(self, message: str) -> None:
        self.messages.append(message)


class VerveWorkflowRunnerCom:
    """Excel COM implementation to preserve workbook internals (drawings/external links)."""

    XL_SHIFT_UP = -4162
    XL_SHIFT_LEFT = -4159
    XL_NONE = -4142
    XL_FILEFORMAT_XLSM = 52
    XL_FORMULAS = -4123
    XL_PART = 2
    XL_BY_ROWS = 1
    XL_NEXT = 1

    BLUE_COLORS = {"0066FF", "0070C0"}
    BLUE_ASSUMPTIONS = {"0066FF", "0070C0", "00B0F0", "1F497D"}
    FAST_SKIP_COSMETIC_SCANS = False

    SHEET_SCAN_LIMITS = {
        "Executive Summary": (220, 40),
        "Development Summary": (260, 40),
        "Cash Flow": (260, 40),
        "Assumptions": (450, 70),
        "Development": (420, 50),
        "Sale Proceeds": (220, 30),
        "Returns Exhibit": (220, 30),
        "1-Yr Waterfall": (260, 40),
        "3-Yr Waterfall": (260, 40),
        "4-Yr Waterfall": (260, 40),
    }

    def __init__(self, input_file: Path, option: str = "base"):
        self.input_file = Path(input_file).resolve()
        self.log = RunLog()
        self.excel = None
        self.wb = None
        self.option = option

    def run(self, output_dir: Path | None = None) -> dict[str, Any]:
        output_dir = (Path(output_dir) if output_dir else self.input_file.parent).resolve()
        output_dir.mkdir(parents=True, exist_ok=True)

        try:
            import pythoncom
            import win32com.client as win32

            pythoncom.CoInitialize()
            self.excel = win32.DispatchEx("Excel.Application")
            self.excel.Visible = False
            self.excel.DisplayAlerts = False
            self.excel.ScreenUpdating = False
            self.excel.EnableEvents = False

            self.wb = self.excel.Workbooks.Open(str(self.input_file), UpdateLinks=0, ReadOnly=False)

            city_compact, city_spaced = self._parse_city_from_e6()
            project_type = self._project_type_prefix()
            cash_ws = self._ws("Cash Flow")
            baseline_q46 = cash_ws.Range("Q46").Value if cash_ws is not None else None
            original_q46_formula = cash_ws.Range("Q46").Formula if cash_ws is not None else None
            original_q51_formula = cash_ws.Range("Q51").Formula if cash_ws is not None else None

            self._step_2_blue_to_black("Executive Summary", self.BLUE_COLORS)
            self._step_3_set_project_name(city_spaced, project_type)
            self._step_4_hardcode_cells("Executive Summary", ["J15", "K15", "L15"])
            self._step_5_dscr_check()
            self._step_6_delete_zero_opex_rows()
            self._step_7_clear_n1_u12_preserve_n_border()
            self._step_8_conditional_row_collapse()

            self._step_9_group_hide_cash_flow_rows()
            self._step_10_conditional_delete_cashflow_row()

            if self.FAST_SKIP_COSMETIC_SCANS:
                self.log.add("Step 11 skipped in COM fast mode (cosmetic)")
            else:
                self._step_11_blue_to_black_assumptions()
            self._step_12_remove_comments("Assumptions")
            self._step_13_group_hide_assumptions_dash_rows()
            self._step_14_hardcode_retail_tbd_sf()
            self._step_15_hardcode_assumptions_external_refs()

            self._step_12_remove_comments("Development")
            if self.FAST_SKIP_COSMETIC_SCANS:
                self.log.add("Step 17 skipped in COM fast mode (cosmetic)")
            else:
                self._step_17_non_black_non_white_to_black_development()
            self._step_18_hardcode_development_external_refs()
            self._step_19_group_hide_development_unused_rows()
            self._step_19b_clear_development_comments_description_values()

            self._step_20_group_hide_sale_proceeds_commercial()
            self._clear_to_white("Sale Proceeds", "B75:J112")

            self._hardcode_refs_to_sheet("Building Program")
            self._safe_delete_sheet("Building Program")
            self._hardcode_refs_to_sheet("Construction Pricing")
            self._ensure_cashflow_q46_q51_formulas(original_q46_formula, original_q51_formula)
            self._safe_delete_sheet("Construction Pricing")
            self._hardcode_refs_to_sheet("3-Yr Co-GP Waterfall")
            self._safe_delete_sheet("3-Yr Co-GP Waterfall")
            self._clear_to_white("Returns Exhibit", "J3:O37")

            for s in [
                "Reassessed 1-Yr Waterfall",
                "Reassessed 3-Yr Waterfall",
                "Reassessed 4-Yr Waterfall",
                "Historic OpEx Comparison",
                "Change Log",
                "TermSheet",
                "TermSheet-Ignore",
                "Proforma Comparison",
            ]:
                self._safe_delete_sheet(s)

            self._step_36b_delete_user_requested_tabs()

            self._clear_to_white("Development", "L21:O31")
            self._step_38_clear_yellow_reference_cells()
            if self.FAST_SKIP_COSMETIC_SCANS:
                self.log.add("Step 39 skipped in COM fast mode (cosmetic)")
            else:
                self._step_39_remove_non_approved_fill_colors_assumptions()
            self._step_40_hardcode_residential_parking_spaces()
            self._step_40b_hardcode_residential_parking_rent_stall()
            self._step_41_copy_assumptions_range()
            self._clear_to_white("Assumptions", "W4:AA44")
            self._step_15b_clear_total_interim_income_adjacent()
            self._step_42b_remove_waterfall_comments_notes()

            if self.option in {"front-range", "lp", "lender"}:
                self._apply_front_range_variant()
                self.log.add("Front Range variant steps applied")
            if self.option in {"lp", "lender"}:
                self._apply_lp_variant()
                self.log.add("LP variant steps applied")
            if self.option == "lender":
                self._apply_lender_variant()
                self.log.add("Lender variant steps applied")

            try:
                self.excel.Calculate()
            except Exception:
                self.log.add("Warning: calculation call skipped")
            if self.option == "lender":
                self.log.add("Q46 integrity check skipped for lender option")
            else:
                self._assert_q46_consistency(baseline_q46)

            today = dt.datetime.now().strftime("%Y%m%d")
            market_slug = self._market_slug(default=city_compact)
            out_name = f"{project_type}_{market_slug}_{today}.xlsm"
            out_path = (output_dir / out_name).resolve()
            if out_path.exists():
                out_path.unlink()
            try:
                self.wb.SaveAs(str(out_path), FileFormat=self.XL_FILEFORMAT_XLSM)
            except Exception:
                self.wb.SaveCopyAs(str(out_path))
            self.log.add(f"Step 1: file renamed on output to {out_name}")

            return {"output_file": str(out_path), "log": self.log.messages}
        finally:
            self._cleanup()

    def _cleanup(self) -> None:
        try:
            if self.wb is not None:
                self.wb.Close(SaveChanges=False)
        except Exception:
            pass
        try:
            if self.excel is not None:
                self.excel.Quit()
        except Exception:
            pass
        self.wb = None
        self.excel = None

    def _ws(self, name: str):
        if self.wb is None:
            return None
        for ws in self.wb.Worksheets:
            if ws.Name.lower() == name.lower():
                return ws
        return None

    def _rgb_hex_from_excel_color(self, color_val: Any) -> str | None:
        if color_val in (None, ""):
            return None
        try:
            n = int(color_val)
        except Exception:
            return None
        r = n & 255
        g = (n >> 8) & 255
        b = (n >> 16) & 255
        return f"{r:02X}{g:02X}{b:02X}"

    def _addr(self, row: int, col: int) -> str:
        letters = ""
        n = col
        while n > 0:
            n, rem = divmod(n - 1, 26)
            letters = chr(65 + rem) + letters
        return f"{letters}{row}"


    def _to_matrix(self, data: Any, rows: int, cols: int) -> list[list[Any]]:
        if rows == 1 and cols == 1:
            return [[data]]
        if not isinstance(data, tuple):
            return [[data]]
        if rows == 1:
            if len(data) == 1 and isinstance(data[0], tuple):
                return [list(data[0])]
            return [list(data)]
        if cols == 1:
            out = []
            for item in data:
                if isinstance(item, tuple):
                    out.append([item[0] if item else None])
                else:
                    out.append([item])
            return out
        out = []
        for item in data:
            if isinstance(item, tuple):
                out.append(list(item))
            else:
                out.append([item])
        return out

    def _scan_bounds(self, ws, sheet_name: str | None = None) -> tuple[int, int]:
        name = sheet_name or ws.Name
        lim_row, lim_col = self.SHEET_SCAN_LIMITS.get(name, (500, 80))
        try:
            ur = ws.UsedRange
            last_row = int(ur.Row + ur.Rows.Count - 1)
            last_col = int(ur.Column + ur.Columns.Count - 1)
        except Exception:
            last_row, last_col = lim_row, lim_col
        max_row = max(1, min(last_row, lim_row))
        max_col = max(1, min(last_col, lim_col))
        return max_row, max_col

    def _to_float_or_none(self, v: Any) -> float | None:
        try:
            if v is None or v == "":
                return None
            return float(v)
        except Exception:
            return None

    def _ensure_cashflow_q46_q51_formulas(self, original_q46_formula: Any, original_q51_formula: Any) -> None:
        ws = self._ws("Cash Flow")
        if ws is None:
            self.log.add("Q46/Q51 formula preservation skipped: missing Cash Flow")
            return

        updated = 0
        for addr, original in (("Q46", original_q46_formula), ("Q51", original_q51_formula)):
            if not isinstance(original, str):
                continue
            original_formula = original.lstrip()
            if not original_formula.startswith("="):
                continue
            if "construction pricing" not in original_formula.lower():
                continue
            ws.Range(addr).Formula = self._replace_construction_pricing_refs_in_formula_com(original_formula)
            updated += 1

        if updated > 0:
            self.log.add("Preserved Cash Flow formulas at Q46 and Q51 by replacing Construction Pricing refs in original formulas")
        else:
            self.log.add("Q46/Q51 formula preservation skipped: original formulas did not require Construction Pricing substitution")
    def _assert_q46_consistency(self, baseline_q46: Any) -> None:
        ws = self._ws("Cash Flow")
        if ws is None:
            raise WorkflowError("Q46 integrity check failed: Cash Flow sheet missing")

        final_q46 = ws.Range("Q46").Value
        b = self._to_float_or_none(baseline_q46)
        f = self._to_float_or_none(final_q46)
        if b is None or f is None:
            raise WorkflowError(
                f"Q46 integrity check failed: non-numeric baseline/final values baseline={baseline_q46!r}, final={final_q46!r}"
            )
        if abs(b - f) > 1e-8:
            raise WorkflowError(
                f"Q46 integrity check failed: baseline={b} final={f} (difference={f-b})"
            )
        self.log.add(f"Q46 integrity check passed: baseline={b} final={f}")


    def _tax_abatement_pref(self) -> str | None:
        raw = os.environ.get("VERVE_TAX_ABATEMENT", "").strip().lower()
        if raw in {"yes", "no"}:
            return raw
        return None

    def _project_type_prefix(self) -> str:
        raw = os.environ.get("VERVE_PROJECT_TYPE", "VERVE").strip().upper()
        if raw in {"VERVE", "EVER", "LOCAL"}:
            return raw
        return "VERVE"

    def _market_slug(self, default: str) -> str:
        raw = os.environ.get("VERVE_MARKET", "").strip().upper()
        if not raw:
            raw = default.strip().upper()
        raw = re.sub(r"[^A-Z0-9]+", "_", raw)
        raw = re.sub(r"_+", "_", raw).strip("_")
        return raw or "MARKET"

    def _protected_sheet_names_for_user_delete(self) -> set[str]:
        return {
            "executive summary",
            "development summary",
            "cash flow",
            "assumptions",
            "development",
            "sale proceeds",
            "returns exhibit",
            "1-yr waterfall",
            "3-yr waterfall",
            "4-yr waterfall",
        }

    def _additional_delete_tabs_from_env(self) -> list[str]:
        raw = os.environ.get("VERVE_ADDITIONAL_DELETE_TABS", "").strip()
        if not raw:
            return []

        names: list[str] = []
        seen: set[str] = set()

        parsed: Any = None
        try:
            import json

            parsed = json.loads(raw)
        except Exception:
            parsed = None

        if isinstance(parsed, list):
            items = parsed
        else:
            items = re.split(r"[\n,]+", raw)

        for item in items:
            name = str(item or "").strip()
            if not name:
                continue
            key = name.lower()
            if key in seen:
                continue
            names.append(name)
            seen.add(key)
        return names

    def _step_36b_delete_user_requested_tabs(self) -> None:
        tabs = self._additional_delete_tabs_from_env()
        if not tabs:
            self.log.add("Step 36b: no user-requested additional tabs")
            return

        protected = self._protected_sheet_names_for_user_delete()
        deleted: list[str] = []
        skipped_protected: list[str] = []
        for tab in tabs:
            if tab.lower() in protected:
                skipped_protected.append(tab)
                self.log.add(f"Step 36b: skipped protected sheet '{tab}'")
                continue
            self._hardcode_refs_to_sheet(tab)
            self._safe_delete_sheet(tab)
            deleted.append(tab)

        msg = f"Step 36b: processed {len(tabs)} user-requested tabs"
        if deleted:
            msg += f"; attempted delete/hardcode for: {', '.join(deleted)}"
        if skipped_protected:
            msg += f"; protected skips: {', '.join(skipped_protected)}"
        self.log.add(msg)
    def _parse_city_from_e6(self) -> tuple[str, str]:
        ws = self._ws("Executive Summary")
        if ws is None:
            raise WorkflowError("Missing sheet Executive Summary")
        city_override = os.environ.get("VERVE_CITY_OVERRIDE", "").strip()
        if city_override:
            self.log.add(f"City override used from env: {city_override}")
            return city_override.replace(" ", ""), city_override

        raw = ws.Range("E6").Value
        if isinstance(raw, str) and "," in raw:
            parts = [p.strip() for p in raw.split(",")]
            if len(parts) >= 2 and parts[1]:
                city_spaced = parts[1]
                return city_spaced.replace(" ", ""), city_spaced

        # Fallback 1: derive from Executive Summary!E5 (e.g. "VERVE Chapel Hill")
        e5 = ws.Range("E5").Value
        if isinstance(e5, str) and e5.strip():
            guess = re.sub(r"(?i)^verve\s+", "", e5.strip()).strip()
            if guess:
                self.log.add(f"City parse fallback used E5 ('{guess}') because E6 was not parseable")
                return guess.replace(" ", ""), guess

        # Fallback 2: derive from filename pattern (e.g. Proforma_Limestone_YYYYMMDD*.xlsm)
        stem = self.input_file.stem
        m = re.search(r"(?i)proforma_([A-Za-z][A-Za-z\s-]*)_\d{8}", stem)
        if m:
            guess = m.group(1).replace("_", " ").strip()
            if guess:
                self.log.add(f"City parse fallback used filename ('{guess}') because E6 was not parseable")
                return guess.replace(" ", ""), guess

        raise WorkflowError("Executive Summary!E6 missing parseable address and no fallback city available")

    def _step_2_blue_to_black(self, sheet_name: str, targets: set[str]) -> None:
        ws = self._ws(sheet_name)
        if ws is None:
            self.log.add(f"Step 2 skipped: sheet '{sheet_name}' missing")
            return
        max_row, max_col = self._scan_bounds(ws, sheet_name)
        changed = 0
        for r in range(1, max_row + 1):
            for c in range(1, max_col + 1):
                cell = ws.Cells(r, c)
                hex_color = self._rgb_hex_from_excel_color(cell.Font.Color)
                if hex_color in targets:
                    cell.Font.Color = 0
                    changed += 1
        self.log.add(f"Step 2: updated {changed} cells on {sheet_name}")
    def _step_3_set_project_name(self, city_spaced: str, project_type: str) -> None:
        ws = self._ws("Executive Summary")
        if ws is None:
            return
        ws.Range("E5").Value = f"{project_type} {city_spaced}"
        self.log.add("Step 3: set Executive Summary!E5")

    def _step_4_hardcode_cells(self, sheet_name: str, cells: list[str]) -> None:
        ws = self._ws(sheet_name)
        if ws is None:
            return
        for addr in cells:
            ws.Range(addr).Value = ws.Range(addr).Value
        self.log.add(f"Step 4: hardcoded {sheet_name} {', '.join(cells)}")

    def _step_5_dscr_check(self) -> None:
        ws = self._ws("Executive Summary")
        if ws is None:
            return
        val = ws.Range("E60").Value
        try:
            fval = float(val)
        except Exception:
            self.log.add("Step 5: DSCR pass or unavailable")
            return
        if fval < 1.25:
            ws.Range("E60").Interior.Color = 65535
            self.log.add(f"Step 5 warning: DSCR {fval} below 1.25")
        else:
            self.log.add("Step 5: DSCR pass or unavailable")

    def _find_label_row(self, ws, label: str, min_row: int, max_row: int, col: str = "G") -> int | None:
        for r in range(min_row, max_row + 1):
            v = ws.Range(f"{col}{r}").Value
            if isinstance(v, str) and v.strip().lower() == label.lower():
                return r
        return None

    def _step_6_delete_zero_opex_rows(self) -> None:
        ws = self._ws("Executive Summary")
        if ws is None:
            return
        labels = ["Ground Lease", "Tax Abatement"]
        tax_pref = self._tax_abatement_pref()
        deleted_above = 0
        for label in labels:
            val_row = self._find_label_row(ws, label, 50, 80, col="G")
            if not val_row:
                self.log.add(f"Step 6: label '{label}' not found in OpEx section")
                continue
            val = ws.Range(f"K{val_row}").Value
            target_row = val_row - deleted_above

            force_delete = label == "Tax Abatement" and tax_pref == "no"
            force_keep = label == "Tax Abatement" and tax_pref == "yes"
            should_delete = force_delete or (not force_keep and isinstance(val, (int, float)) and float(val) == 0)

            if should_delete:
                ws.Range(f"G{target_row}:L{target_row}").Delete(Shift=self.XL_SHIFT_UP)
                deleted_above += 1
                mode = "forced by tax_abatement=no" if force_delete else "value==0"
                self.log.add(f"Step 6: deleted {label} row {target_row} ({mode})")
            else:
                mode = "forced keep by tax_abatement=yes" if force_keep else f"value {val}"
                self.log.add(f"Step 6: kept {label} row {target_row} ({mode})")

    def _set_range_white_no_border(self, rng) -> None:
        rng.Clear()
        rng.Interior.Color = 16777215
        for idx in [7, 8, 9, 10, 11, 12]:
            rng.Borders(idx).LineStyle = self.XL_NONE

    def _step_7_clear_n1_u12_preserve_n_border(self) -> None:
        ws = self._ws("Executive Summary")
        if ws is None:
            return
        self._set_range_white_no_border(ws.Range("N1:U12"))
        ws.Range("N1:N12").Borders(7).LineStyle = 1
        ws.Range("N1:N12").Borders(7).Weight = 2
        ws.Range("N1:N12").Borders(7).Color = 0
        self.log.add("Step 7: cleared N1:U12 and restored left border on N1:N12")

    def _step_8_conditional_row_collapse(self) -> None:
        ws = self._ws("Executive Summary")
        if ws is None:
            return
        l6 = ws.Range("L6").Value
        l7 = ws.Range("L7").Value
        cond = False
        if isinstance(l6, (int, float)) and isinstance(l7, (int, float)):
            cond = abs(float(l6) - float(l7)) < 1e-10
        else:
            cond = str(l6).strip() == str(l7).strip()
        if not cond:
            self.log.add(f"Step 8: condition not met (L6={l6}, L7={l7})")
            return

        ws.Range("G7").Clear()
        ws.Range("L7").Clear()
        ws.Range("K5").Value = ws.Range("K6").Value
        ws.Range("K6").Value = ws.Range("K7").Value
        ws.Range("G7").Value = ws.Range("G8").Value
        ws.Range("L7").Formula = ws.Range("L8").Formula
        ws.Range("L7").NumberFormat = '0 "bps"'

        # Explicit formatting lock to match reference workbook in G5:L7.
        ws.Range("K5").Font.Bold = True
        ws.Range("K5").Font.Underline = 2
        ws.Range("K5").HorizontalAlignment = -4152  # right

        ws.Range("G7").Font.Bold = True
        ws.Range("G7").Font.Name = "Cambria"
        ws.Range("G7").HorizontalAlignment = -4131  # left

        ws.Range("L7").Font.Bold = True
        ws.Range("L7").Font.Name = "Cambria"
        ws.Range("L7").HorizontalAlignment = -4108  # center
        ws.Range("G8").Clear()
        ws.Range("L8").Clear()
        ws.Range("K7").Clear()
        ws.Range("G6").Formula = str(ws.Range("G6").Formula).replace("K7", "K6")
        ws.Range("L6").Formula = str(ws.Range("L6").Formula).replace("K7", "K6")
        if isinstance(ws.Range("L7").Formula, str):
            ws.Range("L7").Formula = ws.Range("L7").Formula.replace("L7", "L6")
        self.log.add("Step 8: row collapse executed")

    def _step_9_group_hide_cash_flow_rows(self) -> None:
        ws = self._ws("Cash Flow")
        if ws is None:
            return
        grouped = []
        for row in [14, 30, 41]:
            vals = ws.Range(f"F{row}:Q{row}").Value
            arr = vals[0] if isinstance(vals, tuple) else vals
            nums = [x for x in arr if isinstance(x, (int, float))]
            if nums and all(float(x) == 0 for x in nums):
                ws.Rows(row).Hidden = True
                grouped.append(row)
        self.log.add(f"Step 9: grouped rows {grouped}")

    def _step_10_conditional_delete_cashflow_row(self) -> None:
        ws = self._ws("Cash Flow")
        if ws is None:
            return

        tax_pref = self._tax_abatement_pref()
        if tax_pref != "no":
            self.log.add("Step 10: skipped (tax_abatement != no)")
            return

        target = "return on cost (net of tax abatement)"
        max_row, max_col = self._scan_bounds(ws, "Cash Flow")
        max_col = min(max_col, 40)

        for row in range(1, max_row + 1):
            found = False
            for col in range(1, max_col + 1):
                v = ws.Cells(row, col).Value
                if isinstance(v, str) and target in v.strip().lower():
                    found = True
                    break
            if found:
                ws.Rows(row).Hidden = True
                self.log.add(f"Step 10: hid Cash Flow row {row} for tax_abatement=no")
                return

        self.log.add("Step 10: target label not found")

    def _step_11_blue_to_black_assumptions(self) -> None:
        self._step_2_blue_to_black("Assumptions", self.BLUE_ASSUMPTIONS)

    def _step_12_remove_comments(self, sheet_name: str) -> None:
        ws = self._ws(sheet_name)
        if ws is None:
            return
        deleted = 0
        max_row, max_col = self._scan_bounds(ws, sheet_name)
        for r in range(1, max_row + 1):
            for c in range(1, max_col + 1):
                cell = ws.Cells(r, c)
                try:
                    if cell.Comment is not None:
                        cell.Comment.Delete()
                        deleted += 1
                except Exception:
                    pass
        self.log.add(f"Step comments cleanup on {sheet_name}: removed {deleted} comments")
    def _step_42b_remove_waterfall_comments_notes(self) -> None:
        sheets = ["1-Yr Waterfall", "3-Yr Waterfall", "4-Yr Waterfall"]
        for s in sheets:
            self._step_12_remove_comments(s)
        self.log.add("Step 42b: waterfall comments/notes cleanup complete")

    def _step_13_group_hide_assumptions_dash_rows(self) -> None:
        ws = self._ws("Assumptions")
        if ws is None:
            return
        grouped = 0
        for row in range(49, 286):
            label = ws.Range(f"C{row}").Value
            if label in (None, ""):
                continue
            all_dash = True
            for col in [chr(i) for i in range(ord("J"), ord("V") + 1)]:
                txt = str(ws.Range(f"{col}{row}").Text).strip()
                if txt not in {"-", " -", "- ", "-   ", " - "}:
                    all_dash = False
                    break
            if all_dash:
                ws.Rows(row).Hidden = True
                grouped += 1
        self.log.add(f"Step 13: grouped {grouped} Assumptions rows")

    def _step_14_hardcode_retail_tbd_sf(self) -> None:
        ws = self._ws("Assumptions")
        if ws is None:
            return
        for row in range(40, 51):
            g = ws.Range(f"G{row}").Value
            h = ws.Range(f"H{row}").Value
            i_formula = ws.Range(f"I{row}").Formula
            if g == "Retail" and h == "TBD" and isinstance(i_formula, str) and i_formula.startswith("="):
                ws.Range(f"I{row}").Value = ws.Range(f"I{row}").Value
                self.log.add(f"Step 14: hardcoded Assumptions!I{row}")
                return
        self.log.add("Step 14: target row not found")

    def _find_formula_refs(self, ws, needle: str) -> list[str]:
        refs: list[str] = []
        try:
            rng = ws.UsedRange
            first = rng.Find(
                What=needle,
                LookIn=self.XL_FORMULAS,
                LookAt=self.XL_PART,
                SearchOrder=self.XL_BY_ROWS,
                SearchDirection=self.XL_NEXT,
                MatchCase=False,
            )
            if first is None:
                return refs
            first_addr = first.Address
            cur = first
            while True:
                # Restrict to formulas only and true substring match.
                f = cur.Formula
                if isinstance(f, str) and f.startswith("=") and needle.lower() in f.lower():
                    refs.append(cur.Address)
                cur = rng.FindNext(cur)
                if cur is None or cur.Address == first_addr:
                    break
        except Exception:
            return refs
        return refs

    def _hardcode_formula_refs_in_sheet(self, ws_name: str, needles: tuple[str, ...]) -> int:
        ws = self._ws(ws_name)
        if ws is None:
            return 0
        changed = 0
        seen: set[str] = set()
        for needle in needles:
            for addr in self._find_formula_refs(ws, needle):
                if addr in seen:
                    continue
                seen.add(addr)
                cell = ws.Range(addr)
                cell.Value = cell.Value
                changed += 1
        return changed
    def _verify_no_refs(self, target_sheet_name: str, only_sheet: str | None = None) -> int:
        needle = target_sheet_name.lower()
        remaining = 0
        for ws in self.wb.Worksheets:
            if ws.Name.lower() == needle:
                continue
            if only_sheet and ws.Name.lower() != only_sheet.lower():
                continue
            remaining += len(self._find_formula_refs(ws, needle))
        return remaining
    def _step_15_hardcode_assumptions_external_refs(self) -> None:
        changed = self._hardcode_formula_refs_in_sheet("Assumptions", ("building program", "construction pricing"))
        remaining = self._verify_no_refs("Building Program", only_sheet="Assumptions") + self._verify_no_refs(
            "Construction Pricing", only_sheet="Assumptions"
        )
        if remaining > 0:
            raise WorkflowError(f"Step 15 failed: remaining refs count {remaining}")
        self.log.add(f"Step 15: hardcoded {changed} Assumptions external refs")

    def _step_15b_clear_total_interim_income_adjacent(self) -> None:
        ws = self._ws("Assumptions")
        if ws is None:
            self.log.add("Step 15b skipped: missing Assumptions")
            return

        scan_row, scan_col = self._scan_bounds(ws, "Assumptions")
        hits: list[tuple[int, int]] = []
        for r in range(1, scan_row):
            for c in range(1, scan_col + 1):
                top = ws.Cells(r, c).Value
                bottom = ws.Cells(r + 1, c).Value
                if isinstance(top, str) and isinstance(bottom, str):
                    if top.strip().upper() == "TOTAL" and bottom.strip().upper() == "INTERIM INCOME":
                        hits.append((r, c))

        if not hits:
            self.log.add("Step 15b: condition not met")
            return

        try:
            ur = ws.UsedRange
            max_row = int(ur.Row + ur.Rows.Count - 1)
            max_col = int(ur.Column + ur.Columns.Count - 1)
        except Exception:
            max_row, max_col = scan_row, scan_col

        start_cols = sorted({c + 1 for _, c in hits if c + 1 <= max_col})
        if not start_cols:
            self.log.add("Step 15b: condition found but no right-side columns to clear")
            return

        start_col = min(start_cols)
        rng = ws.Range(f"{self._addr(1, start_col)}:{self._addr(max_row, max_col)}")
        rng.ClearContents()
        for idx in [7, 8, 9, 10, 11, 12]:
            rng.Borders(idx).LineStyle = self.XL_NONE

        cleared = max_row * (max_col - start_col + 1)
        self.log.add(f"Step 15b: cleared {cleared} cells in right-side area -> {self._addr(1, start_col)}:{self._addr(max_row, max_col)}")

    def _step_17_non_black_non_white_to_black_development(self) -> None:
        ws = self._ws("Development")
        if ws is None:
            return
        changed = 0
        max_row, max_col = self._scan_bounds(ws, "Development")
        for r in range(1, max_row + 1):
            for c in range(1, max_col + 1):
                cell = ws.Cells(r, c)
                hx = self._rgb_hex_from_excel_color(cell.Font.Color)
                if hx and hx not in {"000000", "FFFFFF", "FFFFFE"}:
                    cell.Font.Color = 0
                    changed += 1
        self.log.add(f"Step 17: changed {changed} Development text colors")
    def _step_18_hardcode_development_external_refs(self) -> None:
        changed = self._hardcode_formula_refs_in_sheet("Development", ("building program", "construction pricing"))
        remaining = self._verify_no_refs("Building Program", only_sheet="Development") + self._verify_no_refs(
            "Construction Pricing", only_sheet="Development"
        )
        if remaining > 0:
            raise WorkflowError(f"Step 18 failed: remaining refs count {remaining}")
        self.log.add(f"Step 18: hardcoded {changed} Development external refs")

    def _step_19_group_hide_development_unused_rows(self) -> None:
        ws = self._ws("Development")
        if ws is None:
            return
        grouped = 0
        max_row, _ = self._scan_bounds(ws, "Development")
        for row in range(1, max_row + 1):
            lbl = ws.Range(f"F{row}").Value
            if lbl in (None, ""):
                continue
            h = ws.Range(f"H{row}").Value
            ktxt = str(ws.Range(f"K{row}").Text).strip()
            if isinstance(h, (int, float)) and float(h) == 0 and ktxt in {"-", " -", "-   ", "- "}:
                ws.Rows(row).Hidden = True
                grouped += 1
        self.log.add(f"Step 19: grouped {grouped} Development rows")
    def _step_19b_clear_development_comments_description_values(self) -> None:
        ws = self._ws("Development")
        if ws is None:
            self.log.add("Step 19b skipped: missing Development")
            return

        max_row, max_col = self._scan_bounds(ws, "Development")
        hits: list[tuple[int, int]] = []
        for r in range(1, max_row + 1):
            for c in range(1, max_col + 1):
                v = ws.Cells(r, c).Value
                if isinstance(v, str) and v.strip().upper() == "COMMENTS / DESCRIPTION":
                    hits.append((r, c))

        if not hits:
            self.log.add("Step 19b: condition not met")
            return

        headers_by_col: dict[int, list[int]] = {}
        for r, c in hits:
            headers_by_col.setdefault(c, []).append(r)

        cleared = 0
        for c, rows in headers_by_col.items():
            rows_sorted = sorted(rows)
            for i, hr in enumerate(rows_sorted):
                next_header = rows_sorted[i + 1] if i + 1 < len(rows_sorted) else max_row + 1
                for rr in range(hr + 1, next_header):
                    cell = ws.Cells(rr, c)
                    v = cell.Value
                    if v is None:
                        continue
                    if isinstance(v, str) and v.strip().upper() == "COMMENTS / DESCRIPTION":
                        continue
                    cell.Value = ""
                    cleared += 1

        cols = ", ".join(sorted({self._addr(1, c).rstrip('1') for _, c in hits}))
        self.log.add(f"Step 19b: cleared {cleared} Development comment values in column(s) {cols}")

    def _step_20_group_hide_sale_proceeds_commercial(self) -> None:
        ws = self._ws("Sale Proceeds")
        if ws is None:
            return
        all_zero = True
        for row in range(24, 31):
            vals = ws.Range(f"D{row}:M{row}").Value[0]
            for v in vals:
                txt = str(v).strip() if v is not None else ""
                if isinstance(v, (int, float)) and float(v) == 0:
                    continue
                if txt in {"$-", "$ -", "-", "", "0"}:
                    continue
                all_zero = False
                break
            if not all_zero:
                break
        if all_zero:
            ws.Rows("24:30").Hidden = True
            self.log.add("Step 20: grouped Sale Proceeds rows 24:30")
        else:
            self.log.add("Step 20: condition not met")

    def _clear_to_white(self, sheet_name: str, range_address: str) -> None:
        ws = self._ws(sheet_name)
        if ws is None:
            self.log.add(f"clearToWhite skipped, missing sheet {sheet_name}")
            return
        self._set_range_white_no_border(ws.Range(range_address))

    def _formula_literal(self, v: Any) -> str:
        if v is None:
            return "0"
        if isinstance(v, bool):
            return "1" if v else "0"
        if isinstance(v, (int, float)):
            # Excel COM can surface cell errors as large negative ints.
            if float(v) < -1_000_000_000:
                return "0"
            return str(v)

        txt = str(v).strip()
        if txt == "":
            return "0"

        low = txt.lower()
        if low in {"-", "- ", " -", " - ", "$-", "$ -", "#value!", "#ref!", "#n/a", "#div/0!", "#num!", "#name?", "#null!"}:
            return "0"

        num_txt = txt.replace(",", "").replace("$", "").replace(" ", "")
        if num_txt in {"", "-"}:
            return "0"
        if num_txt.startswith("(") and num_txt.endswith(")"):
            num_txt = "-" + num_txt[1:-1]
        try:
            float(num_txt)
            return num_txt
        except Exception:
            pass

        esc = txt.replace('"', '""')
        return f'"{esc}"'
    def _replace_construction_pricing_refs_in_formula_com(self, formula: str) -> str:
        cp_ws = self._ws("Construction Pricing")
        if cp_ws is None:
            return formula
        pattern = re.compile(r"(?:'(?:\[[^\]]+\])?construction pricing'|(?:\[[^\]]+\])?construction pricing)!\$?([A-Z]{1,3})\$?(\d+)", re.IGNORECASE)

        def repl(match: re.Match[str]) -> str:
            col = match.group(1).upper()
            row = int(match.group(2))
            val = cp_ws.Range(f"{col}{row}").Value
            return self._formula_literal(val)

        return pattern.sub(repl, formula)

    def _hardcode_refs_to_sheet(self, target_sheet_name: str) -> None:
        changed = 0
        needle = target_sheet_name.lower()
        for ws in self.wb.Worksheets:
            if ws.Name.lower() == needle:
                continue
            for addr in self._find_formula_refs(ws, needle):
                cell = ws.Range(addr)

                # Step 24 exception: preserve Cash Flow formulas while
                # replacing only Construction Pricing references with constants.
                if needle == "construction pricing" and ws.Name.lower() == "cash flow":
                    f = cell.Formula
                    if isinstance(f, str) and f.lstrip().startswith("="):
                        cell.Formula = self._replace_construction_pricing_refs_in_formula_com(f.lstrip())
                        changed += 1
                    continue

                cell.Value = cell.Value
                changed += 1

        remaining = self._verify_no_refs(target_sheet_name)
        if remaining > 0:
            raise WorkflowError(f"Hardcode incomplete for {target_sheet_name}: {remaining} refs remain")
        self.log.add(f"Hardcoded refs to {target_sheet_name}: {changed}")
    def _safe_delete_sheet(self, sheet_name: str) -> None:
        ws = self._ws(sheet_name)
        if ws is None:
            self.log.add(f"Delete skipped: sheet '{sheet_name}' not present")
            return
        remaining = self._verify_no_refs(sheet_name)
        if remaining > 0:
            raise WorkflowError(f"Cannot delete '{sheet_name}', {remaining} references remain")
        ws.Delete()
        if self._ws(sheet_name) is not None:
            raise WorkflowError(f"Failed to delete sheet '{sheet_name}'")
        self.log.add(f"Deleted sheet {sheet_name}")

    def _step_38_clear_yellow_reference_cells(self) -> None:
        cleared = 0
        for sname in ["Development", "Assumptions"]:
            ws = self._ws(sname)
            if ws is None:
                continue
            max_row, max_col = self._scan_bounds(ws, sname)
            for r in range(1, max_row + 1):
                for c in range(1, max_col + 1):
                    cell = ws.Cells(r, c)
                    f = cell.Formula
                    low = f.lower() if isinstance(f, str) else ""
                    has_cashflow_ref = ("'cash flow'!" in low) or ("cash flow!" in low)
                    fill_hex = self._rgb_hex_from_excel_color(cell.Interior.Color)
                    if has_cashflow_ref:
                        cell.Value = ""
                        if fill_hex == "FFFF00":
                            cell.Interior.Color = 16777215
                        for idx in [7, 8, 9, 10, 11, 12]:
                            cell.Borders(idx).LineStyle = self.XL_NONE
                        cleared += 1
        self.log.add(f"Step 38: cleared {cleared} Cash Flow reference cells in Development/Assumptions")
    def _step_39_remove_non_approved_fill_colors_assumptions(self) -> None:
        ws = self._ws("Assumptions")
        if ws is None:
            return
        approved = {"002060", "DCE6F1", "FFFFFF"}
        changed = 0
        max_row, max_col = self._scan_bounds(ws, "Assumptions")
        for r in range(1, max_row + 1):
            for c in range(1, max_col + 1):
                cell = ws.Cells(r, c)
                hx = self._rgb_hex_from_excel_color(cell.Interior.Color)
                if hx and hx not in approved:
                    cell.Interior.Color = 16777215
                    changed += 1
        self.log.add(f"Step 39: reset {changed} non-approved fills")
    def _step_40_hardcode_residential_parking_spaces(self) -> None:
        ws = self._ws("Assumptions")
        if ws is None:
            raise WorkflowError("Step 40 failed: Assumptions sheet missing")

        # Enforce explicit hardcode target required by workflow assertions.
        target = ws.Range("D21")
        target.Value = target.Value
        if isinstance(target.Formula, str) and target.Formula.startswith("="):
            raise WorkflowError("Step 40 failed: Assumptions!D21 still contains formula after hardcode")
        self.log.add("Step 40: hardcoded Assumptions!D21")
    def _step_40b_hardcode_residential_parking_rent_stall(self) -> None:
        ws = self._ws("Assumptions")
        if ws is None:
            raise WorkflowError("Step 40b failed: Assumptions sheet missing")

        max_row, max_col = self._scan_bounds(ws, "Assumptions")
        for r in range(1, max_row + 1):
            for c in range(1, max_col):
                cell = ws.Cells(r, c)
                label = str(cell.Value).strip().lower() if cell.Value is not None else ""
                if ("residential parking rent" in label and ("stall" in label or "space" in label)):
                    target = ws.Cells(r, c + 1)
                    formula = target.Formula
                    if isinstance(formula, str) and formula.lstrip().startswith("="):
                        target.Value = target.Value
                        self.log.add(f"Step 40b: hardcoded Assumptions!{self._addr(r, c + 1)}")
                    else:
                        self.log.add("Step 40b: target found but right cell was not a formula")
                    return

        self.log.add("Step 40b: label containing 'Residential Parking Rent' + ('Stall'/'Space') not found")
    def _step_41_copy_assumptions_range(self) -> None:
        ws = self._ws("Assumptions")
        if ws is None:
            return
        ws.Range("W22:X26").Copy(ws.Range("AC22:AD26"))
        self.log.add("Step 41: copied W22:X26 to AC22:AD26")












    def _apply_front_range_variant(self) -> None:
        self._fr_step_1_text_to_black()
        self._fr_step_2_remove_gp_fees_all()
        self._fr_step_3_5_remove_gp_return_blocks()

    def _fr_step_1_text_to_black(self) -> None:
        changed_by_sheet: list[str] = []
        for sheet_name in ["1-Yr Waterfall", "3-Yr Waterfall", "4-Yr Waterfall"]:
            ws = self._ws(sheet_name)
            if ws is None:
                changed_by_sheet.append(f"{sheet_name}:missing")
                continue
            max_row, max_col = self._scan_bounds(ws, sheet_name)
            changed = 0
            for r in range(1, max_row + 1):
                for c in range(1, max_col + 1):
                    cell = ws.Cells(r, c)
                    hx = self._rgb_hex_from_excel_color(cell.Font.Color)
                    if hx and hx != "FFFFFF" and hx != "000000":
                        cell.Font.Color = 0
                        changed += 1
            changed_by_sheet.append(f"{sheet_name}:{changed}")
        self.log.add("FR Step 1 text color updates -> " + ", ".join(changed_by_sheet))

    def _fr_step_2_remove_gp_fees_all(self) -> None:
        results = []
        for sheet_name in ["1-Yr Waterfall", "3-Yr Waterfall", "4-Yr Waterfall"]:
            result = self._fr_remove_gp_fees_block(sheet_name)
            results.append(f"{sheet_name}:{result}")
        self.log.add("FR Step 2 GP Fees -> " + ", ".join(results))

    def _fr_step_3_5_remove_gp_return_blocks(self) -> None:
        labels_1yr = {
            0: "GP Return on Equity",
            1: "GP Promote",
            2: "Total GP Return",
            4: "Co-GP Split",
            6: "Co-GP Total Return",
            7: "Subtext Total Return",
            8: "Total",
        }
        labels_3yr4yr = {
            0: "GP Return on Equity",
            1: "GP Promote",
            2: "Total GP Return",
            5: "Co-GP Split",
            7: "Co-GP Total Return",
            8: "Subtext Total Return",
            9: "Total",
        }
        r1 = self._fr_remove_gp_return_block("1-Yr Waterfall", 9, labels_1yr)
        r2 = self._fr_remove_gp_return_block("3-Yr Waterfall", 10, labels_3yr4yr)
        r3 = self._fr_remove_gp_return_block("4-Yr Waterfall", 10, labels_3yr4yr)
        self.log.add(f"FR Steps 3-5 GP Return blocks -> 1Yr:{r1}, 3Yr:{r2}, 4Yr:{r3}")

    def _fr_find_exact_label_row(self, ws, label: str) -> int | None:
        max_row, _ = self._scan_bounds(ws, ws.Name)
        target = label.strip().lower()
        for r in range(1, max_row + 1):
            v = ws.Range(f"A{r}:XFD{r}").Value
            row_vals = self._to_matrix(v, 1, 16384)[0] if isinstance(v, tuple) else []
            for item in row_vals:
                if isinstance(item, str) and item.strip().lower() == target:
                    return r
        return None

    def _fr_remove_gp_fees_block(self, sheet_name: str) -> str:
        ws = self._ws(sheet_name)
        if ws is None:
            return "missing"

        row = self._fr_find_label_in_column(ws, "L", "GP Fees")
        if row is None:
            return "gp_fees_not_found"

        expected = ["GP Fees", "Development Fee", "Construction Management Fee", "GP Total Return"]
        for i, exp in enumerate(expected):
            actual = ws.Range(f"L{row + i}").Value
            if str(actual).strip().lower() != exp.lower():
                return f"label_mismatch@L{row+i}"

        addresses = {f"{col}{r}" for r in range(row, row + 4) for col in ["L", "M", "N", "O", "P", "Q", "R"]}
        refs = self._fr_find_external_refs_to_addresses(sheet_name, addresses, skip_range=(sheet_name, row, row + 3, "L", "R"))
        for ref_ws, ref_addr in refs:
            cell = ref_ws.Range(ref_addr)
            cell.Value = cell.Value

        block = ws.Range(f"L{row}:R{row + 3}")
        block.ClearContents()
        for idx in [7, 8, 9, 10, 11, 12]:
            block.Borders(idx).LineStyle = self.XL_NONE
        block.Interior.Color = 16777215
        return f"cleared_L{row}:R{row+3}_refs{len(refs)}"

    def _fr_remove_gp_return_block(self, sheet_name: str, block_rows: int, label_offsets: dict[int, str]) -> str:
        ws = self._ws(sheet_name)
        if ws is None:
            return "missing"

        start_row = self._fr_find_label_in_column(ws, "T", "GP Return on Equity")
        if start_row is None:
            return "gp_return_not_found"

        for offset, expected in label_offsets.items():
            actual = ws.Range(f"T{start_row + offset}").Value
            if str(actual).strip().lower() != expected.lower():
                return f"label_mismatch@T{start_row+offset}"

        end_row = start_row + block_rows - 1
        addresses = {f"{col}{r}" for r in range(start_row, end_row + 1) for col in ["S", "T"]}
        refs = self._fr_find_external_refs_to_addresses(
            sheet_name,
            addresses,
            skip_range=(sheet_name, start_row, end_row, "S", "T"),
        )
        for ref_ws, ref_addr in refs:
            cell = ref_ws.Range(ref_addr)
            cell.Value = cell.Value

        block = ws.Range(f"S{start_row}:T{end_row}")
        block.ClearContents()
        for idx in [7, 8, 9, 10, 11, 12]:
            block.Borders(idx).LineStyle = self.XL_NONE
        block.Interior.Color = 16777215

        for rr in [start_row, start_row + 1]:
            b = ws.Range(f"S{rr}").Borders(7)
            b.LineStyle = 1
            b.Weight = 4
            b.Color = 0

        return f"cleared_S{start_row}:T{end_row}_refs{len(refs)}"

    def _fr_find_label_in_column(self, ws, col: str, label: str) -> int | None:
        max_row, _ = self._scan_bounds(ws, ws.Name)
        want = label.strip().lower()
        for r in range(1, max_row + 1):
            v = ws.Range(f"{col}{r}").Value
            if isinstance(v, str) and v.strip().lower() == want:
                return r
        return None

    def _fr_find_external_refs_to_addresses(
        self,
        target_sheet_name: str,
        addresses: set[str],
        skip_range: tuple[str, int, int, str, str] | None = None,
    ) -> list[tuple[Any, str]]:
        refs: list[tuple[Any, str]] = []
        seen: set[tuple[str, str]] = set()

        skip_sheet = None
        skip_r1 = skip_r2 = None
        skip_c1 = skip_c2 = None
        if skip_range is not None:
            skip_sheet, skip_r1, skip_r2, skip_col1, skip_col2 = skip_range
            skip_c1 = self._col_to_num(skip_col1)
            skip_c2 = self._col_to_num(skip_col2)

        patterns = []
        for a in addresses:
            patterns.append(f"''{target_sheet_name}''!{a}".replace("''", "'").lower())
            patterns.append(f"{target_sheet_name}!{a}".lower())

        for ws in self.wb.Worksheets:
            for addr in self._find_formula_refs(ws, target_sheet_name):
                norm = str(addr).replace("$", "")
                cell = ws.Range(addr)
                formula = cell.Formula
                if not (isinstance(formula, str) and formula.startswith("=")):
                    continue

                low = formula.lower()
                if not any(p in low for p in patterns):
                    continue

                m = re.match(r"([A-Z]+)(\d+)", norm)
                if m is None:
                    continue
                c = self._col_to_num(m.group(1))
                r = int(m.group(2))

                if (
                    skip_sheet is not None
                    and ws.Name.lower() == str(skip_sheet).lower()
                    and skip_r1 <= r <= skip_r2
                    and skip_c1 <= c <= skip_c2
                ):
                    continue

                key = (ws.Name.lower(), norm)
                if key in seen:
                    continue
                seen.add(key)
                refs.append((ws, norm))

        return refs

    def _col_to_num(self, col: str) -> int:
        n = 0
        for ch in col.upper():
            n = n * 26 + (ord(ch) - 64)
        return n





























    def _apply_lp_variant(self) -> None:
        sheets = ["1-Yr Waterfall", "3-Yr Waterfall", "4-Yr Waterfall"]
        self._lp_step_1_clear_project_level_subblock(sheets)
        self._lp_step_2_copy_equity_multiple_to_leveraged(sheets)
        self._lp_step_3_link_profit_to_net_cash_flow(sheets)
        self._lp_step_4_link_irr_to_leveraged_irr(sheets)
        self._lp_step_5_link_em_to_leveraged_em(sheets)
        self._lp_step_6_clear_waterfall_terms_block(sheets)
        self._lp_step_7_move_cash_flow_summary_up(sheets)
        self._lp_step_8_delete_empty_rows_58_175(sheets)
        self._lp_step_9_clear_summary_distributions_to_340(sheets)
        self._lp_step_10_link_executive_summary_to_waterfalls()
        self._lp_step_11_rename_waterfall_tabs()
        self._lp_step_12_13_delete_lp_tabs()

    def _lp_find_exact_in_column(self, ws, col: str, label: str) -> int | None:
        max_row, _ = self._scan_bounds(ws, ws.Name)
        want = label.strip().lower()
        for r in range(1, max_row + 1):
            v = ws.Range(f"{col}{r}").Value
            if isinstance(v, str) and v.strip().lower() == want:
                return r
        return None

    def _lp_is_blank(self, value: Any) -> bool:
        if value is None:
            return True
        if isinstance(value, str) and value.strip() == "":
            return True
        return False

    def _lp_step_1_clear_project_level_subblock(self, sheets: list[str]) -> None:
        out = []
        for sheet_name in sheets:
            ws = self._ws(sheet_name)
            if ws is None:
                out.append(f"{sheet_name}:missing")
                continue
            project_row = self._lp_find_exact_in_column(ws, "L", "Project Level")
            if project_row is None:
                out.append(f"{sheet_name}:project_level_not_found")
                continue
            start_row = project_row + 1
            end_row = project_row + 5
            rng = ws.Range(f"L{start_row}:R{end_row}")
            rng.Clear()
            rng.Interior.Color = 16777215
            for idx in [7, 8, 9, 10, 11, 12]:
                rng.Borders(idx).LineStyle = self.XL_NONE
            top = ws.Range(f"L{start_row}:R{start_row}").Borders(8)
            top.LineStyle = 1
            top.Weight = 2
            top.Color = 0
            out.append(f"{sheet_name}:L{start_row}:R{end_row}")
        self.log.add("LP Step 1 Project Level clear -> " + ", ".join(out))

    def _lp_step_2_copy_equity_multiple_to_leveraged(self, sheets: list[str]) -> None:
        out = []
        for sheet_name in sheets:
            ws = self._ws(sheet_name)
            if ws is None:
                out.append(f"{sheet_name}:missing")
                continue

            ucf_row = self._lp_find_exact_in_column(ws, "B", "Unleveraged Cash Flow")
            if ucf_row is None:
                out.append(f"{sheet_name}:ucf_not_found")
                continue
            if str(ws.Range(f"B{ucf_row+1}").Value).strip().lower() != "irr" or str(ws.Range(f"B{ucf_row+2}").Value).strip().lower() != "equity multiple":
                out.append(f"{sheet_name}:test1_failed")
                continue
            source_em_row = ucf_row + 2

            dscr_row = self._lp_find_exact_in_column(ws, "B", "DSCR")
            if dscr_row is None:
                out.append(f"{sheet_name}:dscr_not_found")
                continue
            if str(ws.Range(f"B{dscr_row-2}").Value).strip().lower() != "irr" or not self._lp_is_blank(ws.Range(f"B{dscr_row-1}").Value):
                out.append(f"{sheet_name}:test2_failed")
                continue
            dest_row = dscr_row - 1

            ncf_row = self._lp_find_exact_in_column(ws, "B", "Net Cash Flow")
            if ncf_row is None:
                out.append(f"{sheet_name}:ncf_not_found")
                continue

            ws.Range(f"B{source_em_row}:D{source_em_row}").Copy(ws.Range(f"B{dest_row}:D{dest_row}"))
            ws.Range(f"D{dest_row}").Formula = f"=IFERROR(-SUMIF($F{ncf_row}:$IL{ncf_row},\">0\")/SUMIF($F{ncf_row}:$IL{ncf_row},\"<0\"),0)"
            out.append(f"{sheet_name}:B{source_em_row}:D{source_em_row}->B{dest_row}:D{dest_row}")
        self.log.add("LP Step 2 Equity Multiple copy -> " + ", ".join(out))

    def _lp_step_3_link_profit_to_net_cash_flow(self, sheets: list[str]) -> None:
        out = []
        for sheet_name in sheets:
            ws = self._ws(sheet_name)
            if ws is None:
                out.append(f"{sheet_name}:missing")
                continue
            profit_header_row = self._lp_find_exact_in_column(ws, "P", "Profit")
            ncf_row = self._lp_find_exact_in_column(ws, "B", "Net Cash Flow")
            if profit_header_row is None or ncf_row is None:
                out.append(f"{sheet_name}:missing_profit_or_ncf")
                continue
            target_row = profit_header_row + 2
            ws.Range(f"P{target_row}").Formula = f"=E{ncf_row}"
            out.append(f"{sheet_name}:P{target_row}=E{ncf_row}")
        self.log.add("LP Step 3 Profit links -> " + ", ".join(out))

    def _lp_step_4_link_irr_to_leveraged_irr(self, sheets: list[str]) -> None:
        out = []
        for sheet_name in sheets:
            ws = self._ws(sheet_name)
            if ws is None:
                out.append(f"{sheet_name}:missing")
                continue
            irr_header_row = self._lp_find_exact_in_column(ws, "Q", "IRR")
            irr_row = self._lp_find_exact_in_column(ws, "B", "IRR")
            if irr_header_row is None or irr_row is None:
                out.append(f"{sheet_name}:irr_not_found")
                continue
            target_row = irr_header_row + 2
            ws.Range(f"Q{target_row}").Formula = f"=D{irr_row}"
            out.append(f"{sheet_name}:Q{target_row}=D{irr_row}")
        self.log.add("LP Step 4 IRR links -> " + ", ".join(out))

    def _lp_step_5_link_em_to_leveraged_em(self, sheets: list[str]) -> None:
        out = []
        for sheet_name in sheets:
            ws = self._ws(sheet_name)
            if ws is None:
                out.append(f"{sheet_name}:missing")
                continue
            em_header_row = self._lp_find_exact_in_column(ws, "R", "Equity Multiple")
            em_row = self._lp_find_exact_in_column(ws, "B", "Equity Multiple")
            if em_header_row is None or em_row is None:
                out.append(f"{sheet_name}:em_not_found")
                continue
            target_row = em_header_row + 2
            ws.Range(f"R{target_row}").Formula = f"=D{em_row}"
            out.append(f"{sheet_name}:R{target_row}=D{em_row}")
        self.log.add("LP Step 5 Equity Multiple links -> " + ", ".join(out))

    def _lp_step_6_clear_waterfall_terms_block(self, sheets: list[str]) -> None:
        out = []
        for sheet_name in sheets:
            ws = self._ws(sheet_name)
            if ws is None:
                out.append(f"{sheet_name}:missing")
                continue
            guard = str(ws.Range("B19").Value or "").strip().upper()
            if "CASH FLOW SUMMARY" not in guard:
                out.append(f"{sheet_name}:guard_failed")
                continue
            rng = ws.Range("B7:J18")
            rng.Clear()
            rng.Interior.Color = 16777215
            for idx in [7, 8, 9, 10, 11, 12]:
                rng.Borders(idx).LineStyle = self.XL_NONE
            out.append(f"{sheet_name}:B7:J18")
        self.log.add("LP Step 6 clear terms -> " + ", ".join(out))

    def _lp_step_7_move_cash_flow_summary_up(self, sheets: list[str]) -> None:
        out = []
        for sheet_name in sheets:
            ws = self._ws(sheet_name)
            if ws is None:
                out.append(f"{sheet_name}:missing")
                continue
            b19 = str(ws.Range("B19").Value or "").strip().upper()
            b7 = ws.Range("B7").Value
            if "CASH FLOW SUMMARY" not in b19:
                out.append(f"{sheet_name}:guard_b19_failed")
                continue
            if not self._lp_is_blank(b7):
                out.append(f"{sheet_name}:guard_b7_not_blank")
                continue
            ws.Range("B19:J39").Cut(ws.Range("B7"))
            out.append(f"{sheet_name}:moved_B19:J39_to_B7")
        self.log.add("LP Step 7 move summary -> " + ", ".join(out))

    def _lp_step_8_delete_empty_rows_58_175(self, sheets: list[str]) -> None:
        out = []
        for sheet_name in sheets:
            ws = self._ws(sheet_name)
            if ws is None:
                out.append(f"{sheet_name}:missing")
                continue
            _, max_col = self._scan_bounds(ws, ws.Name)
            end_col_addr = self._addr(1, max_col)
            end_col = re.sub(r"\d+", "", end_col_addr)
            rows_to_clear: list[int] = []
            for row in range(58, 176):
                empty = True
                for col in range(1, max_col + 1):
                    v = ws.Cells(row, col).Value
                    if not self._lp_is_blank(v):
                        empty = False
                        break
                if empty:
                    rows_to_clear.append(row)

            for row in rows_to_clear:
                rng = ws.Range(f"A{row}:{end_col}{row}")
                rng.ClearContents()
                rng.Interior.Color = 16777215
                for idx in [7, 8, 9, 10, 11, 12]:
                    rng.Borders(idx).LineStyle = self.XL_NONE

            out.append(f"{sheet_name}:cleared_{len(rows_to_clear)}")
        self.log.add("LP Step 8 clear empty rows -> " + ", ".join(out))

    def _lp_step_9_clear_summary_distributions_to_340(self, sheets: list[str]) -> None:
        out = []
        for sheet_name in sheets:
            ws = self._ws(sheet_name)
            if ws is None:
                out.append(f"{sheet_name}:missing")
                continue

            start_row = 75
            keep_top = 358
            keep_bottom = 361

            try:
                ur = ws.UsedRange
                max_row = int(ur.Row + ur.Rows.Count - 1)
                max_col = int(ur.Column + ur.Columns.Count - 1)
            except Exception:
                max_row, max_col = self._scan_bounds(ws, ws.Name)

            if max_row < start_row:
                out.append(f"{sheet_name}:no_rows_at_or_below_{start_row}")
                continue

            cleared_ranges: list[str] = []

            top_end = min(max_row, keep_top - 1)
            if top_end >= start_row:
                rng1 = ws.Range(f"A{start_row}:{self._addr(top_end, max_col)}")
                rng1.ClearContents()
                rng1.Interior.Color = 16777215
                for idx in [7, 8, 9, 10, 11, 12]:
                    rng1.Borders(idx).LineStyle = self.XL_NONE
                cleared_ranges.append(f"A{start_row}:{self._addr(top_end, max_col)}")

            bottom_start = max(start_row, keep_bottom + 1)
            if max_row >= bottom_start:
                rng2 = ws.Range(f"A{bottom_start}:{self._addr(max_row, max_col)}")
                rng2.ClearContents()
                rng2.Interior.Color = 16777215
                for idx in [7, 8, 9, 10, 11, 12]:
                    rng2.Borders(idx).LineStyle = self.XL_NONE
                cleared_ranges.append(f"A{bottom_start}:{self._addr(max_row, max_col)}")

            if cleared_ranges:
                out.append(f"{sheet_name}:{';'.join(cleared_ranges)}|kept_B358:C361")
            else:
                out.append(f"{sheet_name}:no_clear_ranges|kept_B358:C361")

        self.log.add("LP Step 9 clear summary distributions -> " + ", ".join(out))

    def _lp_step_10_link_executive_summary_to_waterfalls(self) -> None:
        ws = self._ws("Executive Summary")
        if ws is None:
            self.log.add("LP Step 10 skipped: missing Executive Summary")
            return
        ws.Range("J16").Formula = "='1-Yr Waterfall'!Q22"
        ws.Range("J17").Formula = "='1-Yr Waterfall'!R22"
        ws.Range("K16").Formula = "='3-Yr Waterfall'!Q22"
        ws.Range("K17").Formula = "='3-Yr Waterfall'!R22"
        ws.Range("L16").Formula = "='4-Yr Waterfall'!Q22"
        ws.Range("L17").Formula = "='4-Yr Waterfall'!R22"
        self.log.add("LP Step 10 linked Executive Summary to waterfall Q22/R22")

    def _lp_step_11_rename_waterfall_tabs(self) -> None:
        rename_map = {
            "1-Yr Waterfall": "1-Yr Sale",
            "3-Yr Waterfall": "3-Yr Sale",
            "4-Yr Waterfall": "4-Yr Sale",
        }
        out = []
        for old_name, new_name in rename_map.items():
            old_ws = self._ws(old_name)
            if old_ws is None:
                out.append(f"{old_name}:missing")
                continue
            if self._ws(new_name) is not None:
                out.append(f"{old_name}:target_exists")
                continue
            old_ws.Name = new_name
            out.append(f"{old_name}->{new_name}")
        self.log.add("LP Step 11 rename tabs -> " + ", ".join(out))

    def _lp_step_12_13_delete_lp_tabs(self) -> None:
        self._safe_delete_sheet("FR Waterfall Analysis")
        self._safe_delete_sheet("Returns Exhibit")
        self.log.add("LP Steps 12-13 deletion checks completed")



    def _apply_lender_variant(self) -> None:
        keepers = ["Development Summary", "Cash Flow", "Assumptions"]
        self._lender_step_1_hardcode_keepers(keepers)
        self._lender_step_2_safety_and_delete_nonkeepers(keepers)
        self._lender_step_3_conditional_delete_cashflow_rows()
        self._lender_step_4_clear_below_noi()
        self._lender_step_5_clear_below_blue_col_b_assumptions()
        self._lender_step_6_clear_assumptions_blue_header_band()
        self._lender_step_7_delete_current_year_columns()
        self._lender_step_8_delete_current_rent_column()
        self._lender_step_9_clear_right_of_total_sf()

    def _lender_used_bounds(self, ws) -> tuple[int, int, int, int]:
        max_row, max_col = self._scan_bounds(ws, ws.Name)
        return 1, 1, max_row, max_col

    def _lender_value_zeroish(self, v: Any) -> bool:
        if v is None:
            return True
        if isinstance(v, (int, float)):
            return float(v) == 0
        if isinstance(v, str):
            t = v.strip()
            return t in {"", "-", " - ", "$-", "$ -"}
        return False

    def _lender_find_label_positions(self, ws, label: str) -> list[tuple[int, int]]:
        row1, col1, row2, col2 = self._lender_used_bounds(ws)
        rng = ws.Range(f"{self._addr(row1, col1)}:{self._addr(row2, col2)}")
        values = self._to_matrix(rng.Value, row2 - row1 + 1, col2 - col1 + 1)
        want = label.strip().lower()
        hits: list[tuple[int, int]] = []
        for r in range(len(values)):
            for c in range(len(values[r])):
                v = values[r][c]
                if isinstance(v, str) and v.strip().lower() == want:
                    hits.append((row1 + r, col1 + c))
        return hits

    def _lender_step_1_hardcode_keepers(self, keepers: list[str]) -> None:
        results: list[str] = []
        for name in keepers:
            ws = self._ws(name)
            if ws is None:
                results.append(f"{name}:missing")
                continue
            row1, col1, row2, col2 = self._lender_used_bounds(ws)
            ur = ws.UsedRange
            rows = row2 - row1 + 1
            cols = col2 - col1 + 1
            formulas = self._to_matrix(ur.Formula, rows, cols)
            values = self._to_matrix(ur.Value, rows, cols)
            new_values = [row[:] for row in values]
            replaced = 0
            errors = 0
            for r in range(rows):
                for c in range(cols):
                    f = formulas[r][c]
                    if not (isinstance(f, str) and f.startswith("=")):
                        continue
                    raw = values[r][c]
                    if isinstance(raw, str) and raw.startswith("#"):
                        new_values[r][c] = 0
                        errors += 1
                    else:
                        new_values[r][c] = raw
                    replaced += 1
            ur.Value = tuple(tuple(row) for row in new_values)
            results.append(f"{name}:replaced_{replaced}:errors_{errors}")
        self.log.add("Lender Step 1 hardcode -> " + ", ".join(results))

    def _lender_step_2_safety_and_delete_nonkeepers(self, keepers: list[str]) -> None:
        all_names = [ws.Name for ws in self.wb.Worksheets]
        missing = [k for k in keepers if k not in all_names]
        if missing:
            raise WorkflowError(f"Lender Step 2 failed: missing keeper sheets {missing}")

        to_delete = [n for n in all_names if n not in keepers]
        live_formula_count = 0
        cross_ref_count = 0

        for name in keepers:
            ws = self._ws(name)
            row1, col1, row2, col2 = self._lender_used_bounds(ws)
            ur = ws.UsedRange
            formulas = self._to_matrix(ur.Formula, row2 - row1 + 1, col2 - col1 + 1)
            for r in range(len(formulas)):
                for c in range(len(formulas[r])):
                    f = formulas[r][c]
                    if not (isinstance(f, str) and f.startswith("=")):
                        continue
                    live_formula_count += 1
                    low = f.lower()
                    for d in to_delete:
                        if d.lower() in low:
                            cross_ref_count += 1
                            break

        if live_formula_count > 0:
            raise WorkflowError(f"Lender Step 2 TEST A failed: {live_formula_count} live formulas remain in keeper sheets")
        if cross_ref_count > 0:
            raise WorkflowError(f"Lender Step 2 TEST B failed: {cross_ref_count} refs to to-be-deleted sheets in keeper sheets")

        deleted: list[str] = []
        for name in to_delete:
            ws = self._ws(name)
            if ws is None:
                continue
            ws.Delete()
            deleted.append(name)

        remaining = sorted([ws.Name for ws in self.wb.Worksheets])
        if sorted(keepers) != remaining:
            raise WorkflowError(f"Lender Step 2 failed final verification. Remaining={remaining}")
        self.log.add("Lender Step 2 delete non-keepers -> " + ", ".join(deleted))

    def _lender_step_3_conditional_delete_cashflow_rows(self) -> None:
        ws = self._ws("Cash Flow")
        if ws is None:
            raise WorkflowError("Lender Step 3 failed: Cash Flow sheet missing")
        labels = ["Commercial Parking", "Ground Lease", "Tax Abatement"]
        rows_to_delete: list[int] = []
        kept: list[str] = []
        for label in labels:
            hits = self._lender_find_label_positions(ws, label)
            if not hits:
                kept.append(f"{label}:not_found")
                continue
            row = hits[0][0]
            f_val = ws.Range(f"F{row}").Value
            if self._lender_value_zeroish(f_val):
                rows_to_delete.append(row)
            else:
                kept.append(f"{label}:{row}:{f_val}")
        for row in sorted(rows_to_delete, reverse=True):
            ws.Rows(row).Delete()
        self.log.add(f"Lender Step 3 row deletes -> deleted={sorted(rows_to_delete, reverse=True)} kept={kept}")

    def _lender_step_4_clear_below_noi(self) -> None:
        ws = self._ws("Cash Flow")
        if ws is None:
            raise WorkflowError("Lender Step 4 failed: Cash Flow sheet missing")
        hits = self._lender_find_label_positions(ws, "NET OPERATING INCOME (LESS RESERVES)")
        if not hits:
            raise WorkflowError("Lender Step 4 failed: NOI anchor not found")
        anchor_row = hits[0][0]
        row1, col1, row2, col2 = self._lender_used_bounds(ws)
        start_row = anchor_row + 1
        if start_row > row2:
            self.log.add("Lender Step 4 skipped: no rows below NOI")
            return
        rng = ws.Range(f"{self._addr(start_row, col1)}:{self._addr(row2, col2)}")
        rng.ClearContents()
        rng.Interior.Color = 16777215
        for idx in [7, 8, 9, 10, 11, 12]:
            rng.Borders(idx).LineStyle = self.XL_NONE
        self.log.add(f"Lender Step 4 cleared below NOI -> {self._addr(start_row, col1)}:{self._addr(row2, col2)}")

    def _lender_step_5_clear_below_blue_col_b_assumptions(self) -> None:
        ws = self._ws("Assumptions")
        if ws is None:
            raise WorkflowError("Lender Step 5 failed: Assumptions sheet missing")
        row1, col1, row2, col2 = self._lender_used_bounds(ws)
        found_row = None
        for r in range(max(46, row1), row2 + 1):
            hx = self._rgb_hex_from_excel_color(ws.Cells(r, 2).Interior.Color)
            if hx == "002060":
                found_row = r
                break
        if found_row is None:
            raise WorkflowError("Lender Step 5 failed: no #002060 cell found in col B below row 45")
        rng = ws.Range(f"B{found_row}:{self._addr(row2, col2)}")
        rng.ClearContents()
        rng.Interior.Color = 16777215
        for idx in [7, 8, 9, 10, 11, 12]:
            rng.Borders(idx).LineStyle = self.XL_NONE
        self.log.add(f"Lender Step 5 cleared assumptions below blue B -> B{found_row}:{self._addr(row2, col2)}")

    def _lender_step_6_clear_assumptions_blue_header_band(self) -> None:
        ws = self._ws("Assumptions")
        if ws is None:
            raise WorkflowError("Lender Step 6 failed: Assumptions sheet missing")
        hits = self._lender_find_label_positions(ws, "ASSUMPTIONS")
        hits = [h for h in hits if h[1] == 3]
        if not hits:
            raise WorkflowError("Lender Step 6 failed: ASSUMPTIONS not found in col C")
        anchor_row, anchor_col = hits[0]
        row1, col1, row2, col2 = self._lender_used_bounds(ws)
        end_col = anchor_col
        for c in range(anchor_col, col2 + 1):
            hx = self._rgb_hex_from_excel_color(ws.Cells(anchor_row, c).Interior.Color)
            if hx == "002060":
                end_col = c
            else:
                break
        rng = ws.Range(f"{self._addr(anchor_row, anchor_col)}:{self._addr(row2, end_col)}")
        rng.ClearContents()
        rng.Interior.Color = 16777215
        for idx in [7, 8, 9, 10, 11, 12]:
            rng.Borders(idx).LineStyle = self.XL_NONE
        self.log.add(f"Lender Step 6 cleared assumptions header band -> {self._addr(anchor_row, anchor_col)}:{self._addr(row2, end_col)}")

    def _lender_step_7_delete_current_year_columns(self) -> None:
        ws = self._ws("Assumptions")
        if ws is None:
            raise WorkflowError("Lender Step 7 failed: Assumptions sheet missing")
        p1 = self._lender_find_label_positions(ws, "Current Year Per Bed Input")
        p2 = self._lender_find_label_positions(ws, "Current Year Rent PSF")
        if not p1 or not p2:
            raise WorkflowError("Lender Step 7 failed: required labels not found")
        r1, c1 = p1[0]
        r2, c2 = p2[0]
        if r1 != r2:
            raise WorkflowError("Lender Step 7 failed: labels not in same row")
        label_row = r1
        left_col, right_col = sorted((c1, c2))

        blue_row = None
        for r in range(label_row - 1, 0, -1):
            all_blue = True
            for c in range(left_col, right_col + 1):
                if self._rgb_hex_from_excel_color(ws.Cells(r, c).Interior.Color) != "002060":
                    all_blue = False
                    break
            if all_blue:
                blue_row = r
                break
        if blue_row is None:
            raise WorkflowError("Lender Step 7 failed: blue header row above labels not found")

        row1, _, row2, _ = self._lender_used_bounds(ws)
        light_hits = 0
        bottom_row = None
        for r in range(label_row, row2 + 1):
            hx = self._rgb_hex_from_excel_color(ws.Cells(r, left_col).Interior.Color)
            if hx == "DCE6F1":
                light_hits += 1
                if light_hits == 2:
                    bottom_row = r
                    break
        if bottom_row is None:
            raise WorkflowError("Lender Step 7 failed: second #DCE6F1 boundary not found")

        for rr in [bottom_row + 1, bottom_row + 2]:
            for c in range(left_col, right_col + 1):
                if not self._lp_is_blank(ws.Cells(rr, c).Value):
                    raise WorkflowError("Lender Step 7 failed: expected two blank spacer rows after boundary")

        delete_addr = f"{self._addr(blue_row, left_col)}:{self._addr(bottom_row, right_col)}"
        ws.Range(delete_addr).Delete(Shift=self.XL_SHIFT_LEFT)
        self.log.add(f"Lender Step 7 deleted current-year cols -> {delete_addr}")

    def _lender_step_8_delete_current_rent_column(self) -> None:
        ws = self._ws("Assumptions")
        if ws is None:
            raise WorkflowError("Lender Step 8 failed: Assumptions sheet missing")
        hits = self._lender_find_label_positions(ws, "Current Rent/Yr")
        if not hits:
            raise WorkflowError("Lender Step 8 failed: Current Rent/Yr not found")
        label_row, col = hits[0]

        blue_row = None
        for r in range(label_row - 1, 0, -1):
            if self._rgb_hex_from_excel_color(ws.Cells(r, col).Interior.Color) == "002060":
                blue_row = r
                break
        if blue_row is None:
            raise WorkflowError("Lender Step 8 failed: blue header row above Current Rent/Yr not found")

        _, _, row2, _ = self._lender_used_bounds(ws)
        bottom_row = None
        for r in range(label_row, row2 + 1):
            if self._rgb_hex_from_excel_color(ws.Cells(r, col).Interior.Color) == "DCE6F1":
                bottom_row = r
                break
        if bottom_row is None:
            raise WorkflowError("Lender Step 8 failed: first #DCE6F1 boundary not found")

        delete_addr = f"{self._addr(blue_row, col)}:{self._addr(bottom_row, col)}"
        ws.Range(delete_addr).Delete(Shift=self.XL_SHIFT_LEFT)
        self.log.add(f"Lender Step 8 deleted Current Rent/Yr col -> {delete_addr}")

    def _lender_step_9_clear_right_of_total_sf(self) -> None:
        ws = self._ws("Assumptions")
        if ws is None:
            raise WorkflowError("Lender Step 9 failed: Assumptions sheet missing")
        hits = self._lender_find_label_positions(ws, "Total SF")
        if not hits:
            raise WorkflowError("Lender Step 9 failed: Total SF not found")
        hits = sorted(hits, key=lambda x: x[0])
        label_row, label_col = hits[0]

        top_row = None
        for r in range(label_row, 0, -1):
            if self._rgb_hex_from_excel_color(ws.Cells(r, label_col).Interior.Color) == "002060":
                top_row = r
                break
        if top_row is None:
            raise WorkflowError("Lender Step 9 failed: top #002060 boundary not found")

        _, _, last_row, last_col = self._lender_used_bounds(ws)
        light_hits = 0
        bottom_row = None
        for r in range(label_row, last_row + 1):
            if self._rgb_hex_from_excel_color(ws.Cells(r, label_col).Interior.Color) == "DCE6F1":
                light_hits += 1
                if light_hits == 2:
                    bottom_row = r
                    break
        if bottom_row is None:
            raise WorkflowError("Lender Step 9 failed: second #DCE6F1 boundary not found")

        start_col = label_col + 1
        if start_col > last_col:
            self.log.add("Lender Step 9 skipped: no columns to the right of Total SF")
            return

        rng = ws.Range(f"{self._addr(top_row, start_col)}:{self._addr(bottom_row, last_col)}")
        rng.ClearContents()
        rng.Interior.Color = 16777215
        for idx in [7, 8, 9, 10, 11, 12]:
            rng.Borders(idx).LineStyle = self.XL_NONE
        self.log.add(f"Lender Step 9 cleared right of Total SF -> {self._addr(top_row, start_col)}:{self._addr(bottom_row, last_col)}")

