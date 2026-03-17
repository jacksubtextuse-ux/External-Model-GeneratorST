from __future__ import annotations

import datetime as dt
import re
import os
from copy import copy
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

from openpyxl import load_workbook
from openpyxl.styles import Border, PatternFill, Side
from openpyxl.cell.cell import MergedCell
from openpyxl.utils import column_index_from_string, get_column_letter
from openpyxl.formula.translate import Translator

WHITE_FILL = PatternFill(fill_type="solid", fgColor="FFFFFFFF")
NO_BORDER = Border(
    left=Side(style=None),
    right=Side(style=None),
    top=Side(style=None),
    bottom=Side(style=None),
    diagonal=Side(style=None),
    diagonalDown=False,
    diagonalUp=False,
    outline=False,
    vertical=Side(style=None),
    horizontal=Side(style=None),
)


@dataclass
class RunLog:
    messages: list[str] = field(default_factory=list)

    def add(self, message: str) -> None:
        self.messages.append(message)


class WorkflowError(RuntimeError):
    pass


class VerveWorkflowRunner:
    def __init__(self, input_file: Path, option: str = "base"):
        self.input_file = Path(input_file)
        self.wb = load_workbook(self.input_file, data_only=False, keep_vba=True)
        self.values_wb = load_workbook(self.input_file, data_only=True, keep_vba=True)
        self.log = RunLog()
        self.option = option

    def run(self, output_dir: Path | None = None) -> dict[str, Any]:
        output_dir = Path(output_dir) if output_dir else self.input_file.parent
        output_dir.mkdir(parents=True, exist_ok=True)
        city_compact, city_spaced = self._parse_city_from_e6()
        project_type = self._project_type_prefix()
        baseline_q46 = self._value("Cash Flow", "Q46")

        self._step_2_blue_to_black("Executive Summary", {"0066FF", "0070C0"})
        self._step_3_set_project_name(city_spaced, project_type)
        self._step_4_hardcode_cells("Executive Summary", ["J15", "K15", "L15"])
        self._step_5_dscr_check()
        self._step_6_delete_zero_opex_rows()
        self._step_7_clear_n1_u12_preserve_n_border()
        self._step_8_conditional_row_collapse()

        self._step_9_group_hide_cash_flow_rows()
        self._step_10_conditional_delete_cashflow_row()

        self._step_11_blue_to_black_assumptions()
        self._step_12_remove_comments("Assumptions")
        self._step_13_group_hide_assumptions_dash_rows()
        self._step_14_hardcode_retail_tbd_sf()
        self._step_15_hardcode_assumptions_external_refs()

        self._step_12_remove_comments("Development")
        self._step_17_non_black_non_white_to_black_development()
        self._step_18_hardcode_development_external_refs()
        self._step_19_group_hide_development_unused_rows()

        self._step_20_group_hide_sale_proceeds_commercial()
        self._clear_to_white("Sale Proceeds", "B75:J112")

        self._hardcode_refs_to_sheet("Building Program")
        self._safe_delete_sheet("Building Program")
        self._hardcode_refs_to_sheet("Construction Pricing")
        self._safe_delete_sheet("Construction Pricing")
        self._hardcode_refs_to_sheet("3-Yr Co-GP Waterfall")
        self._safe_delete_sheet("3-Yr Co-GP Waterfall")
        self._step_28_clear_returns_exhibit()

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

        self._step_37_clear_development_range()
        self._step_38_clear_yellow_reference_cells()
        self._step_39_remove_non_approved_fill_colors_assumptions()
        self._step_40_hardcode_residential_parking_spaces()
        self._step_40b_hardcode_residential_parking_rent_stall()
        self._step_41_copy_assumptions_range()
        self._step_42_clear_assumptions_range()

        if self.option in {"front-range", "lp", "lender"}:
            self._apply_front_range_variant()
            self.log.add("Front Range variant steps applied")
        if self.option in {"lp", "lender"}:
            self._apply_lp_variant()
            self.log.add("LP variant steps applied")
        if self.option == "lender":
            self._apply_lender_variant()
            self.log.add("Lender variant steps applied")

        if self.option == "lender":
            self.log.add("Q46 integrity check skipped for lender option")
        else:
            self._assert_q46_consistency_best_effort(baseline_q46)
        today = dt.datetime.now().strftime("%Y%m%d")
        market_slug = self._market_slug(default=city_compact)
        out_name = f"{project_type}_{market_slug}_{today}.xlsm"
        out_path = output_dir / out_name
        self.wb.save(out_path)
        self.log.add(f"Step 1: file renamed on output to {out_name}")
        return {"output_file": str(out_path), "log": self.log.messages}

    def _sheet(self, name: str):
        for s in self.wb.worksheets:
            if s.title.lower() == name.lower():
                return s
        return None

    def _values_sheet(self, name: str):
        for s in self.values_wb.worksheets:
            if s.title.lower() == name.lower():
                return s
        return None

    def _value(self, sheet: str, cell: str) -> Any:
        ws = self._values_sheet(sheet)
        if not ws:
            return None
        return ws[cell].value

    def _assert_q46_consistency_best_effort(self, baseline_q46: Any) -> None:
        ws = self._sheet("Cash Flow")
        if ws is None:
            raise WorkflowError("Q46 integrity check failed: Cash Flow sheet missing")
        cell = ws["Q46"].value
        if isinstance(cell, str) and cell.startswith("="):
            # openpyxl cannot recalc formulas; keep as warning-only in fallback mode
            self.log.add("Q46 integrity check skipped in openpyxl mode (formula not recalculated)")
            return
        try:
            b = float(baseline_q46)
            f = float(cell)
        except Exception:
            raise WorkflowError(f"Q46 integrity check failed: non-numeric baseline/final values baseline={baseline_q46!r}, final={cell!r}")
        if abs(b - f) > 1e-8:
            raise WorkflowError(f"Q46 integrity check failed: baseline={b} final={f} (difference={f-b})")
        self.log.add(f"Q46 integrity check passed: baseline={b} final={f}")

    def _parse_city_from_e6(self) -> tuple[str, str]:
        v = self._value("Executive Summary", "E6")
        if not isinstance(v, str) or "," not in v:
            raise WorkflowError("Executive Summary!E6 missing parseable address")
        parts = [p.strip() for p in v.split(",")]
        if len(parts) < 2:
            raise WorkflowError("Executive Summary!E6 address parse failed")
        city_spaced = parts[1]
        city_compact = city_spaced.replace(" ", "")
        return city_compact, city_spaced


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

    def _normalize_rgb(self, rgb: Any) -> str | None:
        if not rgb:
            return None
        raw = str(rgb).upper()
        if len(raw) == 8:
            raw = raw[2:]
        if len(raw) == 6:
            return raw
        return None

    def _step_2_blue_to_black(self, sheet: str, targets: set[str]) -> None:
        ws = self._sheet(sheet)
        if not ws:
            self.log.add(f"Step 2 skipped: sheet '{sheet}' missing")
            return
        changed = 0
        for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
            for c in row:
                color = self._normalize_rgb(getattr(c.font.color, "rgb", None) if c.font and c.font.color else None)
                if color in targets:
                    c.font = copy(c.font)
                    c.font = c.font.copy(color="000000")
                    changed += 1
        self.log.add(f"Step 2: updated {changed} cells on {sheet}")

    def _step_3_set_project_name(self, city_spaced: str, project_type: str) -> None:
        ws = self._sheet("Executive Summary")
        if ws:
            ws["E5"].value = f"{project_type} {city_spaced}"
            self.log.add("Step 3: set Executive Summary!E5")

    def _step_4_hardcode_cells(self, sheet: str, cells: list[str]) -> None:
        ws = self._sheet(sheet)
        if not ws:
            return
        for addr in cells:
            ws[addr].value = self._value(sheet, addr)
        self.log.add(f"Step 4: hardcoded {sheet} {', '.join(cells)}")

    def _step_5_dscr_check(self) -> None:
        ws = self._sheet("Executive Summary")
        v = self._value("Executive Summary", "E60")
        if ws and isinstance(v, (int, float)) and v < 1.25:
            ws["E60"].fill = PatternFill(fill_type="solid", fgColor="FFFFFF00")
            self.log.add(f"Step 5 warning: DSCR {v} below 1.25")
        else:
            self.log.add("Step 5: DSCR pass or unavailable")

    def _find_label_row(
        self,
        ws,
        column: str,
        label: str,
        min_row: int = 1,
        max_row: int | None = None,
        exact: bool = False,
    ) -> int | None:
        col_idx = column_index_from_string(column)
        max_row = max_row or ws.max_row
        want = label.strip().lower()
        for r in range(min_row, max_row + 1):
            val = ws.cell(r, col_idx).value
            if not isinstance(val, str):
                continue
            have = val.strip().lower()
            if (exact and have == want) or ((not exact) and want in have):
                return r
        return None
    def _step_6_delete_zero_opex_rows(self) -> None:
        ws = self._sheet("Executive Summary")
        vws = self._values_sheet("Executive Summary")
        if not ws or not vws:
            return

        labels = ["Ground Lease", "Tax Abatement"]
        tax_pref = self._tax_abatement_pref()
        deleted_above = 0
        for label in labels:
            val_row = self._find_label_row(vws, "G", label, min_row=50, max_row=80, exact=True)
            if not val_row:
                self.log.add(f"Step 6: label '{label}' not found in OpEx section")
                continue

            val = vws[f"K{val_row}"].value
            target_row = val_row - deleted_above

            force_delete = label == "Tax Abatement" and tax_pref == "no"
            force_keep = label == "Tax Abatement" and tax_pref == "yes"
            should_delete = force_delete or (not force_keep and isinstance(val, (int, float)) and val == 0)

            if should_delete:
                ws.move_range(f"G{target_row + 1}:L{ws.max_row}", rows=-1, cols=0, translate=True)
                for col in range(column_index_from_string("G"), column_index_from_string("L") + 1):
                    ws.cell(ws.max_row, col).value = None
                    ws.cell(ws.max_row, col).border = copy(NO_BORDER)
                    ws.cell(ws.max_row, col).fill = copy(WHITE_FILL)
                deleted_above += 1
                mode = "forced by tax_abatement=no" if force_delete else "value==0"
                self.log.add(f"Step 6: deleted {label} row {target_row} ({mode})")
            else:
                mode = "forced keep by tax_abatement=yes" if force_keep else f"value {val}"
                self.log.add(f"Step 6: kept {label} row {target_row} ({mode})")
    def _clear_range(self, ws, start_row: int, end_row: int, start_col: int, end_col: int) -> None:
        for r in range(start_row, end_row + 1):
            for c in range(start_col, end_col + 1):
                cell = ws.cell(r, c)
                if isinstance(cell, MergedCell):
                    continue
                cell.value = None
                cell.border = copy(NO_BORDER)
                cell.fill = copy(WHITE_FILL)

    def _step_7_clear_n1_u12_preserve_n_border(self) -> None:
        ws = self._sheet("Executive Summary")
        if not ws:
            return
        self._clear_range(ws, 1, 12, column_index_from_string("N"), column_index_from_string("U"))
        for r in range(1, 13):
            c = ws[f"N{r}"]
            b = copy(c.border)
            b.left = Side(style="thin", color="FF000000")
            c.border = b
        self.log.add("Step 7: cleared N1:U12 and restored left border on N1:N12")

    def _step_8_conditional_row_collapse(self) -> None:
        ws = self._sheet("Executive Summary")
        if not ws:
            return
        l6 = self._value("Executive Summary", "L6")
        l7 = self._value("Executive Summary", "L7")

        condition_met = False
        if isinstance(l6, (int, float)) and isinstance(l7, (int, float)):
            condition_met = abs(float(l6) - float(l7)) < 1e-10
        else:
            condition_met = str(l6).strip() == str(l7).strip()

        if not condition_met:
            self.log.add(f"Step 8: condition not met (L6={l6}, L7={l7})")
            return
        ws["G7"].value = None
        ws["L7"].value = None
        ws["K5"].value = ws["K6"].value
        ws["K5"]._style = copy(ws["K6"]._style)
        ws["K6"].value = ws["K7"].value
        ws["K6"]._style = copy(ws["K7"]._style)
        ws["G7"].value = ws["G8"].value
        ws["G7"]._style = copy(ws["G8"]._style)
        ws["L7"].value = ws["L8"].value
        ws["L7"]._style = copy(ws["L8"]._style)
        ws["L7"].number_format = '0 "bps"'

        # Explicit formatting lock to match reference workbook in G5:L7.
        ws["K5"].font = copy(ws["K5"].font)
        ws["K5"].font = ws["K5"].font.copy(bold=True, underline="single")
        ws["K5"].alignment = copy(ws["K5"].alignment)
        ws["K5"].alignment = ws["K5"].alignment.copy(horizontal="right")

        ws["G7"].font = copy(ws["G7"].font)
        ws["G7"].font = ws["G7"].font.copy(bold=True, name="Cambria")
        ws["G7"].alignment = copy(ws["G7"].alignment)
        ws["G7"].alignment = ws["G7"].alignment.copy(horizontal="left")

        ws["L7"].font = copy(ws["L7"].font)
        ws["L7"].font = ws["L7"].font.copy(bold=True, name="Cambria")
        ws["L7"].alignment = copy(ws["L7"].alignment)
        ws["L7"].alignment = ws["L7"].alignment.copy(horizontal="center")
        for c in ["G8", "L8", "K7"]:
            ws[c].value = None
            ws[c].border = copy(NO_BORDER)
            ws[c].fill = copy(WHITE_FILL)
        if isinstance(ws["G6"].value, str):
            ws["G6"].value = ws["G6"].value.replace("K7", "K6")
        if isinstance(ws["L6"].value, str):
            ws["L6"].value = ws["L6"].value.replace("K7", "K6")
        if isinstance(ws["L7"].value, str):
            ws["L7"].value = ws["L7"].value.replace("L7", "L6")
        self.log.add("Step 8: row collapse executed")
    def _set_row_group_hidden(self, ws, row: int, outline: int = 1) -> None:
        rd = ws.row_dimensions[row]
        rd.hidden = True
        rd.outlineLevel = max(rd.outlineLevel or 0, outline)

    def _value_zeroish(self, v: Any) -> bool:
        if v is None:
            return True
        if isinstance(v, (int, float)):
            return v == 0
        if isinstance(v, str):
            t = v.strip()
            return t in {"", "-", " - ", "$-", "$ -"}
        return False

    def _step_9_group_hide_cash_flow_rows(self) -> None:
        ws = self._sheet("Cash Flow")
        vws = self._values_sheet("Cash Flow")
        if not ws or not vws:
            return
        grouped = []
        for row in [14, 30, 41]:
            values = [vws.cell(row, c).value for c in range(column_index_from_string("F"), column_index_from_string("Q") + 1)]
            nums = [x for x in values if isinstance(x, (int, float))]
            if nums and all(x == 0 for x in nums):
                self._set_row_group_hidden(ws, row)
                grouped.append(row)
        self.log.add(f"Step 9: grouped rows {grouped}")

    def _step_10_conditional_delete_cashflow_row(self) -> None:
        ws = self._sheet("Cash Flow")
        vws = self._values_sheet("Cash Flow")
        if not ws or not vws:
            return
        q46 = vws["Q46"].value
        q47 = vws["Q47"].value
        if q46 == q47:
            ws.delete_rows(47, 1)
            self.log.add("Step 10: deleted Cash Flow row 47")
        else:
            self.log.add("Step 10: condition not met")

    def _step_11_blue_to_black_assumptions(self) -> None:
        self._step_2_blue_to_black("Assumptions", {"0066FF", "0070C0", "00B0F0", "1F497D"})

    def _step_12_remove_comments(self, sheet: str) -> None:
        ws = self._sheet(sheet)
        if not ws:
            return
        deleted = 0
        for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
            for c in row:
                if c.comment is not None:
                    c.comment = None
                    deleted += 1
        self.log.add(f"Step comments cleanup on {sheet}: removed {deleted} comments")

    def _step_13_group_hide_assumptions_dash_rows(self) -> None:
        ws = self._sheet("Assumptions")
        vws = self._values_sheet("Assumptions")
        if not ws or not vws:
            return
        grouped = []
        for row in range(49, 286):
            label = ws[f"C{row}"].value
            if label in (None, ""):
                continue
            vals = [vws.cell(row, c).value for c in range(column_index_from_string("J"), column_index_from_string("V") + 1)]
            if vals and all(self._value_zeroish(v) for v in vals):
                self._set_row_group_hidden(ws, row)
                grouped.append(row)
        self.log.add(f"Step 13: grouped {len(grouped)} Assumptions rows")

    def _step_14_hardcode_retail_tbd_sf(self) -> None:
        ws = self._sheet("Assumptions")
        vws = self._values_sheet("Assumptions")
        if not ws or not vws:
            return
        for row in range(40, 51):
            g = ws[f"G{row}"].value
            h = ws[f"H{row}"].value
            i = ws[f"I{row}"].value
            if g == "Retail" and h == "TBD" and isinstance(i, str) and i.startswith("="):
                ws[f"I{row}"].value = vws[f"I{row}"].value
                self.log.add(f"Step 14: hardcoded Assumptions!I{row}")
                return
        self.log.add("Step 14: target row not found")

    def _hardcode_formula_refs_in_sheet(self, ws_name: str, targets: tuple[str, ...]) -> int:
        ws = self._sheet(ws_name)
        vws = self._values_sheet(ws_name)
        if not ws or not vws:
            return 0
        changed = 0
        for r in range(1, ws.max_row + 1):
            for c in range(1, ws.max_column + 1):
                cell = ws.cell(r, c)
                formula = cell.value
                if isinstance(formula, str) and formula.startswith("="):
                    low = formula.lower()
                    if any(t in low for t in targets):
                        cell.value = vws.cell(r, c).value
                        changed += 1
        return changed

    def _verify_no_refs_in_sheet(self, sheet_name: str, target_sheet_name: str) -> int:
        ws = self._sheet(sheet_name)
        if not ws:
            return 0
        needle = target_sheet_name.lower()
        remaining = 0
        for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
            for cell in row:
                if isinstance(cell.value, str) and cell.value.startswith("=") and needle in cell.value.lower():
                    remaining += 1
        return remaining
    def _verify_no_refs(self, target_sheet_name: str) -> int:
        needle = target_sheet_name.lower()
        remaining = 0
        for ws in self.wb.worksheets:
            if ws.title.lower() == needle:
                continue
            for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
                for cell in row:
                    if isinstance(cell.value, str) and cell.value.startswith("=") and needle in cell.value.lower():
                        remaining += 1
        return remaining

    def _step_15_hardcode_assumptions_external_refs(self) -> None:
        changed = self._hardcode_formula_refs_in_sheet("Assumptions", ("building program", "construction pricing"))
        remaining = self._verify_no_refs_in_sheet("Assumptions", "Building Program") + self._verify_no_refs_in_sheet("Assumptions", "Construction Pricing")
        if remaining > 0:
            raise WorkflowError(f"Step 15 failed: remaining refs count {remaining}")
        self.log.add(f"Step 15: hardcoded {changed} Assumptions external refs")

    def _step_17_non_black_non_white_to_black_development(self) -> None:
        ws = self._sheet("Development")
        if not ws:
            return
        changed = 0
        for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
            for c in row:
                color = self._normalize_rgb(getattr(c.font.color, "rgb", None) if c.font and c.font.color else None)
                if color and color not in {"000000", "FFFFFF", "FFFFFE"}:
                    c.font = copy(c.font)
                    c.font = c.font.copy(color="000000")
                    changed += 1
        self.log.add(f"Step 17: changed {changed} Development text colors")

    def _step_18_hardcode_development_external_refs(self) -> None:
        changed = self._hardcode_formula_refs_in_sheet("Development", ("building program", "construction pricing"))
        remaining = self._verify_no_refs_in_sheet("Development", "Building Program") + self._verify_no_refs_in_sheet("Development", "Construction Pricing")
        if remaining > 0:
            raise WorkflowError(f"Step 18 failed: remaining refs count {remaining}")
        self.log.add(f"Step 18: hardcoded {changed} Development external refs")

    def _step_19_group_hide_development_unused_rows(self) -> None:
        ws = self._sheet("Development")
        vws = self._values_sheet("Development")
        if not ws or not vws:
            return
        grouped = []
        for row in range(1, ws.max_row + 1):
            f_label = ws[f"F{row}"].value
            if f_label in (None, ""):
                continue
            h_val = vws[f"H{row}"].value
            k_val = vws[f"K{row}"].value
            k_text = "" if k_val is None else str(k_val).strip()
            if h_val == 0 and k_text in {"-", " -"}:
                self._set_row_group_hidden(ws, row)
                grouped.append(row)
        self.log.add(f"Step 19: grouped {len(grouped)} Development rows")

    def _step_20_group_hide_sale_proceeds_commercial(self) -> None:
        ws = self._sheet("Sale Proceeds")
        vws = self._values_sheet("Sale Proceeds")
        if not ws or not vws:
            return
        all_zero = True
        for row in range(24, 31):
            vals = [vws.cell(row, c).value for c in range(column_index_from_string("D"), column_index_from_string("M") + 1)]
            if not all(self._value_zeroish(v) for v in vals):
                all_zero = False
                break
        if all_zero:
            for row in range(24, 31):
                self._set_row_group_hidden(ws, row)
            self.log.add("Step 20: grouped Sale Proceeds rows 24:30")
        else:
            self.log.add("Step 20: condition not met")

    def _clear_to_white(self, sheet_name: str, range_address: str) -> None:
        ws = self._sheet(sheet_name)
        if not ws:
            self.log.add(f"clearToWhite skipped, missing sheet {sheet_name}")
            return
        start, end = range_address.split(":")
        start_col = column_index_from_string("".join([x for x in start if x.isalpha()]))
        start_row = int("".join([x for x in start if x.isdigit()]))
        end_col = column_index_from_string("".join([x for x in end if x.isalpha()]))
        end_row = int("".join([x for x in end if x.isdigit()]))
        self._clear_range(ws, start_row, end_row, start_col, end_col)

    def _step_28_clear_returns_exhibit(self) -> None:
        self._clear_to_white("Returns Exhibit", "J3:O37")
        self.log.add("Step 28: cleared Returns Exhibit J3:O37")

    def _step_37_clear_development_range(self) -> None:
        self._clear_to_white("Development", "L21:O31")
        self.log.add("Step 37: cleared Development L21:O31")

    def _step_38_clear_yellow_reference_cells(self) -> None:
        refs = ["'cash flow'!q46", "'executive summary'!l6", "'executive summary'!r10"]
        cleared = 0
        for sname in ["Development", "Assumptions"]:
            ws = self._sheet(sname)
            if not ws:
                continue
            for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
                for c in row:
                    formula = c.value if isinstance(c.value, str) and c.value.startswith("=") else ""
                    low = formula.lower()
                    fill = self._normalize_rgb(getattr(c.fill.fgColor, "rgb", None) if c.fill else None)
                    if fill == "FFFF00" and any(r in low for r in refs):
                        c.value = ""
                        c.fill = copy(WHITE_FILL)
                        c.border = copy(NO_BORDER)
                        cleared += 1
        self.log.add(f"Step 38: cleared {cleared} yellow reference cells")

    def _step_39_remove_non_approved_fill_colors_assumptions(self) -> None:
        ws = self._sheet("Assumptions")
        if not ws:
            return
        approved = {"002060", "DCE6F1", "FFFFFF"}
        changed = 0
        for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
            for c in row:
                fill = self._normalize_rgb(getattr(c.fill.fgColor, "rgb", None) if c.fill else None)
                if fill and fill not in approved:
                    c.fill = copy(WHITE_FILL)
                    changed += 1
        self.log.add(f"Step 39: reset {changed} non-approved fills")

    def _step_40_hardcode_residential_parking_spaces(self) -> None:
        ws = self._sheet("Assumptions")
        vws = self._values_sheet("Assumptions")
        if not ws or not vws:
            raise WorkflowError("Step 40 failed: Assumptions sheet missing")
        found = None
        for r in range(1, ws.max_row + 1):
            for c in range(1, ws.max_column + 1):
                v = ws.cell(r, c).value
                if isinstance(v, str) and "residential parking spaces" in v.lower():
                    found = (r, c)
                    break
            if found:
                break
        if not found:
            raise WorkflowError("Step 40 failed: label 'Residential Parking Spaces' not found")
        rr, cc = found
        target = ws.cell(rr, cc + 1)
        if isinstance(target.value, str) and target.value.startswith("="):
            target.value = vws.cell(rr, cc + 1).value
        else:
            target.value = vws.cell(rr, cc + 1).value
        if isinstance(target.value, str) and target.value.startswith("="):
            raise WorkflowError("Step 40 failed: target cell still contains formula after hardcode")
        self.log.add(f"Step 40: hardcoded Assumptions!{get_column_letter(cc + 1)}{rr}")

    def _step_40b_hardcode_residential_parking_rent_stall(self) -> None:
        ws = self._sheet("Assumptions")
        vws = self._values_sheet("Assumptions")
        if not ws or not vws:
            raise WorkflowError("Step 40b failed: Assumptions sheet missing")

        for r in range(1, ws.max_row + 1):
            for c in range(1, ws.max_column):
                val = ws.cell(r, c).value
                label = str(val).strip().lower() if val is not None else ""
                if ("residential parking rent" in label and ("stall" in label or "space" in label)):
                    target = ws.cell(r, c + 1)
                    if isinstance(target.value, str) and target.value.lstrip().startswith("="):
                        target.value = vws.cell(r, c + 1).value
                        self.log.add(f"Step 40b: hardcoded Assumptions!{get_column_letter(c + 1)}{r}")
                    else:
                        self.log.add("Step 40b: target found but right cell was not a formula")
                    return

        self.log.add("Step 40b: label containing 'Residential Parking Rent' + ('Stall'/'Space') not found")
    def _step_41_copy_assumptions_range(self) -> None:
        ws = self._sheet("Assumptions")
        if not ws:
            return
        for r_off in range(0, 5):
            for c_off in range(0, 2):
                src = ws.cell(22 + r_off, column_index_from_string("W") + c_off)
                dst = ws.cell(22 + r_off, column_index_from_string("AC") + c_off)
                if isinstance(src.value, str) and src.value.startswith("="):
                    dst.value = Translator(src.value, origin=src.coordinate).translate_formula(dst.coordinate)
                else:
                    dst.value = src.value
                dst._style = copy(src._style)
                dst.number_format = src.number_format
                dst.protection = copy(src.protection)
                dst.alignment = copy(src.alignment)
        self.log.add("Step 41: copied W22:X26 to AC22:AD26")
    def _step_42_clear_assumptions_range(self) -> None:
        self._clear_to_white("Assumptions", "W4:AA44")
        self.log.add("Step 42: cleared Assumptions W4:AA44")

    def _formula_literal(self, v: Any) -> str:
        if v is None:
            return "0"
        if isinstance(v, bool):
            return "1" if v else "0"
        if isinstance(v, (int, float)):
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
    def _replace_construction_pricing_refs_in_formula(self, formula: str) -> str:
        cp_vws = self._values_sheet("Construction Pricing")
        if not cp_vws:
            return formula

        pattern = re.compile(r"(?:'(?:\[[^\]]+\])?construction pricing'|(?:\[[^\]]+\])?construction pricing)!\$?([A-Z]{1,3})\$?(\d+)", re.IGNORECASE)

        def repl(match: re.Match[str]) -> str:
            col = match.group(1).upper()
            row = int(match.group(2))
            val = cp_vws[f"{col}{row}"].value
            return self._formula_literal(val)

        return pattern.sub(repl, formula)

    def _hardcode_refs_to_sheet(self, target_sheet_name: str) -> None:
        changed = 0
        needle = target_sheet_name.lower()
        for ws in self.wb.worksheets:
            if ws.title.lower() == needle:
                continue
            vws = self._values_sheet(ws.title)
            if not vws:
                continue
            for r in range(1, ws.max_row + 1):
                for c in range(1, ws.max_column + 1):
                    cell = ws.cell(r, c)
                    if isinstance(cell.value, str) and cell.value.startswith("=") and needle in cell.value.lower():
                        addr = f"{get_column_letter(c)}{r}"

                        # Step 24 exception: preserve Cash Flow formulas Q46/Q51 while
                        # replacing only Construction Pricing references with constants.
                        if (
                            needle == "construction pricing"
                            and ws.title.lower() == "cash flow"
                            and addr in {"Q46", "Q51"}
                        ):
                            new_formula = self._replace_construction_pricing_refs_in_formula(cell.value)
                            cell.value = new_formula
                            changed += 1
                            continue

                        cell.value = vws.cell(r, c).value
                        changed += 1
        remaining = self._verify_no_refs(target_sheet_name)
        if remaining > 0:
            raise WorkflowError(f"Hardcode incomplete for {target_sheet_name}: {remaining} refs remain")
        self.log.add(f"Hardcoded refs to {target_sheet_name}: {changed}")
    def _safe_delete_sheet(self, sheet_name: str) -> None:
        ws = self._sheet(sheet_name)
        if not ws:
            self.log.add(f"Delete skipped: sheet '{sheet_name}' not present")
            return
        remaining = self._verify_no_refs(sheet_name)
        if remaining > 0:
            raise WorkflowError(f"Cannot delete '{sheet_name}', {remaining} references remain")
        self.wb.remove(ws)
        if self._sheet(sheet_name):
            raise WorkflowError(f"Failed to delete sheet '{sheet_name}'")
        self.log.add(f"Deleted sheet {sheet_name}")





























    def _scan_bounds(self, ws, sheet_name: str | None = None) -> tuple[int, int]:
        limits = {
            "Executive Summary": (220, 40),
            "Development Summary": (260, 40),
            "Cash Flow": (260, 60),
            "Assumptions": (500, 120),
            "Development": (450, 80),
            "Sale Proceeds": (240, 40),
            "Returns Exhibit": (240, 40),
            "1-Yr Waterfall": (420, 120),
            "3-Yr Waterfall": (420, 120),
            "4-Yr Waterfall": (420, 120),
            "1-Yr Sale": (420, 120),
            "3-Yr Sale": (420, 120),
            "4-Yr Sale": (420, 120),
        }
        lim_row, lim_col = limits.get(sheet_name or ws.title, (600, 160))
        max_row = max(1, min(ws.max_row, lim_row))
        max_col = max(1, min(ws.max_column, lim_col))
        return max_row, max_col

    def _cell_fill_hex(self, cell) -> str | None:
        rgb = None
        try:
            rgb = cell.fill.fgColor.rgb if cell.fill and cell.fill.fgColor else None
        except Exception:
            rgb = None
        return self._normalize_rgb(rgb)

    def _copy_cell(self, src, dst, translate_formula: bool = False) -> None:
        if isinstance(src.value, str) and src.value.startswith("=") and translate_formula:
            try:
                dst.value = Translator(src.value, origin=src.coordinate).translate_formula(dst.coordinate)
            except Exception:
                dst.value = src.value
        else:
            dst.value = src.value
        dst._style = copy(src._style)
        dst.number_format = src.number_format
        dst.protection = copy(src.protection)
        dst.alignment = copy(src.alignment)

    def _set_cell_white_noborder(self, cell) -> None:
        cell.value = None
        cell.border = copy(NO_BORDER)
        cell.fill = copy(WHITE_FILL)

    def _fr_find_label_in_column(self, ws, col: str, label: str) -> int | None:
        max_row, _ = self._scan_bounds(ws, ws.title)
        want = label.strip().lower()
        col_idx = column_index_from_string(col)
        for r in range(1, max_row + 1):
            v = ws.cell(r, col_idx).value
            if isinstance(v, str) and v.strip().lower() == want:
                return r
        return None

    def _fr_find_external_refs_to_addresses(
        self,
        target_sheet_name: str,
        addresses: set[str],
        skip_range: tuple[str, int, int, str, str] | None = None,
    ) -> list[tuple[str, str]]:
        refs: list[tuple[str, str]] = []
        seen: set[tuple[str, str]] = set()

        skip_sheet = None
        skip_r1 = skip_r2 = None
        skip_c1 = skip_c2 = None
        if skip_range is not None:
            skip_sheet, skip_r1, skip_r2, skip_col1, skip_col2 = skip_range
            skip_c1 = column_index_from_string(skip_col1)
            skip_c2 = column_index_from_string(skip_col2)

        patterns = []
        for a in addresses:
            patterns.append(f"'{target_sheet_name}'!{a}".lower())
            patterns.append(f"{target_sheet_name}!{a}".lower())

        target_low = target_sheet_name.lower()
        for ws in self.wb.worksheets:
            max_row, max_col = self._scan_bounds(ws, ws.title)
            for r in range(1, max_row + 1):
                for c in range(1, max_col + 1):
                    cell = ws.cell(r, c)
                    formula = cell.value
                    if not (isinstance(formula, str) and formula.startswith("=")):
                        continue
                    low = formula.lower()
                    if target_low not in low:
                        continue
                    if not any(p in low for p in patterns):
                        continue

                    if (
                        skip_sheet is not None
                        and ws.title.lower() == str(skip_sheet).lower()
                        and skip_r1 <= r <= skip_r2
                        and skip_c1 <= c <= skip_c2
                    ):
                        continue

                    addr = f"{get_column_letter(c)}{r}"
                    key = (ws.title.lower(), addr)
                    if key in seen:
                        continue
                    seen.add(key)
                    refs.append((ws.title, addr))

        return refs

    def _fr_hardcode_refs(self, refs: list[tuple[str, str]]) -> None:
        for ws_name, addr in refs:
            ws = self._sheet(ws_name)
            vws = self._values_sheet(ws_name)
            if ws is None or vws is None:
                continue
            ws[addr].value = vws[addr].value

    def _fr_remove_gp_fees_block(self, sheet_name: str) -> str:
        ws = self._sheet(sheet_name)
        if ws is None:
            return "missing"

        row = self._fr_find_label_in_column(ws, "L", "GP Fees")
        if row is None:
            return "gp_fees_not_found"

        expected = ["GP Fees", "Development Fee", "Construction Management Fee", "GP Total Return"]
        for i, exp in enumerate(expected):
            actual = ws[f"L{row + i}"].value
            if str(actual).strip().lower() != exp.lower():
                return f"label_mismatch@L{row+i}"

        addresses = {f"{col}{r}" for r in range(row, row + 4) for col in ["L", "M", "N", "O", "P", "Q", "R"]}
        refs = self._fr_find_external_refs_to_addresses(sheet_name, addresses, skip_range=(sheet_name, row, row + 3, "L", "R"))
        self._fr_hardcode_refs(refs)

        self._clear_range(ws, row, row + 3, column_index_from_string("L"), column_index_from_string("R"))
        return f"cleared_L{row}:R{row+3}_refs{len(refs)}"

    def _fr_remove_gp_return_block(self, sheet_name: str, block_rows: int, label_offsets: dict[int, str]) -> str:
        ws = self._sheet(sheet_name)
        if ws is None:
            return "missing"

        start_row = self._fr_find_label_in_column(ws, "T", "GP Return on Equity")
        if start_row is None:
            return "gp_return_not_found"

        for offset, expected in label_offsets.items():
            actual = ws[f"T{start_row + offset}"].value
            if str(actual).strip().lower() != expected.lower():
                return f"label_mismatch@T{start_row+offset}"

        end_row = start_row + block_rows - 1
        addresses = {f"{col}{r}" for r in range(start_row, end_row + 1) for col in ["S", "T"]}
        refs = self._fr_find_external_refs_to_addresses(
            sheet_name,
            addresses,
            skip_range=(sheet_name, start_row, end_row, "S", "T"),
        )
        self._fr_hardcode_refs(refs)

        self._clear_range(ws, start_row, end_row, column_index_from_string("S"), column_index_from_string("T"))

        for rr in [start_row, start_row + 1]:
            cell = ws[f"S{rr}"]
            b = copy(cell.border)
            b.left = Side(style="thick", color="FF000000")
            cell.border = b

        return f"cleared_S{start_row}:T{end_row}_refs{len(refs)}"

    def _apply_front_range_variant(self) -> None:
        self._fr_step_1_text_to_black()
        self._fr_step_2_remove_gp_fees_all()
        self._fr_step_3_5_remove_gp_return_blocks()

    def _fr_step_1_text_to_black(self) -> None:
        changed_by_sheet: list[str] = []
        for sheet_name in ["1-Yr Waterfall", "3-Yr Waterfall", "4-Yr Waterfall"]:
            ws = self._sheet(sheet_name)
            if ws is None:
                changed_by_sheet.append(f"{sheet_name}:missing")
                continue
            max_row, max_col = self._scan_bounds(ws, sheet_name)
            changed = 0
            for r in range(1, max_row + 1):
                for c in range(1, max_col + 1):
                    cell = ws.cell(r, c)
                    hx = self._normalize_rgb(cell.font.color.rgb if cell.font and cell.font.color else None)
                    if hx and hx != "FFFFFF" and hx != "000000":
                        cell.font = copy(cell.font)
                        cell.font = cell.font.copy(color="000000")
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

    def _apply_lp_variant(self) -> None:
        sheets = ["1-Yr Waterfall", "3-Yr Waterfall", "4-Yr Waterfall"]
        self._lp_step_1_clear_project_level_subblock(sheets)
        self._lp_step_2_copy_equity_multiple_to_leveraged(sheets)
        self._lp_step_3_link_profit_to_net_cash_flow(sheets)
        self._lp_step_4_link_irr_to_leveraged_irr(sheets)
        self._lp_step_5_link_em_to_leveraged_em(sheets)
        self._lp_step_6_clear_waterfall_terms_block(sheets)
        self._lp_step_7_move_cash_flow_summary_up(sheets)
        self._lp_step_8_delete_empty_rows_29_45(sheets)
        self._lp_step_9_clear_summary_distributions_to_340(sheets)
        self._lp_step_10_link_executive_summary_to_waterfalls()
        self._lp_step_11_rename_waterfall_tabs()
        self._lp_step_12_13_delete_lp_tabs()

    def _lp_find_exact_in_column(self, ws, col: str, label: str) -> int | None:
        max_row, _ = self._scan_bounds(ws, ws.title)
        want = label.strip().lower()
        col_idx = column_index_from_string(col)
        for r in range(1, max_row + 1):
            v = ws.cell(r, col_idx).value
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
            ws = self._sheet(sheet_name)
            if ws is None:
                out.append(f"{sheet_name}:missing")
                continue
            project_row = self._lp_find_exact_in_column(ws, "L", "Project Level")
            if project_row is None:
                out.append(f"{sheet_name}:project_level_not_found")
                continue
            start_row = project_row + 1
            end_row = project_row + 5
            self._clear_range(ws, start_row, end_row, column_index_from_string("L"), column_index_from_string("R"))
            for c in range(column_index_from_string("L"), column_index_from_string("R") + 1):
                cell = ws.cell(start_row, c)
                b = copy(cell.border)
                b.top = Side(style="thin", color="FF000000")
                cell.border = b
            out.append(f"{sheet_name}:L{start_row}:R{end_row}")
        self.log.add("LP Step 1 Project Level clear -> " + ", ".join(out))

    def _copy_row_span(self, ws, src_row: int, dst_row: int, start_col: str, end_col: str) -> None:
        c1 = column_index_from_string(start_col)
        c2 = column_index_from_string(end_col)
        for c in range(c1, c2 + 1):
            self._copy_cell(ws.cell(src_row, c), ws.cell(dst_row, c), translate_formula=True)

    def _lp_step_2_copy_equity_multiple_to_leveraged(self, sheets: list[str]) -> None:
        out = []
        for sheet_name in sheets:
            ws = self._sheet(sheet_name)
            if ws is None:
                out.append(f"{sheet_name}:missing")
                continue

            ucf_row = self._lp_find_exact_in_column(ws, "B", "Unleveraged Cash Flow")
            if ucf_row is None:
                out.append(f"{sheet_name}:ucf_not_found")
                continue
            if str(ws[f"B{ucf_row+1}"].value).strip().lower() != "irr" or str(ws[f"B{ucf_row+2}"].value).strip().lower() != "equity multiple":
                out.append(f"{sheet_name}:test1_failed")
                continue
            source_em_row = ucf_row + 2

            dscr_row = self._lp_find_exact_in_column(ws, "B", "DSCR")
            if dscr_row is None:
                out.append(f"{sheet_name}:dscr_not_found")
                continue
            if str(ws[f"B{dscr_row-2}"].value).strip().lower() != "irr" or not self._lp_is_blank(ws[f"B{dscr_row-1}"].value):
                out.append(f"{sheet_name}:test2_failed")
                continue
            dest_row = dscr_row - 1

            ncf_row = self._lp_find_exact_in_column(ws, "B", "Net Cash Flow")
            if ncf_row is None:
                out.append(f"{sheet_name}:ncf_not_found")
                continue

            self._copy_row_span(ws, source_em_row, dest_row, "B", "D")
            ws[f"D{dest_row}"].value = f'=IFERROR(-SUMIF($F{ncf_row}:$IL{ncf_row},">0")/SUMIF($F{ncf_row}:$IL{ncf_row},"<0"),0)'
            out.append(f"{sheet_name}:B{source_em_row}:D{source_em_row}->B{dest_row}:D{dest_row}")
        self.log.add("LP Step 2 Equity Multiple copy -> " + ", ".join(out))

    def _lp_step_3_link_profit_to_net_cash_flow(self, sheets: list[str]) -> None:
        out = []
        for sheet_name in sheets:
            ws = self._sheet(sheet_name)
            if ws is None:
                out.append(f"{sheet_name}:missing")
                continue
            profit_header_row = self._lp_find_exact_in_column(ws, "P", "Profit")
            ncf_row = self._lp_find_exact_in_column(ws, "B", "Net Cash Flow")
            if profit_header_row is None or ncf_row is None:
                out.append(f"{sheet_name}:missing_profit_or_ncf")
                continue
            target_row = profit_header_row + 2
            ws[f"P{target_row}"].value = f"=E{ncf_row}"
            out.append(f"{sheet_name}:P{target_row}=E{ncf_row}")
        self.log.add("LP Step 3 Profit links -> " + ", ".join(out))

    def _lp_step_4_link_irr_to_leveraged_irr(self, sheets: list[str]) -> None:
        out = []
        for sheet_name in sheets:
            ws = self._sheet(sheet_name)
            if ws is None:
                out.append(f"{sheet_name}:missing")
                continue
            irr_header_row = self._lp_find_exact_in_column(ws, "Q", "IRR")
            irr_row = self._lp_find_exact_in_column(ws, "B", "IRR")
            if irr_header_row is None or irr_row is None:
                out.append(f"{sheet_name}:irr_not_found")
                continue
            target_row = irr_header_row + 2
            ws[f"Q{target_row}"].value = f"=D{irr_row}"
            out.append(f"{sheet_name}:Q{target_row}=D{irr_row}")
        self.log.add("LP Step 4 IRR links -> " + ", ".join(out))

    def _lp_step_5_link_em_to_leveraged_em(self, sheets: list[str]) -> None:
        out = []
        for sheet_name in sheets:
            ws = self._sheet(sheet_name)
            if ws is None:
                out.append(f"{sheet_name}:missing")
                continue
            em_header_row = self._lp_find_exact_in_column(ws, "R", "Equity Multiple")
            em_row = self._lp_find_exact_in_column(ws, "B", "Equity Multiple")
            if em_header_row is None or em_row is None:
                out.append(f"{sheet_name}:em_not_found")
                continue
            target_row = em_header_row + 2
            ws[f"R{target_row}"].value = f"=D{em_row}"
            out.append(f"{sheet_name}:R{target_row}=D{em_row}")
        self.log.add("LP Step 5 Equity Multiple links -> " + ", ".join(out))

    def _lp_step_6_clear_waterfall_terms_block(self, sheets: list[str]) -> None:
        out = []
        for sheet_name in sheets:
            ws = self._sheet(sheet_name)
            if ws is None:
                out.append(f"{sheet_name}:missing")
                continue
            guard = str(ws["B19"].value or "").strip().upper()
            if "CASH FLOW SUMMARY" not in guard:
                out.append(f"{sheet_name}:guard_failed")
                continue
            self._clear_range(ws, 7, 18, column_index_from_string("B"), column_index_from_string("J"))
            out.append(f"{sheet_name}:B7:J18")
        self.log.add("LP Step 6 clear terms -> " + ", ".join(out))

    def _lp_step_7_move_cash_flow_summary_up(self, sheets: list[str]) -> None:
        out = []
        for sheet_name in sheets:
            ws = self._sheet(sheet_name)
            if ws is None:
                out.append(f"{sheet_name}:missing")
                continue
            b19 = str(ws["B19"].value or "").strip().upper()
            b7 = ws["B7"].value
            if "CASH FLOW SUMMARY" not in b19:
                out.append(f"{sheet_name}:guard_b19_failed")
                continue
            if not self._lp_is_blank(b7):
                out.append(f"{sheet_name}:guard_b7_not_blank")
                continue
            ws.move_range("B19:J39", rows=-12, cols=0, translate=True)
            out.append(f"{sheet_name}:moved_B19:J39_to_B7")
        self.log.add("LP Step 7 move summary -> " + ", ".join(out))

    def _lp_step_8_delete_empty_rows_29_45(self, sheets: list[str]) -> None:
        out = []
        for sheet_name in sheets:
            ws = self._sheet(sheet_name)
            if ws is None:
                out.append(f"{sheet_name}:missing")
                continue
            _, max_col = self._scan_bounds(ws, ws.title)
            rows_to_delete: list[int] = []
            for row in range(29, 46):
                empty = True
                for col in range(1, max_col + 1):
                    if not self._lp_is_blank(ws.cell(row, col).value):
                        empty = False
                        break
                if empty:
                    rows_to_delete.append(row)
            for row in reversed(rows_to_delete):
                ws.delete_rows(row, 1)
            out.append(f"{sheet_name}:deleted_{len(rows_to_delete)}")
        self.log.add("LP Step 8 delete empty rows -> " + ", ".join(out))

    def _lp_step_9_clear_summary_distributions_to_340(self, sheets: list[str]) -> None:
        out = []
        for sheet_name in sheets:
            ws = self._sheet(sheet_name)
            if ws is None:
                out.append(f"{sheet_name}:missing")
                continue
            start_row = self._lp_find_exact_in_column(ws, "A", "SUMMARY DISTRIBUTIONS")
            if start_row is None:
                out.append(f"{sheet_name}:summary_distributions_not_found")
                continue
            if start_row > 340:
                out.append(f"{sheet_name}:start_row_gt_340")
                continue
            _, max_col = self._scan_bounds(ws, ws.title)
            self._clear_range(ws, start_row, 340, 1, max_col)
            out.append(f"{sheet_name}:A{start_row}:{get_column_letter(max_col)}340")
        self.log.add("LP Step 9 clear summary distributions -> " + ", ".join(out))

    def _lp_step_10_link_executive_summary_to_waterfalls(self) -> None:
        ws = self._sheet("Executive Summary")
        if ws is None:
            self.log.add("LP Step 10 skipped: missing Executive Summary")
            return
        ws["J16"].value = "='1-Yr Waterfall'!Q22"
        ws["J17"].value = "='1-Yr Waterfall'!R22"
        ws["K16"].value = "='3-Yr Waterfall'!Q22"
        ws["K17"].value = "='3-Yr Waterfall'!R22"
        ws["L16"].value = "='4-Yr Waterfall'!Q22"
        ws["L17"].value = "='4-Yr Waterfall'!R22"
        self.log.add("LP Step 10 linked Executive Summary to waterfall Q22/R22")

    def _replace_sheet_name_refs(self, old_name: str, new_name: str) -> int:
        changed = 0
        q_old = f"'{old_name}'!"
        q_new = f"'{new_name}'!"
        u_old = f"{old_name}!"
        u_new = f"{new_name}!"
        for ws in self.wb.worksheets:
            max_row, max_col = self._scan_bounds(ws, ws.title)
            for r in range(1, max_row + 1):
                for c in range(1, max_col + 1):
                    cell = ws.cell(r, c)
                    v = cell.value
                    if not (isinstance(v, str) and v.startswith("=")):
                        continue
                    nv = v.replace(q_old, q_new).replace(u_old, u_new)
                    if nv != v:
                        cell.value = nv
                        changed += 1
        return changed

    def _lp_step_11_rename_waterfall_tabs(self) -> None:
        rename_map = {
            "1-Yr Waterfall": "1-Yr Sale",
            "3-Yr Waterfall": "3-Yr Sale",
            "4-Yr Waterfall": "4-Yr Sale",
        }
        out = []
        for old_name, new_name in rename_map.items():
            old_ws = self._sheet(old_name)
            if old_ws is None:
                out.append(f"{old_name}:missing")
                continue
            if self._sheet(new_name) is not None:
                out.append(f"{old_name}:target_exists")
                continue
            old_ws.title = new_name
            rewrites = self._replace_sheet_name_refs(old_name, new_name)
            out.append(f"{old_name}->{new_name}:refs_rewritten_{rewrites}")
        self.log.add("LP Step 11 rename tabs -> " + ", ".join(out))

    def _lp_step_12_13_delete_lp_tabs(self) -> None:
        self._safe_delete_sheet("FR Waterfall Analysis")
        self._safe_delete_sheet("Returns Exhibit")
        self.log.add("LP Steps 12-13 deletion checks completed")


    def _addr(self, row: int, col: int) -> str:
        return f"{get_column_letter(col)}{row}"

    def _lender_used_bounds(self, ws) -> tuple[int, int, int, int]:
        max_row, max_col = self._scan_bounds(ws, ws.title)
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
        want = label.strip().lower()
        hits: list[tuple[int, int]] = []
        for r in range(row1, row2 + 1):
            for c in range(col1, col2 + 1):
                v = ws.cell(r, c).value
                if isinstance(v, str) and v.strip().lower() == want:
                    hits.append((r, c))
        return hits

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

    def _lender_step_1_hardcode_keepers(self, keepers: list[str]) -> None:
        results: list[str] = []
        for name in keepers:
            ws = self._sheet(name)
            vws = self._values_sheet(name)
            if ws is None or vws is None:
                results.append(f"{name}:missing")
                continue
            row1, col1, row2, col2 = self._lender_used_bounds(ws)
            replaced = 0
            errors = 0
            for r in range(row1, row2 + 1):
                for c in range(col1, col2 + 1):
                    cell = ws.cell(r, c)
                    if not (isinstance(cell.value, str) and cell.value.startswith("=")):
                        continue
                    raw = vws.cell(r, c).value
                    if isinstance(raw, str) and raw.startswith("#"):
                        cell.value = 0
                        errors += 1
                    else:
                        cell.value = raw
                    replaced += 1
            results.append(f"{name}:replaced_{replaced}:errors_{errors}")
        self.log.add("Lender Step 1 hardcode -> " + ", ".join(results))

    def _lender_step_2_safety_and_delete_nonkeepers(self, keepers: list[str]) -> None:
        all_names = [ws.title for ws in self.wb.worksheets]
        missing = [k for k in keepers if k not in all_names]
        if missing:
            raise WorkflowError(f"Lender Step 2 failed: missing keeper sheets {missing}")

        to_delete = [n for n in all_names if n not in keepers]
        live_formula_count = 0
        cross_ref_count = 0

        for name in keepers:
            ws = self._sheet(name)
            if ws is None:
                continue
            row1, col1, row2, col2 = self._lender_used_bounds(ws)
            for r in range(row1, row2 + 1):
                for c in range(col1, col2 + 1):
                    f = ws.cell(r, c).value
                    if not (isinstance(f, str) and f.startswith("=")):
                        continue
                    live_formula_count += 1
                    low = f.lower()
                    if any(d.lower() in low for d in to_delete):
                        cross_ref_count += 1

        if live_formula_count > 0:
            raise WorkflowError(f"Lender Step 2 TEST A failed: {live_formula_count} live formulas remain in keeper sheets")
        if cross_ref_count > 0:
            raise WorkflowError(f"Lender Step 2 TEST B failed: {cross_ref_count} refs to to-be-deleted sheets in keeper sheets")

        deleted: list[str] = []
        for name in to_delete:
            ws = self._sheet(name)
            if ws is None:
                continue
            self.wb.remove(ws)
            deleted.append(name)

        remaining = sorted([ws.title for ws in self.wb.worksheets])
        if sorted(keepers) != remaining:
            raise WorkflowError(f"Lender Step 2 failed final verification. Remaining={remaining}")
        self.log.add("Lender Step 2 delete non-keepers -> " + ", ".join(deleted))

    def _lender_step_3_conditional_delete_cashflow_rows(self) -> None:
        ws = self._sheet("Cash Flow")
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
            f_val = ws[f"F{row}"].value
            if self._lender_value_zeroish(f_val):
                rows_to_delete.append(row)
            else:
                kept.append(f"{label}:{row}:{f_val}")
        for row in sorted(set(rows_to_delete), reverse=True):
            ws.delete_rows(row, 1)
        self.log.add(f"Lender Step 3 row deletes -> deleted={sorted(set(rows_to_delete), reverse=True)} kept={kept}")

    def _lender_step_4_clear_below_noi(self) -> None:
        ws = self._sheet("Cash Flow")
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
        self._clear_range(ws, start_row, row2, col1, col2)
        self.log.add(f"Lender Step 4 cleared below NOI -> {self._addr(start_row,col1)}:{self._addr(row2,col2)}")

    def _lender_step_5_clear_below_blue_col_b_assumptions(self) -> None:
        ws = self._sheet("Assumptions")
        if ws is None:
            raise WorkflowError("Lender Step 5 failed: Assumptions sheet missing")
        row1, col1, row2, col2 = self._lender_used_bounds(ws)
        found_row = None
        for r in range(max(46, row1), row2 + 1):
            if self._cell_fill_hex(ws.cell(r, 2)) == "002060":
                found_row = r
                break
        if found_row is None:
            raise WorkflowError("Lender Step 5 failed: no #002060 cell found in col B below row 45")
        self._clear_range(ws, found_row, row2, 2, col2)
        self.log.add(f"Lender Step 5 cleared assumptions below blue B -> B{found_row}:{self._addr(row2,col2)}")

    def _lender_step_6_clear_assumptions_blue_header_band(self) -> None:
        ws = self._sheet("Assumptions")
        if ws is None:
            raise WorkflowError("Lender Step 6 failed: Assumptions sheet missing")
        hits = [h for h in self._lender_find_label_positions(ws, "ASSUMPTIONS") if h[1] == 3]
        if not hits:
            raise WorkflowError("Lender Step 6 failed: ASSUMPTIONS not found in col C")
        anchor_row, anchor_col = hits[0]
        row1, col1, row2, col2 = self._lender_used_bounds(ws)
        end_col = anchor_col
        for c in range(anchor_col, col2 + 1):
            if self._cell_fill_hex(ws.cell(anchor_row, c)) == "002060":
                end_col = c
            else:
                break
        self._clear_range(ws, anchor_row, row2, anchor_col, end_col)
        self.log.add(f"Lender Step 6 cleared assumptions header band -> {self._addr(anchor_row,anchor_col)}:{self._addr(row2,end_col)}")

    def _shift_row_segment_left(self, ws, row: int, start_col: int, end_col: int, limit_col: int) -> None:
        count = end_col - start_col + 1
        if count <= 0:
            return
        for c in range(start_col, limit_col - count + 1):
            src = ws.cell(row, c + count)
            dst = ws.cell(row, c)
            self._copy_cell(src, dst, translate_formula=True)
        for c in range(limit_col - count + 1, limit_col + 1):
            self._set_cell_white_noborder(ws.cell(row, c))

    def _shift_range_left(self, ws, row_start: int, row_end: int, col_start: int, col_end: int, limit_col: int) -> None:
        for r in range(row_start, row_end + 1):
            self._shift_row_segment_left(ws, r, col_start, col_end, limit_col)

    def _lender_step_7_delete_current_year_columns(self) -> None:
        ws = self._sheet("Assumptions")
        if ws is None:
            self.log.add("Lender Step 7 skipped: Assumptions sheet missing")
            return
        p1 = self._lender_find_label_positions(ws, "Current Year Per Bed Input")
        p2 = self._lender_find_label_positions(ws, "Current Year Rent PSF")
        if not p1 or not p2:
            self.log.add("Lender Step 7 skipped: required labels not found")
            return
        r1, c1 = p1[0]
        r2, c2 = p2[0]
        if r1 != r2:
            self.log.add("Lender Step 7 skipped: labels not in same row")
            return
        label_row = r1
        left_col, right_col = sorted((c1, c2))

        blue_row = None
        for r in range(label_row - 1, 0, -1):
            all_blue = True
            for c in range(left_col, right_col + 1):
                if self._cell_fill_hex(ws.cell(r, c)) != "002060":
                    all_blue = False
                    break
            if all_blue:
                blue_row = r
                break
        if blue_row is None:
            self.log.add("Lender Step 7 skipped: blue header row above labels not found")
            return

        _, _, row2, max_col = self._lender_used_bounds(ws)
        light_hits = 0
        bottom_row = None
        for r in range(label_row, row2 + 1):
            if self._cell_fill_hex(ws.cell(r, left_col)) == "DCE6F1":
                light_hits += 1
                if light_hits == 2:
                    bottom_row = r
                    break
        if bottom_row is None:
            self.log.add("Lender Step 7 skipped: second #DCE6F1 boundary not found")
            return

        for rr in [bottom_row + 1, bottom_row + 2]:
            if rr > row2:
                continue
            for c in range(left_col, right_col + 1):
                if not self._lp_is_blank(ws.cell(rr, c).value):
                    self.log.add("Lender Step 7 warning: expected blank spacer rows after boundary check failed")

        self._shift_range_left(ws, blue_row, bottom_row, left_col, right_col, max_col)
        self.log.add(f"Lender Step 7 deleted current-year cols -> {self._addr(blue_row,left_col)}:{self._addr(bottom_row,right_col)}")

    def _lender_step_8_delete_current_rent_column(self) -> None:
        ws = self._sheet("Assumptions")
        if ws is None:
            self.log.add("Lender Step 8 skipped: Assumptions sheet missing")
            return
        hits = self._lender_find_label_positions(ws, "Current Rent/Yr")
        if not hits:
            self.log.add("Lender Step 8 skipped: Current Rent/Yr not found")
            return
        label_row, col = hits[0]

        blue_row = None
        for r in range(label_row - 1, 0, -1):
            if self._cell_fill_hex(ws.cell(r, col)) == "002060":
                blue_row = r
                break
        if blue_row is None:
            self.log.add("Lender Step 8 skipped: blue header row above Current Rent/Yr not found")
            return

        _, _, row2, max_col = self._lender_used_bounds(ws)
        bottom_row = None
        for r in range(label_row, row2 + 1):
            if self._cell_fill_hex(ws.cell(r, col)) == "DCE6F1":
                bottom_row = r
                break
        if bottom_row is None:
            self.log.add("Lender Step 8 skipped: first #DCE6F1 boundary not found")
            return

        self._shift_range_left(ws, blue_row, bottom_row, col, col, max_col)
        self.log.add(f"Lender Step 8 deleted Current Rent/Yr col -> {self._addr(blue_row,col)}:{self._addr(bottom_row,col)}")

    def _lender_step_9_clear_right_of_total_sf(self) -> None:
        ws = self._sheet("Assumptions")
        if ws is None:
            self.log.add("Lender Step 9 skipped: Assumptions sheet missing")
            return
        hits = self._lender_find_label_positions(ws, "Total SF")
        if not hits:
            self.log.add("Lender Step 9 skipped: Total SF not found")
            return
        hits = sorted(hits, key=lambda x: x[0])
        label_row, label_col = hits[0]

        top_row = None
        for r in range(label_row, 0, -1):
            if self._cell_fill_hex(ws.cell(r, label_col)) == "002060":
                top_row = r
                break
        if top_row is None:
            self.log.add("Lender Step 9 skipped: top #002060 boundary not found")
            return

        _, _, last_row, last_col = self._lender_used_bounds(ws)
        light_hits = 0
        bottom_row = None
        for r in range(label_row, last_row + 1):
            if self._cell_fill_hex(ws.cell(r, label_col)) == "DCE6F1":
                light_hits += 1
                if light_hits == 2:
                    bottom_row = r
                    break
        if bottom_row is None:
            self.log.add("Lender Step 9 skipped: second #DCE6F1 boundary not found")
            return

        start_col = label_col + 1
        if start_col > last_col:
            self.log.add("Lender Step 9 skipped: no columns to the right of Total SF")
            return

        self._clear_range(ws, top_row, bottom_row, start_col, last_col)
        self.log.add(f"Lender Step 9 cleared right of Total SF -> {self._addr(top_row,start_col)}:{self._addr(bottom_row,last_col)}")
