from __future__ import annotations

from pathlib import Path
import json
import os
from uuid import uuid4

from flask import Flask, render_template, request, send_from_directory, abort, redirect, url_for

from app.engine import WorkflowError
from app.engine_factory import get_runner
from app.report import build_side_by_side_report
from app.validator import validate_workbook

BASE = Path(__file__).resolve().parent.parent
RUNS_ROOT = BASE / "runs"
WEB_UPLOADS = RUNS_ROOT / "web-submissions"
WEB_UPLOADS.mkdir(parents=True, exist_ok=True)
ASSERTIONS_BASE = BASE / "config" / "assertions" / "VERVE-Proforma-Cleaner-v1.1-Assertions.json"
ASSERTIONS_LP = BASE / "config" / "assertions" / "VERVE-Proforma-Cleaner-v1.1-LP-Assertions.json"
ASSERTIONS_LENDER = BASE / "config" / "assertions" / "VERVE-Proforma-Cleaner-v1.1-Lender-Assertions.json"

app = Flask(
    __name__,
    template_folder=str(Path(__file__).resolve().parent / "templates"),
    static_folder=str(Path(__file__).resolve().parent / "static"),
)
app.secret_key = os.getenv("VERVE_APP_SECRET", "").strip() or os.urandom(32).hex()


def login_required(view_func):
    return view_func


@app.route("/login", methods=["GET", "POST"])
def login():
    return redirect(url_for("index"))


@app.route("/logout", methods=["POST"])
def logout():
    return redirect(url_for("index"))


def build_run_diagnostics(run_log: list[str], validation: dict, report: dict) -> dict:
    log_lines = run_log or []
    validation_errors = list(validation.get("errors", [])) if isinstance(validation, dict) else []

    warn_tokens = (
        "warning",
        "failed",
        "error",
        "abort",
        "aborted",
        "skipped",
        "not found",
        "missing",
        "condition not met",
        "still present",
        "incomplete",
    )

    flagged = [line for line in log_lines if any(tok in line.lower() for tok in warn_tokens)]

    expected_remaining = 0
    try:
        expected_remaining = len(report.get("sheet_diff", {}).get("expected_removed_still_present", []))
    except Exception:
        expected_remaining = 0

    hints: list[str] = []
    low = "\n".join(log_lines).lower()
    if "condition not met" in low:
        hints.append("One or more conditional rules were skipped because source values/labels did not match expected patterns.")
    if "not found" in low or "missing" in low:
        hints.append("At least one expected label/sheet/range was not found. This often indicates a layout variation in the uploaded proforma.")
    if expected_remaining > 0:
        hints.append("Some sheets expected to be deleted are still present; check hardcode/delete sequence for that option.")
    if validation_errors:
        hints.append("Validation reported hard failures. Review those first before trusting output metrics.")
    if not hints:
        hints.append("No high-risk diagnostics detected in this run.")

    return {
        "validation_error_count": len(validation_errors),
        "flagged_log_count": len(flagged),
        "expected_targets_remaining": expected_remaining,
        "flagged_log_lines": flagged[:25],
        "validation_errors": validation_errors,
        "hints": hints,
    }


@app.route("/download/<run_id>/<path:filename>")
@login_required
def download_output(run_id: str, filename: str):
    run_dir = (WEB_UPLOADS / run_id).resolve()
    if not run_dir.exists() or not run_dir.is_dir():
        abort(404)

    target = (run_dir / filename).resolve()
    if run_dir not in target.parents and target != run_dir:
        abort(404)
    if not target.exists() or not target.is_file():
        abort(404)

    return send_from_directory(
        directory=str(run_dir),
        path=filename,
        as_attachment=True,
        download_name=filename,
        mimetype="application/vnd.ms-excel.sheet.macroEnabled.12",
    )


@app.route("/download-diagnostics/<run_id>/<path:filename>")
@login_required
def download_diagnostics(run_id: str, filename: str):
    run_dir = (WEB_UPLOADS / run_id).resolve()
    if not run_dir.exists() or not run_dir.is_dir():
        abort(404)

    target = (run_dir / filename).resolve()
    if run_dir not in target.parents and target != run_dir:
        abort(404)
    if not target.exists() or not target.is_file():
        abort(404)

    return send_from_directory(
        directory=str(run_dir),
        path=filename,
        as_attachment=True,
        download_name=filename,
        mimetype="application/json",
    )


@app.route("/", methods=["GET", "POST"])
@login_required
def index():
    context = {
        "result": None,
        "error": None,
        "form": {
            "project_type": "VERVE",
            "market": "",
            "tax_abatement": "no",
            "option": "base",
            "has_additional_tabs": "no",
            "additional_sheet_names": [],
        },
    }

    if request.method == "POST":
        form_state = context["form"]
        file = request.files.get("workbook")

        option = (request.form.get("option") or "base").strip().lower()
        if option not in {"base", "front-range", "lp", "lender"}:
            option = "base"
        form_state["option"] = option

        project_type = (request.form.get("project_type") or "VERVE").strip().upper()
        if project_type not in {"VERVE", "EVER", "LOCAL"}:
            project_type = "VERVE"
        form_state["project_type"] = project_type

        market = (request.form.get("market") or "").strip()
        form_state["market"] = market
        if not market:
            context["error"] = "Market is required."
            return render_template("index.html", **context)

        tax_abatement = (request.form.get("tax_abatement") or "no").strip().lower()
        if tax_abatement not in {"yes", "no"}:
            tax_abatement = "no"
        form_state["tax_abatement"] = tax_abatement

        has_additional_tabs = (request.form.get("has_additional_tabs") or "no").strip().lower()
        if has_additional_tabs not in {"yes", "no"}:
            has_additional_tabs = "no"

        additional_sheet_names: list[str] = []
        if has_additional_tabs == "yes":
            seen: set[str] = set()
            for raw in request.form.getlist("additional_sheet_name"):
                name = (raw or "").strip()
                if not name:
                    continue
                key = name.lower()
                if key in seen:
                    continue
                additional_sheet_names.append(name)
                seen.add(key)
            if not additional_sheet_names:
                context["error"] = "Add at least one sheet name when 'Additional tabs added to template?' is set to Yes."
                form_state["has_additional_tabs"] = has_additional_tabs
                form_state["additional_sheet_names"] = additional_sheet_names
                return render_template("index.html", **context)

        form_state["has_additional_tabs"] = has_additional_tabs
        form_state["additional_sheet_names"] = additional_sheet_names

        if not file or not file.filename:
            context["error"] = "Select an .xlsm or .xlsx file."
            return render_template("index.html", **context)
        if not (file.filename.lower().endswith(".xlsm") or file.filename.lower().endswith(".xlsx")):
            context["error"] = "Only .xlsm/.xlsx files are supported."
            return render_template("index.html", **context)

        run_id = uuid4().hex[:8]
        run_dir = WEB_UPLOADS / run_id
        run_dir.mkdir(parents=True, exist_ok=True)
        input_path = run_dir / file.filename
        file.save(input_path)

        try:
            previous_tax_setting = os.environ.get("VERVE_TAX_ABATEMENT")
            previous_project_type = os.environ.get("VERVE_PROJECT_TYPE")
            previous_market = os.environ.get("VERVE_MARKET")
            previous_additional_tabs = os.environ.get("VERVE_ADDITIONAL_DELETE_TABS")
            os.environ["VERVE_TAX_ABATEMENT"] = tax_abatement
            os.environ["VERVE_PROJECT_TYPE"] = project_type
            os.environ["VERVE_MARKET"] = market
            os.environ["VERVE_ADDITIONAL_DELETE_TABS"] = json.dumps(additional_sheet_names)
            try:
                runner = get_runner(input_path, option=option)
                run_result = runner.run(output_dir=run_dir)
            finally:
                if previous_tax_setting is None:
                    os.environ.pop("VERVE_TAX_ABATEMENT", None)
                else:
                    os.environ["VERVE_TAX_ABATEMENT"] = previous_tax_setting
                if previous_project_type is None:
                    os.environ.pop("VERVE_PROJECT_TYPE", None)
                else:
                    os.environ["VERVE_PROJECT_TYPE"] = previous_project_type
                if previous_market is None:
                    os.environ.pop("VERVE_MARKET", None)
                else:
                    os.environ["VERVE_MARKET"] = previous_market
                if previous_additional_tabs is None:
                    os.environ.pop("VERVE_ADDITIONAL_DELETE_TABS", None)
                else:
                    os.environ["VERVE_ADDITIONAL_DELETE_TABS"] = previous_additional_tabs

            output_file = Path(run_result["output_file"])
            assertions = ASSERTIONS_BASE
            if option == "lp":
                assertions = ASSERTIONS_LP
            elif option == "lender":
                assertions = ASSERTIONS_LENDER
            validation = validate_workbook(assertions, output_file)
            report = build_side_by_side_report(input_path, output_file)
            diagnostics = build_run_diagnostics(run_result["log"], validation, report)
            diagnostics_filename = "run_diagnostics.json"
            diagnostics_path = run_dir / diagnostics_filename
            diagnostics_payload = {
                "run_id": run_id,
                "option": option,
                "project_type": project_type,
                "market": market,
                "tax_abatement": tax_abatement,
                "has_additional_tabs": has_additional_tabs,
                "additional_sheet_names": additional_sheet_names,
                "input_filename": input_path.name,
                "output_filename": output_file.name,
                "diagnostics_filename": diagnostics_filename,
                "validation": validation,
                "diagnostics": diagnostics,
                "report_summary": {
                    "sheet_diff": report.get("sheet_diff", {}),
                    "changes": report.get("changes", {}),
                },
                "log": run_result["log"],
            }
            diagnostics_path.write_text(
                json.dumps(diagnostics_payload, indent=2, default=str),
                encoding="utf-8",
            )
            context["result"] = {
                "run_id": run_id,
                "option": option,
                "project_type": project_type,
                "market": market,
                "tax_abatement": tax_abatement,
                "has_additional_tabs": has_additional_tabs,
                "additional_sheet_names": additional_sheet_names,
                "input_filename": input_path.name,
                "output_filename": output_file.name,
                "diagnostics_filename": diagnostics_filename,
                "log": run_result["log"],
                "validation": validation,
                "report": report,
                "diagnostics": diagnostics,
            }
        except WorkflowError as exc:
            context["error"] = f"Workflow failed: {exc}"
        except Exception as exc:
            context["error"] = f"Unexpected error: {exc}"

    return render_template("index.html", **context)


if __name__ == "__main__":
    app.run(debug=True, port=5050)