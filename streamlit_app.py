from __future__ import annotations

import json
import os
import re
from pathlib import Path
from uuid import uuid4

import streamlit as st

from app.engine import WorkflowError
from app.engine_factory import get_runner
from app.report import build_side_by_side_report
from app.validator import validate_workbook

BASE = Path(__file__).resolve().parent
RUNS_ROOT = BASE / "runs"
STREAMLIT_UPLOADS = RUNS_ROOT / "web-submissions-streamlit"
STREAMLIT_UPLOADS.mkdir(parents=True, exist_ok=True)

ASSERTIONS_BASE = BASE / "VERVE-Proforma-Cleaner-v1.1-Assertions.json"
ASSERTIONS_LP = BASE / "VERVE-Proforma-Cleaner-v1.1-LP-Assertions.json"
ASSERTIONS_LENDER = BASE / "VERVE-Proforma-Cleaner-v1.1-Lender-Assertions.json"


def _build_run_diagnostics(run_log: list[str], validation: dict, report: dict) -> dict:
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
    expected_remaining = len(report.get("sheet_diff", {}).get("expected_removed_still_present", []))

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


def _require_login() -> None:
    st.set_page_config(page_title="MrClean's External Model Generator", layout="centered")

    app_user = os.getenv("VERVE_APP_USER", "").strip()
    app_password = os.getenv("VERVE_APP_PASSWORD", "").strip()

    if not app_user or not app_password:
        st.error("Authentication is not configured. Set VERVE_APP_USER and VERVE_APP_PASSWORD in deployment secrets.")
        st.stop()

    if "authenticated" not in st.session_state:
        st.session_state.authenticated = False

    if st.session_state.authenticated:
        return

    st.title("MrClean's External Model Generator")
    st.subheader("Sign in")
    with st.form("login"):
        username = st.text_input("Username")
        password = st.text_input("Password", type="password")
        submitted = st.form_submit_button("Sign In")

    if submitted:
        if username.strip() == app_user and password == app_password:
            st.session_state.authenticated = True
            st.rerun()
        st.error("Invalid username or password.")

    st.stop()


def _market_slug(v: str) -> str:
    raw = re.sub(r"[^A-Za-z0-9]+", "_", (v or "").strip().upper())
    raw = re.sub(r"_+", "_", raw).strip("_")
    return raw or "MARKET"


def _assertions_for_option(option: str) -> Path:
    if option == "lp":
        return ASSERTIONS_LP
    if option == "lender":
        return ASSERTIONS_LENDER
    return ASSERTIONS_BASE


def main() -> None:
    _require_login()

    st.title("MrClean's External Model Generator")
    st.caption("Upload one raw .xlsm/.xlsx file and run the workflow.")

    if st.button("Log Out"):
        st.session_state.authenticated = False
        st.rerun()

    with st.form("run-workflow"):
        workbook = st.file_uploader("Workbook", type=["xlsm", "xlsx"])
        project_type = st.selectbox("Project Type", options=["VERVE", "EVER", "LOCAL"], index=0)
        market = st.text_input("Market", placeholder="e.g. GAINESVILLE")
        tax_abatement = st.selectbox("Tax Abatement", options=["no", "yes"], index=0)
        option = st.selectbox("Workflow Option", options=["base", "front-range", "lp", "lender"], index=0)
        submitted = st.form_submit_button("Run Workflow")

    if not submitted:
        return

    if workbook is None:
        st.error("Select an .xlsm or .xlsx file.")
        return
    if not market.strip():
        st.error("Market is required.")
        return


    run_id = uuid4().hex[:8]
    run_dir = STREAMLIT_UPLOADS / run_id
    run_dir.mkdir(parents=True, exist_ok=True)
    input_path = run_dir / workbook.name
    input_path.write_bytes(workbook.getbuffer())

    env_before = {
        "VERVE_TAX_ABATEMENT": os.environ.get("VERVE_TAX_ABATEMENT"),
        "VERVE_PROJECT_TYPE": os.environ.get("VERVE_PROJECT_TYPE"),
        "VERVE_MARKET": os.environ.get("VERVE_MARKET"),
        "VERVE_FORCE_OPENPYXL": os.environ.get("VERVE_FORCE_OPENPYXL"),
    }

    os.environ["VERVE_TAX_ABATEMENT"] = tax_abatement
    os.environ["VERVE_PROJECT_TYPE"] = project_type
    os.environ["VERVE_MARKET"] = market
    os.environ["VERVE_FORCE_OPENPYXL"] = "1"

    progress = st.progress(0, text="Preparing workbook...")
    for p, msg in [
        (5, "Boosting Proforma Returns..."),
        (10, "Hard Coding UROC..."),
        (15, "Lowering Costs..."),
        (20, "Reducing Inflation..."),
        (25, "Boosting Rents..."),
    ]:
        progress.progress(p, text=msg)

    try:
        runner = get_runner(input_path, option=option)
        run_result = runner.run(output_dir=run_dir)
        output_file = Path(run_result["output_file"])
        validation = validate_workbook(_assertions_for_option(option), output_file)
        report = build_side_by_side_report(input_path, output_file)
        diagnostics = _build_run_diagnostics(run_result["log"], validation, report)

        diagnostics_payload = {
            "run_id": run_id,
            "option": option,
            "project_type": project_type,
            "market": _market_slug(market),
            "tax_abatement": tax_abatement,
            "input_filename": input_path.name,
            "output_filename": output_file.name,
            "validation": validation,
            "diagnostics": diagnostics,
            "report_summary": {
                "sheet_diff": report.get("sheet_diff", {}),
                "changes": report.get("changes", {}),
            },
            "log": run_result["log"],
        }
        diagnostics_json = json.dumps(diagnostics_payload, indent=2, default=str)
        (run_dir / "run_diagnostics.json").write_text(diagnostics_json, encoding="utf-8")

        progress.progress(100, text="Workflow complete")
        st.success(f"Run complete: {output_file.name}")

        st.download_button(
            "Download Finished Workbook",
            data=output_file.read_bytes(),
            file_name=output_file.name,
            mime="application/vnd.ms-excel.sheet.macroEnabled.12",
        )
        st.download_button(
            "Download Diagnostics JSON",
            data=diagnostics_json.encode("utf-8"),
            file_name="run_diagnostics.json",
            mime="application/json",
        )

        c1, c2, c3 = st.columns(3)
        c1.metric("Validation errors", diagnostics["validation_error_count"])
        c2.metric("Flagged log lines", diagnostics["flagged_log_count"])
        c3.metric("Expected deletions still present", diagnostics["expected_targets_remaining"])

        with st.expander("Triage Hints", expanded=True):
            for h in diagnostics["hints"]:
                st.write(f"- {h}")

        with st.expander("Execution Log"):
            st.code("\n".join(run_result["log"]))

    except WorkflowError as exc:
        st.error(f"Workflow failed: {exc}")
    except Exception as exc:
        st.error(f"Unexpected error: {exc}")
    finally:
        for key, old in env_before.items():
            if old is None:
                os.environ.pop(key, None)
            else:
                os.environ[key] = old


if __name__ == "__main__":
    main()
