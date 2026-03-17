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


def _inject_theme() -> None:
    st.markdown(
        """
<style>
@import url('https://fonts.googleapis.com/css2?family=Libre+Baskerville:wght@400;700&family=Source+Sans+3:wght@400;500;600;700&display=swap');

:root {
  --slate-gray: #2b2825;
  --everest-green: #16352e;
  --birch: #a95818;
  --brown: #512213;
  --beige: #f7f1e3;
  --lime-green: #c1d100;
  --bg: #2b2825;
  --panel: #221f1c;
  --line: #3a352f;
  --chip: #2d2925;
  --muted: #d9cfbe;
  --ink: #f7f1e3;
}

[data-testid="stAppViewContainer"],
[data-testid="stAppViewContainer"] > .main,
.stApp {
  background:
    radial-gradient(1200px 460px at 100% -90px, #35302b 0%, rgba(53, 48, 43, 0) 58%),
    linear-gradient(180deg, #2a2623 0%, var(--bg) 100%) !important;
  color: var(--ink) !important;
}

[data-testid="stHeader"] {
  background: transparent !important;
}

.block-container {
  max-width: 1120px;
  padding-top: 0.8rem;
  padding-bottom: 2rem;
}

html, body, [class*="css"], p, div, span, label {
  font-family: "Source Sans 3", "Segoe UI", Calibri, sans-serif;
  color: var(--ink);
}

.brand-wrap {
  margin-bottom: 1rem;
}

.brand-mark {
  font-family: "Libre Baskerville", Georgia, serif;
  font-size: clamp(2.2rem, 8vw, 4.2rem);
  line-height: 0.9;
  color: var(--lime-green);
  text-transform: lowercase;
}

.brand-rule {
  margin-top: 0.4rem;
  width: min(460px, 84vw);
  border-top: 5px solid var(--lime-green);
}

.brand-title {
  margin-top: 0.7rem;
  font-family: "Libre Baskerville", Georgia, serif;
  color: var(--lime-green);
  font-size: clamp(1.4rem, 3vw, 2.15rem);
}

.brand-subtitle {
  margin-top: 0.45rem;
  color: var(--muted);
  font-size: 1rem;
}

h1, h2, h3 {
  font-family: "Libre Baskerville", Georgia, serif;
  color: var(--lime-green);
}

[data-testid="stForm"] {
  background: var(--panel);
  border: 1px solid var(--line);
  border-radius: 12px;
  padding: 16px;
  box-shadow: 0 8px 22px rgba(0, 0, 0, 0.25);
}

div[data-baseweb="input"] > div,
div[data-baseweb="base-input"] > div,
div[data-baseweb="select"] > div,
[data-testid="stFileUploaderDropzone"],
input[type="text"],
input[type="password"],
textarea {
  background: #2a2622 !important;
  border: 1px solid #4a443d !important;
  color: var(--ink) !important;
  border-radius: 8px !important;
}

div[data-baseweb="select"] * {
  color: var(--ink) !important;
}

div[data-baseweb="select"] svg {
  fill: var(--ink) !important;
}

[data-testid="stFileUploaderDropzone"] {
  min-height: 88px;
}

[data-testid="stFileUploaderDropzone"] * {
  color: var(--ink) !important;
}

label, .stSelectbox label, .stTextInput label, .stFileUploader label {
  color: var(--ink) !important;
  font-weight: 600 !important;
}

.stButton > button {
  border: 0 !important;
  border-radius: 8px !important;
  background: linear-gradient(180deg, #d4df3f 0%, var(--lime-green) 100%) !important;
  color: var(--everest-green) !important;
  font-weight: 700 !important;
  font-size: 0.95rem !important;
  padding: 0.56rem 0.95rem !important;
}

.stButton > button[kind="secondary"] {
  background: #3a352f !important;
  color: var(--ink) !important;
  border: 1px solid #585046 !important;
}

[data-testid="metric-container"] {
  background: var(--chip);
  border: 1px solid #453f37;
  border-radius: 8px;
  padding: 8px 10px;
}

details {
  background: var(--panel);
  border: 1px solid var(--line);
  border-radius: 10px;
  padding: 6px 10px;
}

.stCodeBlock pre {
  background: #2c2823 !important;
  border: 1px solid #4a433b !important;
  border-radius: 8px !important;
}

.stDownloadButton > button {
  border: 0 !important;
  border-radius: 8px !important;
  font-weight: 700 !important;
  background: linear-gradient(180deg, #d4df3f 0%, var(--lime-green) 100%) !important;
  color: var(--everest-green) !important;
}

[data-testid="stProgressBar"] > div {
  background: var(--chip) !important;
  border: 1px solid #4a433b !important;
  border-radius: 999px !important;
}

[data-testid="stProgressBar"] > div > div {
  background: linear-gradient(90deg, #d4df3f 0%, var(--lime-green) 100%) !important;
}

[data-testid="stAlert"] {
  background: #2d2925 !important;
  border: 1px solid #453f37 !important;
  color: var(--ink) !important;
}

a {
  color: var(--lime-green);
}

@media (max-width: 720px) {
  .brand-mark {
    font-size: 3rem;
  }
  .brand-title {
    font-size: 1.2rem;
  }
  .brand-rule {
    width: 300px;
  }
}
</style>
""",
        unsafe_allow_html=True,
    )


def _render_brand_header(subtitle: str) -> None:
    st.markdown(
        f"""
<div class="brand-wrap">
  <div class="brand-mark">subtext</div>
  <div class="brand-rule"></div>
  <div class="brand-title">MrClean's External Model Generator</div>
  <div class="brand-subtitle">{subtitle}</div>
</div>
""",
        unsafe_allow_html=True,
    )


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
    st.set_page_config(page_title="MrClean's External Model Generator", layout="wide")
    _inject_theme()

    app_user = os.getenv("VERVE_APP_USER", "").strip()
    app_password = os.getenv("VERVE_APP_PASSWORD", "").strip()

    if not app_user or not app_password:
        st.error("Authentication is not configured. Set VERVE_APP_USER and VERVE_APP_PASSWORD in deployment secrets.")
        st.stop()

    if "authenticated" not in st.session_state:
        st.session_state.authenticated = False

    if st.session_state.authenticated:
        return

    _render_brand_header("Sign in to access the workflow.")

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

    left, right = st.columns([8, 2], vertical_alignment="top")
    with left:
        _render_brand_header("Upload one raw .xlsm/.xlsx file and run the workflow.")
    with right:
        st.write("")
        if st.button("Log Out", type="secondary"):
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

