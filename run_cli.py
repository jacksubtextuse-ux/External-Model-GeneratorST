from __future__ import annotations

import argparse
import json
from pathlib import Path

from app.engine_factory import get_runner
from app.validator import validate_workbook

BASE = Path(__file__).resolve().parent
RUNS_ROOT = BASE / "runs"
OPTION_OUTPUT_DIRS = {
    "base": RUNS_ROOT / "base-tests",
    "front-range": RUNS_ROOT / "front-range-tests",
    "lp": RUNS_ROOT / "lp-tests",
    "lender": RUNS_ROOT / "lender-tests",
}
OPTION_ASSERTIONS = {
    "base": BASE / "VERVE-Proforma-Cleaner-v1.1-Assertions.json",
    "front-range": BASE / "VERVE-Proforma-Cleaner-v1.1-Assertions.json",
    "lp": BASE / "VERVE-Proforma-Cleaner-v1.1-LP-Assertions.json",
    "lender": BASE / "VERVE-Proforma-Cleaner-v1.1-Lender-Assertions.json",
}


def main() -> None:
    parser = argparse.ArgumentParser(description="Run VERVE Proforma Cleaner workflow")
    parser.add_argument("input_file", type=Path, help="Path to source .xlsm/.xlsx")
    parser.add_argument("--output-dir", type=Path, default=None, help="Output directory")
    parser.add_argument(
        "--option",
        choices=["base", "front-range", "lp", "lender"],
        default="base",
        help="Workflow option to run",
    )
    parser.add_argument(
        "--assertions",
        type=Path,
        default=None,
        help="Assertions JSON for validation",
    )
    args = parser.parse_args()

    output_dir = args.output_dir
    if output_dir is None:
        output_dir = OPTION_OUTPUT_DIRS.get(args.option, OPTION_OUTPUT_DIRS["base"])
    output_dir.mkdir(parents=True, exist_ok=True)

    assertions_file = args.assertions or OPTION_ASSERTIONS.get(args.option, OPTION_ASSERTIONS["base"])

    runner = get_runner(args.input_file, option=args.option)
    result = runner.run(output_dir)
    output_file = Path(result["output_file"])
    validation = validate_workbook(assertions_file, output_file)

    print(json.dumps({"run": result, "validation": validation}, indent=2))


if __name__ == "__main__":
    main()
