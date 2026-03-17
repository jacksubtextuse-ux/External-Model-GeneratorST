from __future__ import annotations

import os
import platform
from pathlib import Path
from typing import Protocol


class RunnerProtocol(Protocol):
    def run(self, output_dir: Path | None = None) -> dict: ...


def get_runner(input_file: Path, option: str = "base"):
    prefer_com = platform.system() == "Windows" and os.getenv("VERVE_FORCE_OPENPYXL", "").strip() != "1"
    if prefer_com:
        try:
            from app.engine_com import VerveWorkflowRunnerCom

            return VerveWorkflowRunnerCom(input_file, option=option)
        except Exception:
            pass

    from app.engine import VerveWorkflowRunner

    return VerveWorkflowRunner(input_file, option=option)
