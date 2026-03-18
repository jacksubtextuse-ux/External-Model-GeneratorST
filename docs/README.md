# VERVE Proforma Cleaner (v1.2 COM-first)

## Engine behavior
- Default engine: **Excel COM** (`pywin32`) to preserve workbook internals and avoid Excel repair logs from Open XML part loss.
- Fallback engine: `openpyxl` only if COM is unavailable.

## What this includes
- `app/engine_com.py`: COM-based workflow runner (default).
- `app/engine.py`: openpyxl fallback runner.
- `app/engine_factory.py`: runtime engine selector.
- `app/validator.py`: strict output validator using assertions.
- `app/report.py`: side-by-side report generator.
- `app/web.py`: upload UI for one-file-at-a-time testing.
- `run_cli.py`: CLI entrypoint.

## Requirements
- Windows + installed Microsoft Excel desktop app
- Python with:
```powershell
py -m pip install -r requirements.txt
```

## Run web UI
```powershell
py -m app.web
```
Open `http://127.0.0.1:5050`.

## Run CLI
```powershell
py run_cli.py "C:\path\to\input.xlsm" --output-dir "C:\path\to\out"
```

## Notes
- Keep the source workbook closed in Excel before running.
- Step 40 remains strict: adjacent `Residential Parking Spaces` value must be hardcoded (non-formula).
