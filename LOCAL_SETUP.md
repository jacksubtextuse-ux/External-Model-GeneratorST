# MrClean Local Setup

This package runs fully on your Windows machine using Excel COM.

## Quick Start
1. Extract this folder to a local path.
2. Double-click `Start_MrClean.bat`.
3. Wait for setup to complete (first run installs dependencies).
4. Your browser opens at `http://127.0.0.1:5050`.

## What Start_MrClean Does Automatically
1. Detects Python (`py` or `python`).
2. Creates a local virtual environment (`.venv`) if missing.
3. Installs required packages from `config/requirements.txt`.
4. Installs core packages explicitly (`pywin32`, `openpyxl`, `Flask`).
5. Verifies module imports.
6. Verifies Excel COM is available.
7. Starts the local app and opens the website.

## Hard Requirements
1. Windows OS.
2. Microsoft Excel desktop installed and activated for this user.
3. Python 3.10+ installed and available as `py` or `python`.

## Stop the App
1. Double-click `Stop_MrClean.bat`.

## Troubleshooting
1. If startup says Python launcher not found: install Python and retry.
2. If Excel COM check fails: open Excel once manually, sign in/activate Office, then retry.
3. If port 5050 is already in use: run `Stop_MrClean.bat`, then start again.
4. If corporate security blocks installs: run PowerShell as user and execute `Start_MrClean.ps1` to view detailed errors.

