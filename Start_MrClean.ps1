param(
    [int]$Port = 5050
)

$ErrorActionPreference = "Stop"
Set-Location -Path $PSScriptRoot

function Resolve-Python {
    if (Get-Command py -ErrorAction SilentlyContinue) {
        return "py"
    }
    if (Get-Command python -ErrorAction SilentlyContinue) {
        return "python"
    }

    Write-Host "Python was not found. Attempting automatic install with winget..." -ForegroundColor Yellow

    if (-not (Get-Command winget -ErrorAction SilentlyContinue)) {
        Write-Host "ERROR PY001: winget is not available on this machine. Ask Dev Analysts for help." -ForegroundColor Red
        exit 1
    }

    try {
        & winget install --id Python.Python.3.11 -e --source winget --accept-source-agreements --accept-package-agreements --silent
    } catch {
        Write-Host "ERROR PY002: Python auto-install failed. Ask Dev Analysts for help." -ForegroundColor Red
        Write-Host $_.Exception.Message -ForegroundColor Red
        exit 1
    }

    if (Get-Command py -ErrorAction SilentlyContinue) {
        return "py"
    }
    if (Get-Command python -ErrorAction SilentlyContinue) {
        return "python"
    }

    $pyLauncher = Join-Path $env:LOCALAPPDATA "Programs\Python\Launcher\py.exe"
    if (Test-Path $pyLauncher) {
        return $pyLauncher
    }

    Write-Host "ERROR PY003: Python install completed but launcher is still unavailable in this session. Ask Dev Analysts for help." -ForegroundColor Red
    exit 1
}

function Ensure-Venv {
    param(
        [string]$PythonLauncher,
        [string]$VenvPython
    )
    if (-not (Test-Path $VenvPython)) {
        Write-Host "Creating virtual environment..." -ForegroundColor Yellow
        & $PythonLauncher -m venv .venv
    }
}

function Ensure-PipPackages {
    param([string]$VenvPython)

    Write-Host "Installing/updating Python dependencies..." -ForegroundColor Yellow
    & $VenvPython -m pip install --upgrade pip
    & $VenvPython -m pip install -r config\\requirements.txt

    # Force key packages as a safety net for first-time installs.
    & $VenvPython -m pip install pywin32 openpyxl Flask

    # Some environments need this after pywin32 install; safe to ignore if absent.
    try {
        & $VenvPython "$PSScriptRoot\.venv\Scripts\pywin32_postinstall.py" -install | Out-Null
    } catch { }
}

function Test-ModuleImports {
    param([string]$VenvPython)

    Write-Host "Verifying required Python modules..." -ForegroundColor Yellow
    $code = @"
import importlib
required = ["flask", "openpyxl", "win32com.client"]
missing = []
for m in required:
    try:
        importlib.import_module(m)
    except Exception:
        missing.append(m)
if missing:
    raise SystemExit("Missing modules: " + ", ".join(missing))
print("Python module check passed")
"@
    & $VenvPython -c $code
}

function Test-ExcelCom {
    param([string]$VenvPython)

    Write-Host "Checking Excel COM availability..." -ForegroundColor Yellow
    $code = @"
import win32com.client
app = win32com.client.DispatchEx("Excel.Application")
app.Visible = False
app.DisplayAlerts = False
app.Quit()
print("Excel COM check passed")
"@
    & $VenvPython -c $code
}

Write-Host "== MrClean Local Launcher ==" -ForegroundColor Cyan
$pythonLauncher = Resolve-Python
$venvPath = Join-Path $PSScriptRoot ".venv"
$venvPython = Join-Path $venvPath "Scripts\python.exe"

Ensure-Venv -PythonLauncher $pythonLauncher -VenvPython $venvPython
Ensure-PipPackages -VenvPython $venvPython
Test-ModuleImports -VenvPython $venvPython

try {
    Test-ExcelCom -VenvPython $venvPython
} catch {
    Write-Host "" -ForegroundColor Red
    Write-Host "Excel COM check failed." -ForegroundColor Red
    Write-Host "Install/activate Microsoft Excel desktop for this Windows user, then rerun." -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
    exit 1
}

$url = "http://127.0.0.1:$Port"
Write-Host "Starting local server (Ctrl+C to stop)..." -ForegroundColor Green

$browserJob = Start-Job -ScriptBlock {
    param([string]$TargetUrl)
    $deadline = (Get-Date).AddSeconds(90)
    while ((Get-Date) -lt $deadline) {
        try {
            Invoke-WebRequest -Uri $TargetUrl -UseBasicParsing -TimeoutSec 2 | Out-Null
            Start-Process $TargetUrl | Out-Null
            return
        } catch {
            Start-Sleep -Milliseconds 700
        }
    }
} -ArgumentList $url

try {
    & $venvPython -m app.web
} catch {
    Write-Host "" -ForegroundColor Red
    Write-Host "Server startup failed." -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
    Read-Host "Press Enter to close"
    exit 1
} finally {
    if ($browserJob) {
        try {
            Stop-Job -Id $browserJob.Id -ErrorAction SilentlyContinue | Out-Null
            Remove-Job -Id $browserJob.Id -Force -ErrorAction SilentlyContinue | Out-Null
        } catch { }
    }
}
