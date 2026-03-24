param(
    [int]$Port = 5050
)

$ErrorActionPreference = "Stop"
Set-Location -Path $PSScriptRoot

function Test-PythonLauncher {
    param([string]$Launcher)
    try {
        $out = & $Launcher --version 2>&1
        $txt = ($out | Out-String).Trim()
        if ($LASTEXITCODE -eq 0 -and $txt -match "^Python\s+\d+\.\d+") {
            return $true
        }
    } catch { }
    return $false
}

function Resolve-Python {
    $candidates = @()
    if (Get-Command py -ErrorAction SilentlyContinue) {
        $candidates += "py"
    }
    if (Get-Command python -ErrorAction SilentlyContinue) {
        $candidates += "python"
    }
    $pyLauncherPath = Join-Path $env:LOCALAPPDATA "Programs\Python\Launcher\py.exe"
    if (Test-Path $pyLauncherPath) {
        $candidates += $pyLauncherPath
    }

    foreach ($cand in $candidates | Select-Object -Unique) {
        if (Test-PythonLauncher -Launcher $cand) {
            return $cand
        }
    }

    Write-Host "Python is missing or not runnable. Attempting automatic install with winget..." -ForegroundColor Yellow

    if (-not (Get-Command winget -ErrorAction SilentlyContinue)) {
        Write-Host "ERROR PY001: winget is not available on this machine. Ask Dev Analysts for help." -ForegroundColor Red
        exit 1
    }

    try {
        & winget install --id Python.Python.3.11 -e --source winget --accept-source-agreements --accept-package-agreements --silent
        if ($LASTEXITCODE -ne 0) {
            throw "winget exited with code $LASTEXITCODE"
        }
    } catch {
        Write-Host "ERROR PY002: Python auto-install failed. Ask Dev Analysts for help." -ForegroundColor Red
        Write-Host $_.Exception.Message -ForegroundColor Red
        exit 1
    }

    $postCandidates = @()
    if (Get-Command py -ErrorAction SilentlyContinue) {
        $postCandidates += "py"
    }
    if (Get-Command python -ErrorAction SilentlyContinue) {
        $postCandidates += "python"
    }
    if (Test-Path $pyLauncherPath) {
        $postCandidates += $pyLauncherPath
    }
    foreach ($cand in $postCandidates | Select-Object -Unique) {
        if (Test-PythonLauncher -Launcher $cand) {
            return $cand
        }
    }

    Write-Host "ERROR PY003: Python install completed but no runnable launcher is available in this session. Ask Dev Analysts for help." -ForegroundColor Red
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
        if ($LASTEXITCODE -ne 0 -or -not (Test-Path $VenvPython)) {
            Write-Host "ERROR PY004: Failed to create virtual environment python executable. Ask Dev Analysts for help." -ForegroundColor Red
            exit 1
        }
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

# Avoid Start-Job (blocked in some corporate environments). Open browser after short delay.
$browserCmd = "Start-Sleep -Seconds 3; Start-Process '$url'"
Start-Process -FilePath "powershell.exe" -WindowStyle Hidden -ArgumentList @(
    "-NoProfile",
    "-ExecutionPolicy",
    "Bypass",
    "-Command",
    $browserCmd
) | Out-Null

try {
    & $venvPython -m app.web
} catch {
    Write-Host "" -ForegroundColor Red
    Write-Host "Server startup failed." -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
    Read-Host "Press Enter to close"
    exit 1
}
