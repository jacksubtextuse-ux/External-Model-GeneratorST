@echo off
setlocal
cd /d "%~dp0"
powershell -ExecutionPolicy Bypass -File "%~dp0Start_MrClean.ps1"
if errorlevel 1 (
  echo.
  echo Launcher exited with an error. Ask Dev Analysts for help.
  pause
)
endlocal
