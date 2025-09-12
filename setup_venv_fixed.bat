@echo off
setlocal ENABLEDELAYEDEXPANSION
REM === Portable venv setup for Unified Launcher ===

REM Resolve this script's folder (handles spaces & parentheses)
set "ROOT=%~dp0"
pushd "%ROOT%"

REM Ensure tools.json is here
if not exist "tools.json" (
  echo [!] tools.json not found in "%ROOT%"
  echo     Place this .bat next to tools.json and run again.
  pause
  popd
  exit /b 1
)

REM Choose Python launcher
where py >nul 2>&1
if %ERRORLEVEL%==0 (set "PY=py") else (set "PY=python")

echo [i] Creating venv (if needed)...
"%PY%" -m venv "venv"
if not exist "venv\Scripts\python.exe" (
  echo [!] venv creation failed. Ensure Python is installed and on PATH.
  pause
  popd
  exit /b 1
)

set "VENV_PY=%ROOT%venv\Scripts\python.exe"
set "VENV_PIP=%ROOT%venv\Scripts\pip.exe"

echo [i] Upgrading pip...
"%VENV_PY%" -m pip install --upgrade pip

if exist "requirements.txt" (
  echo [i] Installing packages from requirements.txt ...
  "%VENV_PIP%" install -r "requirements.txt"
) else (
  echo [i] No requirements.txt found. Skipping package installs.
)

echo [i] Patching tools.json to use portable interpreter placeholders...
"%VENV_PY%" "setup_venv.py" "%ROOT%tools.json"
if errorlevel 1 (
  echo [!] Failed to update tools.json
  pause
  popd
  exit /b 1
)

echo.
echo [✓] Portable venv is ready at "%ROOT%venv"
echo     You can now zip this entire folder and share it.
echo.
pause
popd
exit /b 0
