@echo off
setlocal
cd /d "%~dp0"

rem Run the GUI from source without packaging.
chcp 65001 >nul
set "PYTHONUTF8=1"
set "PYTHONIOENCODING=utf-8"
set "PIP_DISABLE_PIP_VERSION_CHECK=1"

where py >nul 2>nul
if not errorlevel 1 (
    set "PY=py -3"
) else (
    set "PY=python"
)

%PY% --version
if errorlevel 1 (
    echo Python was not found.
    echo Install 64-bit Python 3.11 or 3.12 and enable Add Python to PATH.
    pause
    exit /b 1
)

if not exist ".venv\Scripts\python.exe" (
    echo Creating virtual environment and installing dependencies...
    %PY% -m venv .venv
    call ".venv\Scripts\activate.bat"
    python -m pip install --upgrade pip
    python -m pip install -r requirements.txt
    if errorlevel 1 (
        echo Dependency installation failed.
        pause
        exit /b 1
    )
) else (
    call ".venv\Scripts\activate.bat"
)

python billing_tool.py
if errorlevel 1 (
    echo.
    echo BillingTool exited with an error.
    pause
)
endlocal
