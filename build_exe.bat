@echo off
setlocal
cd /d "%~dp0"

rem Keep this BAT file ASCII-only to avoid mojibake on Chinese Windows.
chcp 65001 >nul
set "PYTHONUTF8=1"
set "PYTHONIOENCODING=utf-8"
set "PIP_DISABLE_PIP_VERSION_CHECK=1"

echo ========================================
echo BillingTool build script
echo Working directory:
cd
echo ========================================
echo.

where py >nul 2>nul
if not errorlevel 1 (
    set "PY=py -3"
) else (
    set "PY=python"
)

echo Checking Python...
%PY% --version
if errorlevel 1 (
    echo.
    echo Python was not found.
    echo Install 64-bit Python 3.11 or 3.12 and enable Add Python to PATH.
    pause
    exit /b 1
)

echo.
if not exist ".venv\Scripts\python.exe" (
    echo Creating virtual environment .venv ...
    %PY% -m venv .venv
    if errorlevel 1 (
        echo Failed to create virtual environment.
        pause
        exit /b 1
    )
)

call ".venv\Scripts\activate.bat"
if errorlevel 1 (
    echo Failed to activate virtual environment.
    pause
    exit /b 1
)

echo.
echo Installing or updating dependencies...
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
if errorlevel 1 (
    echo.
    echo Dependency installation failed.
    echo Check your network, proxy, and access to pypi.org.
    pause
    exit /b 1
)

echo.
echo Checking required source files...
for %%F in (BillingTool_portable.spec billing_tool.py billing_core.py customer_config.py fuel_rates.py sheet_merge.py money.py rate_store.py test_regression.py requirements.txt) do (
    if not exist "%%F" (
        echo Missing file: %%F
        pause
        exit /b 1
    )
)

echo.
echo Running regression tests...
python test_regression.py
if errorlevel 1 (
    echo.
    echo ========================================
    echo Regression tests failed. Packaging stopped.
    echo Review the output above.
    echo ========================================
    pause
    exit /b 1
)

echo.
echo Cleaning old build folders...
if exist "build" rmdir /s /q "build"
if exist "dist" rmdir /s /q "dist"

echo.
echo Packaging BillingTool.exe ...
python -m PyInstaller --noconfirm --clean "BillingTool_portable.spec"
if errorlevel 1 (
    echo.
    echo ========================================
    echo Packaging failed. Review the output above.
    echo ========================================
    pause
    exit /b 1
)

echo.
echo ========================================
echo Build completed successfully.
echo Output: dist\BillingTool.exe
echo ========================================
pause
endlocal
