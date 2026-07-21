@echo off
echo ============================================
echo   Billing Tool - Build EXE
echo ============================================
echo.

where python >nul 2>nul
if errorlevel 1 (
    echo [ERROR] Python not found. Please install Python 3.9+ first.
    echo Make sure to check "Add python.exe to PATH" during installation.
    echo Download: https://www.python.org/downloads/
    echo.
    pause
    exit /b 1
)

echo Installing/updating required packages, this may take a few minutes...
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
if errorlevel 1 (
    echo [ERROR] Failed to install dependencies. Check your network connection and retry.
    pause
    exit /b 1
)

echo.
echo Building exe, please wait (usually 1-2 minutes)...
python -m PyInstaller --onefile --noconsole --name "BillingTool" --clean billing_tool.py
if errorlevel 1 (
    echo [ERROR] Build failed. Please check the log above.
    pause
    exit /b 1
)

echo.
if exist "dist\BillingTool.exe" (
    copy /Y "dist\BillingTool.exe" "BillingTool.exe" >nul
    echo ============================================
    echo Build succeeded!
    echo exe file created at: %cd%\BillingTool.exe
    echo You can double-click it directly, or copy it anywhere you like.
    echo ============================================
) else (
    echo Could not find the generated exe file. Please check the log above.
)

echo.
pause
