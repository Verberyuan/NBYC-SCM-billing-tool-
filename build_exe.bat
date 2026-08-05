@echo off
chcp 65001 >nul
cd /d "%~dp0"

echo ========================================
echo 当前目录：
cd
echo ========================================
echo.

echo 正在检查 Python...
python --version

if errorlevel 1 (
    echo.
    echo 未找到 Python。
    echo 请确认 Python 已安装并加入 PATH。
    pause
    exit /b 1
)

echo.
echo 正在检查 PyInstaller...
python -m PyInstaller --version

if errorlevel 1 (
    echo.
    echo 未找到 PyInstaller，正在安装...
    python -m pip install pyinstaller

    if errorlevel 1 (
        echo.
        echo PyInstaller 安装失败。
        pause
        exit /b 1
    )
)

echo.
echo 正在检查必要文件...

if not exist "BillingTool.spec" (
    echo 找不到 BillingTool.spec
    pause
    exit /b 1
)

if not exist "billing_tool.py" (
    echo 找不到 billing_tool.py
    pause
    exit /b 1
)

if not exist "billing_tool.ico" (
    echo 找不到 billing_tool.ico
    pause
    exit /b 1
)

echo.
echo 正在清理旧文件...

if exist "build" rmdir /s /q "build"
if exist "dist" rmdir /s /q "dist"

echo.
echo 正在打包，请稍候...
python -m PyInstaller --noconfirm --clean "BillingTool.spec"

if errorlevel 1 (
    echo.
    echo ========================================
    echo 打包失败，请查看上方错误信息。
    echo ========================================
    pause
    exit /b 1
)

echo.
echo ========================================
echo 打包完成。
echo 请查看 dist 文件夹。
echo ========================================
pause
