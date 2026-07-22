@echo off
chcp 65001 >nul
echo ============================================
echo   电商审单核对工具 - 一键打包为 exe
echo ============================================
echo.

REM 进入本批处理所在目录（确保无论从哪双击都能正确执行）
cd /d "%~dp0"

REM 1) 安装 PyInstaller（仅需第一次，已安装会自动跳过）
echo [1/2] 正在检查/安装 PyInstaller ...
pip install pyinstaller
if errorlevel 1 (
    echo.
    echo [错误] pip 安装失败。请确认已安装 Python 且已勾选“Add Python to PATH”。
    echo 下载 Python：https://www.python.org/downloads/  （安装时务必勾选 Add to PATH）
    pause
    exit /b 1
)

REM 2) 执行打包
echo.
echo [2/2] 正在打包，请稍候（首次约 1-3 分钟）...
pyinstaller build.spec
if errorlevel 1 (
    echo.
    echo [错误] 打包失败，请查看上方报错信息。
    pause
    exit /b 1
)

echo.
echo ============================================
echo   打包完成！
echo   生成的程序在：dist\审单核对工具.exe
echo   把「审单核对工具.exe」发给同事，双击即可使用。
echo   （无需安装 Python，无需联网）
echo ============================================
echo.
pause
