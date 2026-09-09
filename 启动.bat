@echo off
REM ============================================================
REM  价格牌处理工具 - 一键启动脚本（双击即可）
REM  自动定位带 tkinter 的 Python，装齐依赖，再启动 GUI
REM ============================================================

setlocal enabledelayedexpansion
chcp 65001 >nul

cd /d "%~dp0"

echo.
echo ==========================================================
echo   价格牌处理工具 - 一键启动
echo ==========================================================
echo.

REM ---------- 1. 选 Python ----------
REM 优先级：py launcher > 系统 Python 3.10+ > 系统 Python 3.9
set "PYEXE="

REM 1.1 尝试 py launcher（推荐）
for /f "delims=" %%i in ('py -3 -c "import sys; print(sys.executable)" 2^>nul') do set "PYEXE=%%i"
if defined PYEXE goto :py_found

REM 1.2 尝试系统 python（用户机器上的）
where python >nul 2>&1
if %errorlevel%==0 (
    for /f "delims=" %%i in ('python -c "import sys; print(sys.executable)" 2^>nul') do set "PYEXE=%%i"
)

if not defined PYEXE (
    echo [错误] 找不到可用的 Python。
    echo 请先安装 Python 3.10 或更高版本（勾选 Add to PATH）：
    echo   https://www.python.org/downloads/
    pause
    exit /b 1
)

:py_found
echo [信息] 使用 Python: %PYEXE%
%PYEXE% --version
echo.

REM ---------- 2. 检查并安装依赖 ----------
echo [信息] 检查依赖（PyMuPDF / openpyxl / Pillow）...
%PYEXE% -c "import fitz, openpyxl, PIL, tkinter" 2>nul
if %errorlevel%==0 (
    echo [信息] 依赖已就绪，跳过安装。
    goto :run
)

echo [信息] 正在安装依赖（首次约 1-3 分钟）...
%PYEXE% -m pip install --upgrade pip >nul 2>&1
%PYEXE% -m pip install -i https://pypi.tuna.tsinghua.edu.cn/simple --timeout 120 -r requirements.txt
if %errorlevel% neq 0 (
    echo.
    echo [错误] 依赖安装失败。
    echo 请检查网络后重试，或手动运行：
    echo   %PYEXE% -m pip install -r requirements.txt
    pause
    exit /b 1
)
echo [信息] 依赖安装完成。
echo.

:run
REM ---------- 3. 启动 GUI ----------
echo [信息] 启动 GUI...
echo.
%PYEXE% run.py
if %errorlevel% neq 0 (
    echo.
    echo [错误] GUI 启动失败，错误码 %errorlevel%。
    pause
)
endlocal
