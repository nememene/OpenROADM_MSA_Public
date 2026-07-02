@echo off
chcp 65001 >nul
title Python 环境检测
cd /d "%~dp0"

echo ========================================
echo   文件夹大小监控 - Python 环境检测
echo ========================================
echo.

set "PYTHON_CMD="
set "NEED_INSTALL=0"

where python >nul 2>&1
if %errorlevel%==0 (
    set "PYTHON_CMD=python"
    goto :found
)

where python3 >nul 2>&1
if %errorlevel%==0 (
    set "PYTHON_CMD=python3"
    goto :found
)

where py >nul 2>&1
if %errorlevel%==0 (
    set "PYTHON_CMD=py -3"
    goto :found
)

set "NEED_INSTALL=1"
goto :install

:found
echo [√] 已检测到 Python：
%PYTHON_CMD% --version
echo.

echo 正在检测 tkinter（图形界面需要）...
%PYTHON_CMD% -c "import tkinter" >nul 2>&1
if %errorlevel%==0 (
    echo [√] tkinter 可用
) else (
    echo [×] tkinter 不可用
    echo     请重新安装 Python，安装时勾选 "tcl/tk and IDLE"
    set "NEED_INSTALL=1"
)

echo.
echo 正在检测 openpyxl（Excel 导出需要）...
%PYTHON_CMD% -c "import openpyxl" >nul 2>&1
if %errorlevel%==0 (
    echo [√] openpyxl 已安装
) else (
    echo [!] openpyxl 未安装，正在自动安装...
    %PYTHON_CMD% -m pip install -r "%~dp0requirements.txt"
    if %errorlevel%==0 (
        echo [√] openpyxl 安装成功
    ) else (
        echo [×] 安装失败，请手动执行: pip install -r requirements.txt
    )
)

echo.
echo ========================================
echo 环境检测完成，可以双击「启动文件夹大小监控.bat」使用程序。
echo ========================================
pause
exit /b 0

:install
echo [×] 未检测到 Python 环境
echo.
echo 正在尝试自动安装 Python...

where winget >nul 2>&1
if %errorlevel%==0 (
    echo 使用 winget 安装 Python 3.12 ...
    winget install --id Python.Python.3.12 -e --accept-source-agreements --accept-package-agreements
    if %errorlevel%==0 (
        echo.
        echo [√] Python 安装完成，请关闭本窗口后重新运行此脚本确认环境。
        pause
        exit /b 0
    )
)

echo.
echo 自动安装失败，请手动安装：
echo   1. 打开 https://www.python.org/downloads/
echo   2. 下载并安装 Python 3
echo   3. 安装时务必勾选 "Add Python to PATH"
echo   4. 安装时务必勾选 "tcl/tk and IDLE"
echo   5. 安装完成后重新运行本脚本
echo.
start https://www.python.org/downloads/
pause
exit /b 1
