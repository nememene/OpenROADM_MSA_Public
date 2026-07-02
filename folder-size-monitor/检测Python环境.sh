#!/usr/bin/env bash
# Python environment check for Linux / macOS

set -e
cd "$(dirname "$0")"

echo "========================================"
echo "  文件夹大小监控 - Python 环境检测"
echo "========================================"
echo

PYTHON_CMD=""
for cmd in python3 python; do
    if command -v "$cmd" >/dev/null 2>&1; then
        PYTHON_CMD="$cmd"
        break
    fi
done

if [ -z "$PYTHON_CMD" ]; then
    echo "[×] 未检测到 Python"
    echo
    if command -v apt-get >/dev/null 2>&1; then
        echo "正在尝试安装 Python..."
        sudo apt-get update -qq
        sudo apt-get install -y python3 python3-pip python3-tk
    elif command -v brew >/dev/null 2>&1; then
        echo "正在尝试安装 Python..."
        brew install python python-tk@3.12 2>/dev/null || brew install python-tk
    else
        echo "请手动安装 Python 3: https://www.python.org/downloads/"
        exit 1
    fi
    PYTHON_CMD=python3
fi

echo "[√] 已检测到 Python："
$PYTHON_CMD --version
echo

echo "正在检测 tkinter..."
if $PYTHON_CMD -c "import tkinter" 2>/dev/null; then
    echo "[√] tkinter 可用"
else
    echo "[×] tkinter 不可用"
    if command -v apt-get >/dev/null 2>&1; then
        echo "正在安装 python3-tk..."
        sudo apt-get install -y python3-tk
    elif command -v brew >/dev/null 2>&1; then
        brew install python-tk@3.12 2>/dev/null || brew install python-tk
    else
        echo "请手动安装 tkinter 组件"
    fi
fi

echo
echo "正在检测 openpyxl..."
if $PYTHON_CMD -c "import openpyxl" 2>/dev/null; then
    echo "[√] openpyxl 已安装"
else
    echo "[!] openpyxl 未安装，正在自动安装..."
    $PYTHON_CMD -m pip install -r requirements.txt
    echo "[√] openpyxl 安装成功"
fi

echo
echo "========================================"
echo "环境检测完成，可以运行 ./启动文件夹大小监控.sh"
echo "========================================"
