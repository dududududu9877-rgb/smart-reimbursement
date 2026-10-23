#!/bin/bash
# 智能报销录入系统 — 启动器
# 双击此文件即可启动，首次启动会自动安装依赖

cd "$(dirname "$0")"

echo "═══════════════════════════════════"
echo "  智能报销录入系统 v2.0"
echo "═══════════════════════════════════"

# 检查 Python
if ! command -v python3 &>/dev/null; then
    echo "❌ 需要 Python 3，请先安装: https://www.python.org/downloads/"
    echo "   然后双击此文件重试。"
    read -p "按回车退出..."
    exit 1
fi

# 检查依赖
python3 -c "import openpyxl" 2>/dev/null || {
    echo "正在安装 openpyxl..."
    python3 -m pip install openpyxl Pillow -q
}

# 启动服务
echo "启动中..."
python3 serve.py

read -p "按回车退出..."
