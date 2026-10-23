#!/bin/bash
# ═══════════════════════════════════════════════════════════════
#   智能报销系统 - Mac 打包脚本（生成 .app）
#
#   使用方法：
#     1. 安装 PyInstaller：  pip3 install pyinstaller openpyxl
#     2. 在项目根目录运行：  bash 打包脚本_build_mac.sh
#     3. 生成的 .app 在 dist/ 目录下
# ═══════════════════════════════════════════════════════════════
set -e
cd "$(dirname "$0")"

echo "═══ 智能报销系统打包（Mac .app）═══"

# 检查依赖
python3 -c "import PyInstaller" 2>/dev/null || {
    echo "安装 PyInstaller..."
    python3 -m pip install pyinstaller -q
}
python3 -c "import openpyxl" 2>/dev/null || {
    echo "安装 openpyxl..."
    python3 -m pip install openpyxl -q
}

APP_NAME="智能报销系统"

# 清理旧产物
rm -rf build dist_app *.spec 2>/dev/null || true

echo "📦 开始打包..."
python3 -m PyInstaller \
    --noconfirm \
    --onefile \
    --windowed \
    --name "$APP_NAME" \
    --add-data "index.html:." \
    --add-data "template.xlsx:." \
    --hidden-import openpyxl \
    --hidden-import openpyxl.drawing.image \
    --hidden-import openpyxl.utils \
    serve.py

echo ""
echo "✅ 打包完成！"
echo "   产物： dist/$APP_NAME.app"
echo ""
echo "   说明："
echo "   - 双击 .app 即可启动，不需要装 Python"
echo "   - 数据存在 ~/.reimbursement/ 目录"
echo "   - 首次启动会自动建库"
echo "   - 把 .app 拷给同事即可使用（仅限 Mac）"
