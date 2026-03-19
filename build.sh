#!/bin/bash
# 养老金融通报数据自动化处理系统
# 打包脚本 (macOS/Linux)
#
# 使用方法: ./build.sh

set -e

echo "========================================"
echo "养老金融通报数据处理系统 - 打包脚本"
echo "========================================"
echo ""

# 获取脚本所在目录
SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
cd "$SCRIPT_DIR"

# 检查Python
if ! command -v python3 &> /dev/null; then
    echo "错误: 未找到Python3"
    exit 1
fi

# 检查依赖
echo "检查依赖..."
python3 -c "import openpyxl, pandas" 2>/dev/null || {
    echo "安装依赖..."
    pip3 install -r requirements.txt
}

# 检查PyInstaller
python3 -c "import PyInstaller" 2>/dev/null || {
    echo "安装PyInstaller..."
    pip3 install pyinstaller
}

# 创建输出目录
mkdir -p dist

# 打包
echo ""
echo "开始打包..."
echo ""

pyinstaller pension_financial.spec --clean

echo ""
echo "========================================"
echo "打包完成！"
echo "输出目录: dist/养老金融通报数据处理系统"
echo "========================================"

# macOS特定提示
if [[ "$OSTYPE" == "darwin"* ]]; then
    echo ""
    echo "macOS用户提示:"
    echo "如果遇到'无法打开因为无法验证开发者'错误,"
    echo "请在系统偏好设置 > 安全性与隐私中允许运行此应用"
fi
