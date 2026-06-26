#!/usr/bin/env bash
# Mac / Linux 打包脚本（需在目标系统上执行，无法跨平台交叉打包）
set -euo pipefail
cd "$(dirname "$0")"

PYTHON="${PYTHON:-python3}"

echo "使用 Python: $($PYTHON --version)"
$PYTHON -m pip install -r requirements.txt
$PYTHON -m PyInstaller build.spec --noconfirm

echo ""
echo "打包完成，输出目录: dist/销售订单合同生成工具/"
if [[ "$(uname -s)" == "Darwin" ]]; then
    echo "Mac 可执行文件: dist/销售订单合同生成工具/销售订单合同生成工具"
    echo "首次运行若被拦截，请在「系统设置 → 隐私与安全性」中允许，或执行:"
    echo "  xattr -cr \"dist/销售订单合同生成工具/销售订单合同生成工具\""
fi
