# -*- coding: utf-8 -*-
"""
分析销售订单合同模板布局（命令行工具）

用法:
  python scripts/analyze_template.py sample/销售订单合同.xlsx
  python scripts/analyze_template.py 你的新模板.xlsx -s 模板

更换新模板前可先运行本脚本，确认工具能正确识别产品区、合计行、需方信息区。
"""
import argparse
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from core.template_analyzer import analyze_template_layout


def main() -> int:
    parser = argparse.ArgumentParser(description="分析销售订单合同模板布局，用于自适应裁剪")
    parser.add_argument("template", help="销售订单合同模板 xlsx 路径")
    parser.add_argument("-s", "--sheet", help="指定工作表名称（默认自动识别）")
    args = parser.parse_args()

    template_path = Path(args.template)
    if not template_path.is_file():
        print(f"错误：文件不存在 {template_path}")
        return 1

    try:
        layout = analyze_template_layout(str(template_path), sheet_name=args.sheet)
    except ValueError as exc:
        print(f"分析失败：{exc}")
        return 1

    print("=" * 50)
    print("模板布局分析结果")
    print("=" * 50)
    print(layout.summary())
    print("=" * 50)
    print("若以上信息与您的模板一致，生成时将自动按此布局裁剪留白。")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
