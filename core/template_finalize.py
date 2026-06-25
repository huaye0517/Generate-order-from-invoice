# -*- coding: utf-8 -*-
"""生成流程：分析模板 + 裁剪 + 写入"""
import shutil
from typing import Dict

from .template_analyzer import analyze_template_layout
from .template_layout import TemplateLayout
from .template_trim import trim_template_sheet


def prepare_template_output(
    template_path: str,
    dest_path: str,
    buyer_info: Dict[str, str],
    product_line_count: int,
    layout: TemplateLayout | None = None,
) -> TemplateLayout:
    """
    复制模板并完成自适应裁剪与需方信息写入。
    layout 可传入以复用同批次分析结果，否则自动分析 template_path。
    """
    if layout is None:
        layout = analyze_template_layout(template_path)

    shutil.copy2(template_path, dest_path)
    trim_template_sheet(dest_path, layout, buyer_info, product_line_count)
    return layout
