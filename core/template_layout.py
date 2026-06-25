# -*- coding: utf-8 -*-
"""模板「模板」工作表产品区行号常量"""
PRODUCT_FIRST_ROW = 7
PRODUCT_LAST_ROW = 206


def empty_product_row_range(product_line_count: int) -> tuple[int, int] | None:
    """
    根据产品行数计算需要隐藏的空行范围（含起止行）。
    产品区为第 7–206 行，共 200 行。
    """
    if product_line_count < 0:
        product_line_count = 0
    max_lines = PRODUCT_LAST_ROW - PRODUCT_FIRST_ROW + 1
    if product_line_count >= max_lines:
        return None
    hide_start = PRODUCT_FIRST_ROW + product_line_count
    if hide_start > PRODUCT_LAST_ROW:
        return None
    return hide_start, PRODUCT_LAST_ROW
