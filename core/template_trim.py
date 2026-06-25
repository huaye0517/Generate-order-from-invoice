# -*- coding: utf-8 -*-
"""模板页自适应裁剪：按分析结果删除留白行，保留表格线条"""
from typing import Dict

import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

from .template_layout import TemplateLayout

BUYER_INFO_KEYS = (
    "company",
    "address",
    "signatory",
    "phone",
    "bank",
    "account",
    "credit_code",
)


def _row_is_empty(ws, row: int) -> bool:
    for col in range(1, ws.max_column + 1):
        if ws.cell(row, col).value not in (None, ""):
            return False
    return True


def _replace_row_in_formula(formula: str, old_row: int, new_row: int) -> str:
    if not formula or not isinstance(formula, str):
        return formula
    text = formula.replace(f"I{old_row}", f"I{new_row}")
    text = text.replace(f"G{old_row}", f"G{new_row}")
    return text


def _unmerge_covering(ws: Worksheet, row: int, col: int) -> None:
    """解除覆盖指定单元格的合并，以便写入值"""
    to_remove = []
    for merged in ws.merged_cells.ranges:
        if (
            merged.min_row <= row <= merged.max_row
            and merged.min_col <= col <= merged.max_col
        ):
            to_remove.append(str(merged))
    for ref in to_remove:
        ws.unmerge_cells(ref)


def _write_buyer_value(ws: Worksheet, row: int, col: int, value: str) -> None:
    """写入需方字段值（处理删行后合并单元格错位）"""
    _unmerge_covering(ws, row, col)
    ws.cell(row, col).value = value


def trim_template_sheet(
    dest_path: str,
    layout: TemplateLayout,
    buyer_info: Dict[str, str],
    product_line_count: int,
) -> None:
    """
    根据自动分析的 layout 裁剪模板：
    1. 删除产品区多余空行
    2. 删除需方区块前的空白行
    3. 更新合计/总合计公式
    4. 写入需方静态信息（打开即显示，无需点击重算）
    """
    wb = openpyxl.load_workbook(dest_path)
    ws = wb[layout.template_sheet]

    last_data_row = layout.last_product_row(product_line_count)
    delete_count = layout.rows_to_delete_count(product_line_count)

    if delete_count > 0:
        ws.delete_rows(last_data_row + 1, delete_count)

    total_row = layout.total_row - delete_count
    qty_col = layout.qty_sum_col
    amt_col = layout.amt_sum_col
    qty_letter = get_column_letter(qty_col)
    amt_letter = get_column_letter(amt_col)

    ws.cell(total_row, qty_col).value = (
        f"=SUM({qty_letter}{layout.product_first_row}:{qty_letter}{last_data_row})"
    )
    ws.cell(total_row, amt_col).value = (
        f"=SUM({amt_letter}{layout.product_first_row}:{amt_letter}{last_data_row})"
    )

    grand_row = layout.grand_total_row - delete_count
    ws.cell(grand_row, amt_col).value = f"={amt_letter}{total_row}"

    e_grand = ws.cell(grand_row, 5).value
    if e_grand and isinstance(e_grand, str):
        ws.cell(grand_row, 5).value = _replace_row_in_formula(
            e_grand, layout.total_row, total_row
        )

    gap_deleted = 0
    for gap_row in sorted(layout.empty_rows_before_buyer, reverse=True):
        shifted = gap_row - delete_count
        if shifted > 0 and shifted <= ws.max_row and _row_is_empty(ws, shifted):
            ws.delete_rows(shifted, 1)
            gap_deleted += 1

    row_shift = delete_count + gap_deleted
    buyer_rows = {
        key: row - row_shift
        for key, row in layout.buyer_fields.items()
    }

    party_cell = layout.party_cell
    party_row = ws[party_cell].row
    party_col = ws[party_cell].column
    _unmerge_covering(ws, party_row, party_col)
    ws.cell(party_row, party_col).value = buyer_info.get(
        "party", buyer_info.get("company", "")
    )

    value_col = layout.buyer_value_col
    for key in BUYER_INFO_KEYS:
        row = buyer_rows.get(key)
        if row and row <= ws.max_row:
            _write_buyer_value(ws, row, value_col, buyer_info.get(key, ""))

    # 重置工作表视图，打开时从第1行开始显示
    ws.sheet_view.topLeftCell = "A1"
    if ws.sheet_view.selection:
        ws.sheet_view.selection[0].activeCell = "A1"
        ws.sheet_view.selection[0].sqref = "A1"

    wb.save(dest_path)
    wb.close()
