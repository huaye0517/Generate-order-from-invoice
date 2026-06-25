# -*- coding: utf-8 -*-
"""自动分析销售订单合同模板布局，供裁剪与写入适配使用"""
import re
from typing import Dict, List, Optional, Tuple

import openpyxl
from openpyxl.worksheet.worksheet import Worksheet

from .template_layout import TemplateLayout

# 产品行 B 列典型公式特征
PRODUCT_ROW_PATTERN = re.compile(r'=IF\s*\(\s*ROW\s*\(\s*A\d+\s*\)', re.I)
SUM_PATTERN = re.compile(
    r"=SUM\s*\(\s*([A-Z]+)(\d+)\s*:\s*([A-Z]+)(\d+)\s*\)",
    re.I,
)

# 需方字段标签 -> 语义键
BUYER_LABEL_KEYS: List[Tuple[str, Tuple[str, ...]]] = [
    ("company", ("单位名称",)),
    ("address", ("单位地址",)),
    ("signatory", ("代表签字",)),
    ("phone", ("电话", "电    话", "电  话")),
    ("bank", ("开户银行",)),
    ("account", ("账号", "账    号", "账  号")),
    ("credit_code", ("统一社会信用代码", "社会信用代码")),
]

PARTY_LABELS = ("需求方", "甲  方", "甲方")


def _cell_str(value) -> str:
    if value is None:
        return ""
    return str(value).strip()


def _row_is_empty(ws: Worksheet, row: int) -> bool:
    for col in range(1, ws.max_column + 1):
        if ws.cell(row, col).value not in (None, ""):
            return False
    return True


def _col_letter_to_index(letter: str) -> int:
    return openpyxl.utils.column_index_from_string(letter.upper())


def _find_template_sheet(wb) -> str:
    if "模板" in wb.sheetnames:
        return "模板"
    for name in wb.sheetnames:
        ws = wb[name]
        if _find_total_row(ws) is not None:
            return name
    raise ValueError("未在模板文件中识别到合同工作表（需包含「合计」行及产品公式区）")


def _find_party_cell(ws: Worksheet) -> str:
    for row in range(1, min(15, ws.max_row + 1)):
        for col in range(1, min(8, ws.max_column + 1)):
            label = _cell_str(ws.cell(row, col).value)
            if not label:
                continue
            if any(p in label.replace(" ", "") for p in PARTY_LABELS):
                value_col = col + 2
                if value_col <= ws.max_column:
                    letter = openpyxl.utils.get_column_letter(value_col)
                    return f"{letter}{row}"
    return "E2"


def _find_total_row(ws: Worksheet) -> Optional[int]:
    for row in range(1, ws.max_row + 1):
        for col in range(1, min(6, ws.max_column + 1)):
            text = _cell_str(ws.cell(row, col).value)
            if text == "合计" or (text.endswith("合计") and "大写" not in text):
                return row
    return None


def _parse_sum_range(ws: Worksheet, total_row: int) -> Optional[Tuple[int, int, int, int]]:
    """从合计行解析 SUM 公式，返回 (qty_col, amt_col, first_row, last_row)"""
    qty_col = None
    amt_col = None
    first_row = None
    last_row = None

    for col in range(1, ws.max_column + 1):
        val = ws.cell(total_row, col).value
        if not val or not isinstance(val, str) or not val.upper().startswith("=SUM"):
            continue
        m = SUM_PATTERN.match(val.strip())
        if not m:
            continue
        _, r1, _, r2 = m.groups()
        row_start = int(r1)
        row_end = int(r2)
        if first_row is None:
            first_row, last_row = row_start, row_end
        if qty_col is None:
            qty_col = col
        else:
            amt_col = col

    if first_row is None or last_row is None:
        return None
    if qty_col is None:
        qty_col = 7
    if amt_col is None:
        amt_col = 9
    return qty_col, amt_col, first_row, last_row


def _is_product_formula_row(ws: Worksheet, row: int) -> bool:
    val = ws.cell(row, 2).value
    if not val or not isinstance(val, str):
        return False
    return bool(PRODUCT_ROW_PATTERN.search(val))


def _refine_product_range(ws: Worksheet, first_row: int, last_row: int) -> Tuple[int, int]:
    """根据 B 列公式校正产品区起止行"""
    while first_row <= last_row and not _is_product_formula_row(ws, first_row):
        first_row += 1
    while last_row >= first_row and not _is_product_formula_row(ws, last_row):
        last_row -= 1
    if first_row > last_row:
        raise ValueError("未能识别产品明细公式区")
    return first_row, last_row


def _find_grand_total_row(ws: Worksheet, total_row: int) -> int:
    for row in range(total_row + 1, min(total_row + 5, ws.max_row + 1)):
        for col in range(1, 6):
            text = _cell_str(ws.cell(row, col).value)
            if "总合计" in text or "大写" in text:
                return row
    return total_row + 1


def _find_buyer_header_row(ws: Worksheet, start_row: int) -> Optional[int]:
    for row in range(start_row, ws.max_row + 1):
        for col in range(1, 8):
            text = _cell_str(ws.cell(row, col).value)
            if text.startswith("需方"):
                return row
    return None


def _writable_cell(ws: Worksheet, row: int, col: int):
    """若目标格在合并区域内，返回可写入的左上角单元格"""
    for merged in ws.merged_cells.ranges:
        if (
            merged.min_row <= row <= merged.max_row
            and merged.min_col <= col <= merged.max_col
        ):
            return ws.cell(merged.min_row, merged.min_col)
    return ws.cell(row, col)


def _resolve_value_column(ws: Worksheet, row: int, col: int) -> int:
    """解析实际应写入的列（合并单元格取左上角列）"""
    for merged in ws.merged_cells.ranges:
        if (
            merged.min_row <= row <= merged.max_row
            and merged.min_col <= col <= merged.max_col
        ):
            return merged.min_col
    return col


def _is_in_label_merge(ws: Worksheet, row: int, col: int) -> bool:
    """判断单元格是否属于从C列（标签列）起始的合并区域"""
    for merged in ws.merged_cells.ranges:
        if (
            merged.min_row <= row <= merged.max_row
            and merged.min_col == 3  # 从C列开始
            and merged.min_col <= col <= merged.max_col
        ):
            return True
    return False


def _find_buyer_fields(ws: Worksheet, start_row: int) -> Tuple[Dict[str, int], int]:
    """扫描需方字段行，返回 {语义键: 行号} 及值写入列"""
    fields: Dict[str, int] = {}
    value_col = 5

    for row in range(start_row, ws.max_row + 1):
        label = _cell_str(ws.cell(row, 3).value)
        if not label:
            continue
        if label.startswith("供方"):
            break
        for key, patterns in BUYER_LABEL_KEYS:
            if key in fields:
                continue
            if any(label.startswith(p) for p in patterns):
                fields[key] = row
                # 优先查找公式单元格（=E2或VLOOKUP等），再查非空且非标签合并的单元格
                formula_col = None
                fallback_col = None
                for col in range(4, 8):
                    cell = ws.cell(row, col)
                    if isinstance(cell.value, str) and cell.value.startswith("="):
                        if formula_col is None:
                            formula_col = _resolve_value_column(ws, row, col)
                    elif cell.value and not _cell_str(cell.value).startswith("="):
                        if fallback_col is None and not _is_in_label_merge(ws, row, col):
                            fallback_col = _resolve_value_column(ws, row, col)
                if formula_col is not None:
                    value_col = formula_col
                elif fallback_col is not None:
                    value_col = fallback_col
                break
    return fields, value_col


def _find_empty_rows_before_buyer(ws: Worksheet, from_row: int, buyer_header_row: int) -> List[int]:
    empty_rows: List[int] = []
    for row in range(from_row, buyer_header_row):
        if _row_is_empty(ws, row):
            empty_rows.append(row)
    return empty_rows


def analyze_template_layout(template_path: str, sheet_name: Optional[str] = None) -> TemplateLayout:
    """
    分析上传的销售订单合同模板，自动识别产品区、合计区、需方区位置。
    用户更换新模板时，只要结构类似（产品公式区 + 合计 + 需方信息表），即可自动适配。
    """
    wb = openpyxl.load_workbook(template_path, data_only=False)
    try:
        if sheet_name:
            if sheet_name not in wb.sheetnames:
                raise ValueError(f"工作表不存在: {sheet_name}")
            ws = wb[sheet_name]
            used_sheet = sheet_name
        else:
            used_sheet = _find_template_sheet(wb)
            ws = wb[used_sheet]

        total_row = _find_total_row(ws)
        if total_row is None:
            raise ValueError("未找到「合计」行，请确认模板格式")

        parsed = _parse_sum_range(ws, total_row)
        if parsed is None:
            raise ValueError("合计行未找到 SUM 公式，无法识别产品区范围")

        qty_col, amt_col, first_row, last_row = parsed
        first_row, last_row = _refine_product_range(ws, first_row, last_row)

        party_cell = _find_party_cell(ws)
        grand_total_row = _find_grand_total_row(ws, total_row)

        terms_end = grand_total_row
        for row in range(grand_total_row + 1, ws.max_row + 1):
            text = _cell_str(ws.cell(row, 3).value)
            if text.startswith("需方") or text.startswith("10、"):
                if text.startswith("10、"):
                    terms_end = row
                break
            if text:
                terms_end = row

        buyer_header = _find_buyer_header_row(ws, grand_total_row + 1)
        empty_before_buyer: List[int] = []
        buyer_fields: Dict[str, int] = {}
        buyer_value_col = 5

        if buyer_header:
            empty_before_buyer = _find_empty_rows_before_buyer(ws, terms_end + 1, buyer_header)
            scan_start = buyer_header
            buyer_fields, buyer_value_col = _find_buyer_fields(ws, scan_start)
        else:
            buyer_fields, buyer_value_col = _find_buyer_fields(ws, grand_total_row + 1)

        if not buyer_fields:
            raise ValueError("未找到需方信息区（单位名称/地址等标签）")

        return TemplateLayout(
            template_sheet=used_sheet,
            party_cell=party_cell,
            product_first_row=first_row,
            product_last_row=last_row,
            total_row=total_row,
            qty_sum_col=qty_col,
            amt_sum_col=amt_col,
            grand_total_row=grand_total_row,
            buyer_value_col=buyer_value_col,
            buyer_fields=buyer_fields,
            empty_rows_before_buyer=empty_before_buyer,
        )
    finally:
        wb.close()
