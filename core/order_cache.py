# -*- coding: utf-8 -*-
"""订单表缓存，避免重复扫描大表"""
from typing import List, Tuple
import openpyxl

RowTuple = Tuple[str, str, float, float]


def load_order_rows(template_path: str) -> List[RowTuple]:
    wb = openpyxl.load_workbook(template_path, read_only=True, data_only=True)
    ws = wb["订单"]
    rows: List[RowTuple] = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row:
            continue
        c_val = row[2] if len(row) > 2 else ""
        k_val = row[10] if len(row) > 10 else ""
        c_str = str(c_val).strip() if c_val is not None else ""
        k_str = str(k_val).strip() if k_val is not None else ""
        qty = float(row[5] or 0) if len(row) > 5 and isinstance(row[5], (int, float)) else 0.0
        amt = float(row[7] or 0) if len(row) > 7 and isinstance(row[7], (int, float)) else 0.0
        if not isinstance(row[5], (int, float)):
            try:
                qty = float(row[5] or 0)
            except (TypeError, ValueError):
                qty = 0.0
        if not isinstance(row[7], (int, float)):
            try:
                amt = float(row[7] or 0)
            except (TypeError, ValueError):
                amt = 0.0
        rows.append((c_str, k_str, qty, amt))
    wb.close()
    return rows
