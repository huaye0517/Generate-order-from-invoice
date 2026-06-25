# -*- coding: utf-8 -*-
"""开票信息映射"""
from typing import Dict, List, Optional, Tuple
import openpyxl


def load_invoice_mapping(template_path: str) -> List[Tuple[str, str]]:
    """读取开票信息表 A列(合同名称) -> C列(单位名称)"""
    wb = openpyxl.load_workbook(template_path, read_only=True, data_only=True)
    if "开票信息" not in wb.sheetnames:
        wb.close()
        return []
    ws = wb["开票信息"]
    mapping: List[Tuple[str, str]] = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row:
            continue
        a_val = row[0] if len(row) > 0 else None
        c_val = row[2] if len(row) > 2 else None
        if a_val is None and c_val is None:
            continue
        a_str = str(a_val).strip().lstrip("\ufeff") if a_val is not None else ""
        c_str = str(c_val).strip() if c_val is not None else ""
        if a_str or c_str:
            mapping.append((a_str, c_str))
    wb.close()
    return mapping


def resolve_company_name(customer_name: str, mapping: List[Tuple[str, str]]) -> str:
    """目录客户名 -> 公司名（开票信息 C 列）"""
    cat = str(customer_name).strip()
    if not cat:
        return cat

    # 精确匹配 A 列
    for a_str, c_str in mapping:
        if a_str and a_str.lower() == cat.lower() and c_str:
            return c_str

    # 前缀/包含匹配（A 列较短时）
    best: Optional[str] = None
    best_len = -1
    for a_str, c_str in mapping:
        if not a_str or not c_str:
            continue
        a_lower = a_str.lower()
        cat_lower = cat.lower()
        if cat_lower.startswith(a_lower) or a_lower in cat_lower:
            if len(a_str) > best_len:
                best = c_str
                best_len = len(a_str)
    if best:
        return best

    return cat
