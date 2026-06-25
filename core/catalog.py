# -*- coding: utf-8 -*-
"""客户目录读取与公司分组"""
from typing import Dict, List
import openpyxl
from .invoice_map import load_invoice_mapping, resolve_company_name
from .models import CatalogRow, CompanyGroup

SKIP_NAMES = {"总计", "合计", "none", "none"}
COL_CUSTOMER = 1
COL_QTY = 2
COL_AMOUNT = 3


def _should_skip(name) -> bool:
    if name is None:
        return True
    text = str(name).strip()
    if not text:
        return True
    if text.lower() in SKIP_NAMES:
        return True
    if text.startswith("(") and "空白" in text:
        return True
    return False


def load_catalog_rows(catalog_path: str) -> List[CatalogRow]:
    """读取客户目录有效行"""
    wb = openpyxl.load_workbook(catalog_path, read_only=True, data_only=True)
    ws = wb.active
    rows: List[CatalogRow] = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row:
            continue
        name = row[COL_CUSTOMER - 1] if len(row) >= COL_CUSTOMER else None
        if _should_skip(name):
            continue
        qty = row[COL_QTY - 1] if len(row) >= COL_QTY else 0
        amt = row[COL_AMOUNT - 1] if len(row) >= COL_AMOUNT else 0
        try:
            qty_f = float(qty or 0)
        except (TypeError, ValueError):
            qty_f = 0.0
        try:
            amt_f = float(amt or 0)
        except (TypeError, ValueError):
            amt_f = 0.0
        rows.append(CatalogRow(str(name).strip(), qty_f, amt_f))
    wb.close()
    return rows


def group_by_company(catalog_path: str, template_path: str) -> List[CompanyGroup]:
    """按开票信息映射的公司分组并合并目录数量/金额"""
    catalog_rows = load_catalog_rows(catalog_path)
    invoice_mapping = load_invoice_mapping(template_path)
    groups: Dict[str, CompanyGroup] = {}

    for row in catalog_rows:
        company = resolve_company_name(row.customer_name, invoice_mapping)
        e2_value = company
        if company not in groups:
            groups[company] = CompanyGroup(
                company_name=company,
                e2_value=e2_value,
                customers=[],
                catalog_quantity=0.0,
                catalog_amount=0.0,
            )
        g = groups[company]
        if row.customer_name not in g.customers:
            g.customers.append(row.customer_name)
        g.catalog_quantity += row.quantity
        g.catalog_amount += row.amount

    return sorted(groups.values(), key=lambda x: x.company_name)
