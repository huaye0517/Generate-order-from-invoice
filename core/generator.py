# -*- coding: utf-8 -*-
"""销售订单合同生成"""
import re
from pathlib import Path
from typing import Callable, List, Optional
from .catalog import group_by_company
from .invoice_map import buyer_info_from_record, load_invoice_records, lookup_invoice_record
from .models import CompanyGroup, GenerationResult
from .order_cache import load_order_rows
from .template_analyzer import analyze_template_layout
from .template_finalize import prepare_template_output
from .validator import count_product_lines, validate_company_group

INVALID_CHARS = re.compile(r'[\\/:*?"<>|]')


def sanitize_filename(name: str) -> str:
    cleaned = INVALID_CHARS.sub("_", str(name).strip())
    return cleaned or "未命名"


def build_output_filename(company_name: str, is_consistent: bool) -> str:
    base = sanitize_filename(company_name)
    if is_consistent:
        return f"{base}销售订单合同.xlsx"
    return f"{base}销售订单合同（数据不一致）.xlsx"


def generate_contracts(
    catalog_path: str,
    template_path: str,
    output_dir: str,
    progress_callback: Optional[Callable[[int, int, str], None]] = None,
) -> List[GenerationResult]:
    output_path = Path(output_dir)
    output_path.mkdir(parents=True, exist_ok=True)

    groups: List[CompanyGroup] = group_by_company(catalog_path, template_path)
    order_rows = load_order_rows(template_path)
    invoice_records = load_invoice_records(template_path)
    template_layout = analyze_template_layout(template_path)
    results: List[GenerationResult] = []
    total = len(groups)

    for idx, group in enumerate(groups, start=1):
        if progress_callback:
            progress_callback(idx, total, group.company_name)

        order_qty, order_amt, consistent, formula_warning, c_values = validate_company_group(order_rows, group)
        filename = build_output_filename(group.company_name, consistent)
        dest = output_path / filename

        if dest.exists():
            stem = dest.stem
            suffix = dest.suffix
            n = 2
            while dest.exists():
                dest = output_path / f"{stem}_{n}{suffix}"
                n += 1

        invoice_record = lookup_invoice_record(group.e2_value, group.customers, invoice_records)
        buyer_info = buyer_info_from_record(invoice_record, group.e2_value)
        actual_e2 = buyer_info.get("party", group.e2_value)

        # 修正E2（party）：确保其与订单C列值精确匹配，使COUNTIF公式能正常过滤产品行
        if c_values and actual_e2 not in c_values:
            # 优先选包含当前e2值的订单C列项，否则按字母序取第一个
            best_key = next(
                (v for v in sorted(c_values) if group.e2_value.lower() in v.lower()),
                sorted(c_values)[0],
            )
            buyer_info = dict(buyer_info)
            buyer_info["party"] = best_key
            actual_e2 = best_key

        # 修正后重新评估formula_warning（基于最终actual_e2）
        formula_warning = not c_values or actual_e2 not in c_values

        # 基于最终E2键计算有效产品行数：只统计匹配E2且有实际数据的行
        # 确保裁剪行数与公式实际显示行数一致，消除末尾空白行
        product_lines = count_product_lines(order_rows, [actual_e2])

        prepare_template_output(
            template_path,
            str(dest),
            buyer_info,
            product_lines,
            layout=template_layout,
        )

        status = "一致" if consistent else "数据不一致"
        message = ""
        if formula_warning:
            message = "警告：E2 与订单 C 列不完全匹配，Excel 中产品行可能为空"
        if not consistent:
            message = (
                f"目录合计 数量={group.catalog_quantity} 金额={group.catalog_amount}；"
                f"订单合计 数量={order_qty} 金额={order_amt}"
            )
            if formula_warning:
                message += "；E2 与订单 C 列不完全匹配"

        results.append(
            GenerationResult(
                company_name=group.company_name,
                e2_value=actual_e2,
                customers=group.customers,
                catalog_quantity=group.catalog_quantity,
                catalog_amount=group.catalog_amount,
                order_quantity=order_qty,
                order_amount=order_amt,
                is_consistent=consistent,
                formula_warning=formula_warning,
                output_filename=dest.name,
                output_path=str(dest),
                status=status,
                message=message,
            )
        )

    return results
