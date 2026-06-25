# -*- coding: utf-8 -*-
"""订单数据校验"""
from typing import List, Set, Tuple
from .models import CompanyGroup
from .order_cache import RowTuple

TOLERANCE = 0.01


def count_product_lines(order_rows: List[RowTuple], customer_names: List[str]) -> int:
    """统计该公司关联客户的订单明细行数（即产品行数）"""
    name_set = {str(n).strip() for n in customer_names}
    count = 0
    for c_str, k_str, _qty, _amt in order_rows:
        if k_str in name_set or c_str in name_set:
            count += 1
    return count


def sum_order_for_customers(order_rows: List[RowTuple], customer_names: List[str]) -> Tuple[float, float, Set[str]]:
    name_set = {str(n).strip() for n in customer_names}
    total_qty = 0.0
    total_amt = 0.0
    c_values: Set[str] = set()

    for c_str, k_str, qty, amt in order_rows:
        if k_str in name_set or c_str in name_set:
            if c_str:
                c_values.add(c_str)
            total_qty += qty
            total_amt += amt

    return total_qty, total_amt, c_values


def is_consistent(catalog_qty: float, catalog_amt: float, order_qty: float, order_amt: float) -> bool:
    return abs(catalog_qty - order_qty) <= TOLERANCE and abs(catalog_amt - order_amt) <= TOLERANCE


def check_formula_warning(e2_value: str, order_c_values: Set[str]) -> bool:
    if not order_c_values:
        return True
    return str(e2_value).strip() not in order_c_values


def validate_company_group(order_rows: List[RowTuple], group: CompanyGroup) -> Tuple[float, float, bool, bool, Set[str]]:
    order_qty, order_amt, c_values = sum_order_for_customers(order_rows, group.customers)
    consistent = is_consistent(group.catalog_quantity, group.catalog_amount, order_qty, order_amt)
    formula_warning = check_formula_warning(group.e2_value, c_values)
    return order_qty, order_amt, consistent, formula_warning, c_values
