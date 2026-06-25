# -*- coding: utf-8 -*-
"""数据模型"""
from dataclasses import dataclass, field
from typing import List


@dataclass
class CatalogRow:
    """客户目录单行"""
    customer_name: str
    quantity: float
    amount: float


@dataclass
class CompanyGroup:
    """按公司分组后的数据"""
    company_name: str
    e2_value: str
    customers: List[str] = field(default_factory=list)
    catalog_quantity: float = 0.0
    catalog_amount: float = 0.0


@dataclass
class GenerationResult:
    """单个公司生成结果"""
    company_name: str
    e2_value: str
    customers: List[str]
    catalog_quantity: float
    catalog_amount: float
    order_quantity: float
    order_amount: float
    is_consistent: bool
    formula_warning: bool
    output_filename: str
    output_path: str
    status: str
    message: str = ""
