# -*- coding: utf-8 -*-
"""模板结构数据模型"""
from dataclasses import dataclass, field
from typing import Dict, List, Optional


@dataclass
class TemplateLayout:
    """从上传的合同模板中自动识别出的布局信息"""

    template_sheet: str
    party_cell: str
    product_first_row: int
    product_last_row: int
    total_row: int
    qty_sum_col: int
    amt_sum_col: int
    grand_total_row: int
    buyer_value_col: int
    buyer_fields: Dict[str, int] = field(default_factory=dict)
    empty_rows_before_buyer: List[int] = field(default_factory=list)

    @property
    def product_capacity(self) -> int:
        return self.product_last_row - self.product_first_row + 1

    def last_product_row(self, product_line_count: int) -> int:
        if product_line_count <= 0:
            return self.product_first_row
        return min(
            self.product_first_row + product_line_count - 1,
            self.product_last_row,
        )

    def rows_to_delete_count(self, product_line_count: int) -> int:
        last_row = self.last_product_row(product_line_count)
        return max(0, self.product_last_row - last_row)

    def summary(self) -> str:
        lines = [
            f"工作表: {self.template_sheet}",
            f"需求方单元格: {self.party_cell}",
            f"产品区: 第 {self.product_first_row}–{self.product_last_row} 行（容量 {self.product_capacity}）",
            f"合计行: 第 {self.total_row} 行（数量列 {self.qty_sum_col}，金额列 {self.amt_sum_col}）",
            f"总合计行: 第 {self.grand_total_row} 行",
            f"需方信息列: 第 {self.buyer_value_col} 列",
            f"需方字段行: {self.buyer_fields}",
            f"需方前空白行: {self.empty_rows_before_buyer or '无'}",
        ]
        return "\n".join(lines)
