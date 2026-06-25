# -*- coding: utf-8 -*-
"""模板结构数据模型"""
from dataclasses import dataclass, field
from typing import Dict, List, Tuple


@dataclass
class TemplateLayout:
    """从上传的合同模板中自动识别出的布局信息（适配不同模板结构）"""

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

    # 总合计（大写）公式所在列及合并范围（从模板自动识别，非固定 E 列）
    grand_formula_col: int = 5
    grand_formula_merge_min_col: int = 5
    grand_formula_merge_max_col: int = 5

    # 产品行中「产品名称」类合并列（删行后需清除的幽灵合并，如 D:E）
    product_row_merges: List[Tuple[int, int]] = field(default_factory=list)

    # 需方信息区标签列（如 C 列），供方区块起始列（如 F 列）
    buyer_label_col: int = 3
    buyer_header_row: int = 0
    supplier_start_col: int = 6

    # 条款区末行（用于定位需方前空白行）
    terms_last_row: int = 0

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
            f"总合计行: 第 {self.grand_total_row} 行（大写公式列 {self.grand_formula_col}，"
            f"合并 {self.grand_formula_merge_min_col}–{self.grand_formula_merge_max_col}）",
            f"产品行合并: {self.product_row_merges or '无'}",
            f"需方标签列: {self.buyer_label_col}，值列: {self.buyer_value_col}，供方起始列: {self.supplier_start_col}",
            f"需方字段行: {self.buyer_fields}",
            f"需方前空白行: {self.empty_rows_before_buyer or '无'}",
        ]
        return "\n".join(lines)
