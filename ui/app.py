# -*- coding: utf-8 -*-
"""销售订单合同生成工具 GUI"""
import os
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from typing import List, Optional
import customtkinter as ctk
from core.generator import generate_contracts
from core.models import GenerationResult

# 蓝色主题
COLOR_PRIMARY = "#0056B3"
COLOR_PRIMARY_HOVER = "#004494"
COLOR_LIGHT = "#E8F0FE"
COLOR_BG = "#F0F2F5"
COLOR_CARD = "#FFFFFF"
COLOR_TEXT_SUB = "#666666"


class ContractGeneratorApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        ctk.set_appearance_mode("light")
        ctk.set_default_color_theme("blue")

        self.title("销售订单合同生成工具")
        self.geometry("980x720")
        self.minsize(900, 650)
        self.configure(fg_color=COLOR_BG)

        self.catalog_path = tk.StringVar(value="尚未选择文件")
        self.template_path = tk.StringVar(value="尚未选择文件")
        self.output_dir = tk.StringVar(value="")
        self.status_text = tk.StringVar(value="请选择客户目录和合同模板后开始生成")
        self.progress_value = tk.DoubleVar(value=0.0)
        self.progress_label = tk.StringVar(value="0%")

        self.stat_total = tk.StringVar(value="--")
        self.stat_success = tk.StringVar(value="--")
        self.stat_mismatch = tk.StringVar(value="--")

        self.results: List[GenerationResult] = []
        self._generating = False

        self._build_ui()

    def _build_ui(self):
        header = ctk.CTkFrame(self, fg_color=COLOR_PRIMARY, corner_radius=0, height=72)
        header.pack(fill="x")
        header.pack_propagate(False)
        ctk.CTkLabel(
            header,
            text="销售订单合同生成工具",
            font=ctk.CTkFont(size=22, weight="bold"),
            text_color="white",
        ).pack(anchor="w", padx=24, pady=(14, 0))
        ctk.CTkLabel(
            header,
            text="上传客户目录与合同模板，按公司自动生成销售订单合同 Excel",
            font=ctk.CTkFont(size=12),
            text_color="#D0E4FF",
        ).pack(anchor="w", padx=24)

        body = ctk.CTkFrame(self, fg_color="transparent")
        body.pack(fill="both", expand=True, padx=20, pady=16)

        self._build_file_section(body)
        self._build_status_section(body)
        self._build_preview_section(body)

    def _card(self, parent, title: str) -> ctk.CTkFrame:
        card = ctk.CTkFrame(parent, fg_color=COLOR_CARD, corner_radius=12, border_width=1, border_color="#E0E0E0")
        card.pack(fill="x", pady=(0, 12))
        ctk.CTkLabel(card, text=title, font=ctk.CTkFont(size=14, weight="bold"), anchor="w").pack(
            anchor="w", padx=16, pady=(12, 8)
        )
        inner = ctk.CTkFrame(card, fg_color="transparent")
        inner.pack(fill="x", padx=16, pady=(0, 14))
        return inner

    def _build_file_section(self, parent):
        inner = self._card(parent, "文件操作")

        row1 = ctk.CTkFrame(inner, fg_color="transparent")
        row1.pack(fill="x", pady=4)
        ctk.CTkEntry(row1, textvariable=self.catalog_path, height=36).pack(side="left", fill="x", expand=True, padx=(0, 8))
        ctk.CTkButton(row1, text="选择目录 Excel", width=130, height=36, fg_color="white", text_color=COLOR_PRIMARY,
                      border_width=1, border_color=COLOR_PRIMARY, hover_color=COLOR_LIGHT,
                      command=self._pick_catalog).pack(side="left")

        row2 = ctk.CTkFrame(inner, fg_color="transparent")
        row2.pack(fill="x", pady=4)
        ctk.CTkEntry(row2, textvariable=self.template_path, height=36).pack(side="left", fill="x", expand=True, padx=(0, 8))
        ctk.CTkButton(row2, text="选择模板 Excel", width=130, height=36, fg_color="white", text_color=COLOR_PRIMARY,
                      border_width=1, border_color=COLOR_PRIMARY, hover_color=COLOR_LIGHT,
                      command=self._pick_template).pack(side="left")

        row3 = ctk.CTkFrame(inner, fg_color="transparent")
        row3.pack(fill="x", pady=(8, 0))
        ctk.CTkButton(row3, text="选择输出文件夹", width=130, height=36, fg_color="white", text_color=COLOR_PRIMARY,
                      border_width=1, border_color=COLOR_PRIMARY, hover_color=COLOR_LIGHT,
                      command=self._pick_output).pack(side="left")
        ctk.CTkLabel(row3, textvariable=self.output_dir, text_color=COLOR_TEXT_SUB, anchor="w").pack(
            side="left", fill="x", expand=True, padx=12
        )
        ctk.CTkButton(row3, text="开始生成", width=110, height=36, fg_color=COLOR_PRIMARY,
                      hover_color=COLOR_PRIMARY_HOVER, command=self._start_generate).pack(side="right", padx=(8, 0))
        ctk.CTkButton(row3, text="打开输出目录", width=110, height=36, fg_color=COLOR_PRIMARY,
                      hover_color=COLOR_PRIMARY_HOVER, command=self._open_output).pack(side="right")

    def _build_status_section(self, parent):
        inner = self._card(parent, "生成状态")
        ctk.CTkLabel(inner, text="请选择客户目录 Excel 和合同模板后开始生成", text_color=COLOR_TEXT_SUB, anchor="w").pack(
            anchor="w", pady=(0, 8)
        )

        prog_row = ctk.CTkFrame(inner, fg_color="transparent")
        prog_row.pack(fill="x")
        ctk.CTkLabel(prog_row, textvariable=self.progress_label, width=40).pack(side="right")
        self.progress_bar = ctk.CTkProgressBar(prog_row, progress_color=COLOR_PRIMARY, height=10)
        self.progress_bar.pack(side="left", fill="x", expand=True, padx=(0, 8))
        self.progress_bar.set(0)

        ctk.CTkLabel(inner, textvariable=self.status_text, text_color=COLOR_TEXT_SUB, anchor="w").pack(
            anchor="w", pady=(8, 10)
        )

        stats = ctk.CTkFrame(inner, fg_color="transparent")
        stats.pack(fill="x")
        for title, var in [("公司总数", self.stat_total), ("生成成功", self.stat_success), ("数据不一致", self.stat_mismatch)]:
            box = ctk.CTkFrame(stats, fg_color=COLOR_LIGHT, corner_radius=8)
            box.pack(side="left", fill="x", expand=True, padx=4)
            ctk.CTkLabel(box, text=title, text_color=COLOR_TEXT_SUB, font=ctk.CTkFont(size=11)).pack(pady=(8, 0))
            ctk.CTkLabel(box, textvariable=var, font=ctk.CTkFont(size=18, weight="bold"), text_color=COLOR_PRIMARY).pack(pady=(0, 8))

    def _build_preview_section(self, parent):
        card = ctk.CTkFrame(parent, fg_color=COLOR_CARD, corner_radius=12, border_width=1, border_color="#E0E0E0")
        card.pack(fill="both", expand=True)

        head = ctk.CTkFrame(card, fg_color="transparent")
        head.pack(fill="x", padx=16, pady=(12, 6))
        ctk.CTkLabel(head, text="结果预览", font=ctk.CTkFont(size=14, weight="bold")).pack(side="left")
        ctk.CTkLabel(head, text="最多显示前 200 行", text_color=COLOR_TEXT_SUB).pack(side="right")

        table_wrap = ctk.CTkFrame(card, fg_color="transparent")
        table_wrap.pack(fill="both", expand=True, padx=12, pady=(0, 12))

        columns = ("company", "customers", "e2", "cat_qty", "cat_amt", "ord_qty", "ord_amt", "status", "file")
        self.tree = ttk.Treeview(table_wrap, columns=columns, show="headings", height=12)
        headings = {
            "company": "公司名",
            "customers": "关联客户",
            "e2": "E2值",
            "cat_qty": "目录数量",
            "cat_amt": "目录金额",
            "ord_qty": "订单数量",
            "ord_amt": "订单金额",
            "status": "状态",
            "file": "输出文件",
        }
        widths = {"company": 120, "customers": 140, "e2": 120, "cat_qty": 70, "cat_amt": 80,
                  "ord_qty": 70, "ord_amt": 80, "status": 80, "file": 180}
        for col in columns:
            self.tree.heading(col, text=headings[col])
            self.tree.column(col, width=widths[col], anchor="center")

        yscroll = ttk.Scrollbar(table_wrap, orient="vertical", command=self.tree.yview)
        xscroll = ttk.Scrollbar(table_wrap, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=yscroll.set, xscrollcommand=xscroll.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        yscroll.grid(row=0, column=1, sticky="ns")
        xscroll.grid(row=1, column=0, sticky="ew")
        table_wrap.grid_rowconfigure(0, weight=1)
        table_wrap.grid_columnconfigure(0, weight=1)

    def _pick_catalog(self):
        path = filedialog.askopenfilename(title="选择客户目录 Excel", filetypes=[("Excel 文件", "*.xlsx *.xls")])
        if path:
            self.catalog_path.set(path)

    def _pick_template(self):
        path = filedialog.askopenfilename(title="选择销售订单合同模板", filetypes=[("Excel 文件", "*.xlsx *.xls")])
        if path:
            self.template_path.set(path)

    def _pick_output(self):
        path = filedialog.askdirectory(title="选择输出文件夹")
        if path:
            self.output_dir.set(path)

    def _open_output(self):
        out = self.output_dir.get().strip()
        if not out or not os.path.isdir(out):
            messagebox.showwarning("提示", "请先选择有效的输出文件夹")
            return
        os.startfile(out)

    def _validate_inputs(self) -> bool:
        cat = self.catalog_path.get().strip()
        tpl = self.template_path.get().strip()
        out = self.output_dir.get().strip()
        if cat in ("", "尚未选择文件") or not os.path.isfile(cat):
            messagebox.showwarning("提示", "请选择有效的客户目录 Excel 文件")
            return False
        if tpl in ("", "尚未选择文件") or not os.path.isfile(tpl):
            messagebox.showwarning("提示", "请选择有效的合同模板 Excel 文件")
            return False
        if not out:
            messagebox.showwarning("提示", "请选择输出文件夹")
            return False
        return True

    def _start_generate(self):
        if self._generating:
            return
        if not self._validate_inputs():
            return
        self._generating = True
        self.status_text.set("正在生成，请稍候...")
        self.progress_bar.set(0)
        self.progress_label.set("0%")
        thread = threading.Thread(target=self._run_generate, daemon=True)
        thread.start()

    def _run_generate(self):
        try:
            results = generate_contracts(
                self.catalog_path.get().strip(),
                self.template_path.get().strip(),
                self.output_dir.get().strip(),
                progress_callback=self._on_progress,
            )
            self.after(0, lambda: self._on_complete(results, None))
        except Exception as exc:
            self.after(0, lambda: self._on_complete([], exc))

    def _on_progress(self, current: int, total: int, company: str):
        pct = current / total if total else 0
        self.after(0, lambda: self._update_progress(pct, current, total, company))

    def _update_progress(self, pct: float, current: int, total: int, company: str):
        self.progress_bar.set(pct)
        self.progress_label.set(f"{int(pct * 100)}%")
        self.status_text.set(f"正在处理 {current}/{total}：{company}")

    def _on_complete(self, results: List[GenerationResult], error: Optional[Exception]):
        self._generating = False
        if error:
            self.status_text.set(f"生成失败：{error}")
            messagebox.showerror("生成失败", str(error))
            return

        self.results = results
        success = len(results)
        mismatch = sum(1 for r in results if not r.is_consistent)
        self.stat_total.set(str(success))
        self.stat_success.set(str(success))
        self.stat_mismatch.set(str(mismatch))
        self.progress_bar.set(1.0)
        self.progress_label.set("100%")
        self.status_text.set(f"生成完成，共 {success} 份，其中 {mismatch} 份数据不一致")
        self._fill_preview(results)
        messagebox.showinfo("完成", f"已生成 {success} 份销售订单合同\n数据不一致：{mismatch} 份")

    def _fill_preview(self, results: List[GenerationResult]):
        for item in self.tree.get_children():
            self.tree.delete(item)
        for r in results[:200]:
            customers = "、".join(r.customers[:3])
            if len(r.customers) > 3:
                customers += f" 等{len(r.customers)}个"
            self.tree.insert(
                "",
                "end",
                values=(
                    r.company_name,
                    customers,
                    r.e2_value,
                    f"{r.catalog_quantity:g}",
                    f"{r.catalog_amount:g}",
                    f"{r.order_quantity:g}",
                    f"{r.order_amount:g}",
                    r.status,
                    r.output_filename,
                ),
            )
