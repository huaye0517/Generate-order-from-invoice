# -*- coding: utf-8 -*-
"""销售订单合同生成工具 - 入口"""
from ui.app import ContractGeneratorApp


def main():
    app = ContractGeneratorApp()
    app.mainloop()


if __name__ == "__main__":
    main()
