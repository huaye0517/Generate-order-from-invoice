# 销售订单合同生成工具

按公司自动生成销售订单合同 Excel 的桌面工具。

## 使用方式

1. 安装依赖：`pip install -r requirements.txt`
2. 运行：`python main.py`
3. 选择客户目录 Excel、合同模板 Excel、输出文件夹，点击「开始生成」

## 打包

双击 `build.bat`，或在项目目录执行：

```
pyinstaller build.spec --noconfirm
```

打包结果位于 `dist/销售订单合同生成工具/`。

## 规则说明

- 按开票信息映射的公司去重，每家公司 1 份合同
- E2 写入开票信息中的公司全称
- 目录与订单数量/金额不一致时，文件名追加「（数据不一致）」
