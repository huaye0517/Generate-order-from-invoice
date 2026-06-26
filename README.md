# 销售订单合同生成工具

按公司自动生成销售订单合同 Excel 的桌面工具。

## 使用方式

1. 安装依赖：`pip install -r requirements.txt`（Mac/Linux 可用 `python3 -m pip install -r requirements.txt`）
2. 运行：`python main.py`（Mac/Linux 可用 `python3 main.py`）
3. 选择客户目录 Excel、合同模板 Excel、输出文件夹，点击「开始生成」
4. 生成完成后可点击「打开输出文件夹」，在 Windows / macOS / Linux 下均会调用系统文件管理器打开目录

## 打包

### Windows

双击 `build.bat`，或在项目目录执行：

```bash
py -3 -m pip install -r requirements.txt
py -3 -m PyInstaller build.spec --noconfirm
```

打包结果：`dist/销售订单合同生成工具/销售订单合同生成工具.exe`

### macOS

**必须在 Mac 本机打包**（PyInstaller 不支持从 Windows 交叉编译 Mac 应用）。

1. 安装 Python 3.10+（建议从 [python.org](https://www.python.org/downloads/macos/) 或 Homebrew 安装）
2. 在项目目录执行：

```bash
chmod +x build.sh
./build.sh
```

或手动：

```bash
python3 -m pip install -r requirements.txt
python3 -m PyInstaller build.spec --noconfirm
```

打包结果：`dist/销售订单合同生成工具/销售订单合同生成工具`（无 `.exe` 后缀）

**首次打开安全提示**：若系统提示「无法验证开发者」，可任选其一：

- 右键该可执行文件 →「打开」→ 确认打开
- 或在终端执行：`xattr -cr "dist/销售订单合同生成工具/销售订单合同生成工具"`

Apple Silicon（M 系列）请使用 arm64 版 Python 打包，Intel Mac 请使用 x86_64 版 Python。

#### M 芯片 Mac 打包与分发（推荐流程）

适用于 Apple Silicon（M1/M2/M3/M4），打包一次后可在同类型 Mac 上直接使用，无需安装 Python。

**一、在 M 芯片 Mac 上准备（只需做一次）**

1. 安装 **arm64 版 Python 3.10+**
   - 推荐：[python.org/downloads/macos](https://www.python.org/downloads/macos/) → 选 **macOS 64-bit universal2 installer** 或 **Apple Silicon** 版
   - 或 Homebrew：`brew install python@3.12`
2. 确认是 arm64（终端执行）：
   ```bash
   python3 -c "import platform; print(platform.machine())"
   ```
   应输出 `arm64`。
3. 把项目拷到 Mac（Git clone / 网盘 / U 盘），进入项目目录：
   ```bash
   cd ~/Downloads/Generate-order-from-invoice   # 按实际路径修改
   chmod +x build.sh
   ./build.sh
   ```
4. 打包完成后，产物在：
   ```
   dist/销售订单合同生成工具/
   ├── 销售订单合同生成工具      ← 主程序（双击或终端运行）
   └── _internal/               ← 依赖，不可删
   ```

**二、首次运行（解除安全拦截）**

Finder 中 **右键** `销售订单合同生成工具` → **打开** → 再点 **打开**。

或在终端（在项目目录下）：
```bash
xattr -cr "dist/销售订单合同生成工具/销售订单合同生成工具"
open "dist/销售订单合同生成工具/销售订单合同生成工具"
```

**三、分发给其他 M 芯片 Mac**

1. 将整个文件夹 `dist/销售订单合同生成工具/` 打成 zip 发送（不要只发单个可执行文件）
2. 对方解压后，同样 **右键 → 打开** 运行一次即可
3. 可拖到「应用程序」或 Dock，以后双击使用

**四、日常使用**

1. 打开「销售订单合同生成工具」
2. 选择：客户目录 Excel → 合同模板 Excel → 输出文件夹
3. 点击「开始生成」→「打开输出文件夹」查看结果

**五、从 Windows 同步新版本**

Windows 上改完代码后，用 Git / 网盘把项目同步到 Mac，在 Mac 上重新执行 `./build.sh`，再分发新的 `dist/销售订单合同生成工具/` 文件夹即可。

> 注意：Windows 打的 `.exe` 不能在 Mac 上用；M 芯片打的包也不能在 Intel Mac 上运行（除非用 Rosetta + x86_64 Python 重新打包）。

### Linux

与 macOS 相同，使用 `build.sh` 或 `python3 -m PyInstaller build.spec --noconfirm` 在本机打包。

## 模板自适应裁剪

工具会在每次生成前**自动分析**您上传的合同模板布局（产品区、合计行、需方信息区等），再按实际产品行数删除留白行。**更换新模板时无需改代码**。

更换新模板后，建议先运行分析脚本确认识别结果：

```bash
python scripts/analyze_template.py sample/销售订单合同.xlsx
python scripts/analyze_template.py 你的新模板.xlsx
```

识别规则概要：
- 产品区：合计行 `SUM` 公式范围 + B 列 `ROW(A` 公式行
- 合计/总合计：含「合计」「总合计」标签的行
- 需方区：「单位名称」「单位地址」等标签行
- 留白：需方标题前的空行、产品区末尾多余行

## 规则说明

- 按开票信息映射的公司去重，每家公司 1 份合同
- E2 写入开票信息中的公司全称
- 目录与订单数量/金额不一致时，文件名追加「（数据不一致）」
