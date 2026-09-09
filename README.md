# 价格牌处理工具 · Python 项目

- **核心算法**用 Python 重写，便于在 IDE 里调试和修改
- **GUI**用 Tkinter（Python 自带），无需额外 GUI 框架
- **EXE**用 PyInstaller 一键打包，分发给最终用户

---

## v1.6 本次更新（尺寸开关 + 页面对齐参考样例 + 拖入导入）

- **标签尺寸开关**（Excel 页签）：新增「自定义标签尺寸」开关——
  - 开：可设置空白标签宽×高 cm，空白标签图上标注设置的尺寸（默认图会**先擦除烙印的旧 4.5/9cm 标注**再写新尺寸；自定义参考图则叠加红色尺寸标注）；
  - 关（默认）：默认 4.5×9.0，输入框置灰、不做任何标注。
- **汇总 PDF 页面对齐参考样例**（P260901923.pdf）：输出页面 = 标签物理尺寸（默认 **4.5×7.5cm**，可调），每页一枚、内容铺满页面；防切边 mm 只作为出血裁切余量、不再放大页面；内容墨迹 + 2mm 为页面自动下限，保证内容永不被裁。
- **拖入导入**：桌面版（tkinterdnd2，已 vendor 到 `vendor/`，缺失时自动降级）与网页版均支持把 Excel / PDF / 图片直接拖到窗口完成选择。
- 桌面 `main.py` v1.5；网页版 `价格牌处理工具_网页版2.html` 同步以上全部改动。

## v1.5 更新回顾（合并 Excel 三项修复 + 网页版改版）

- **空白标签标注尺寸**：未选参考图时，内置默认空白标签图会按用户设置的「标签宽 × 高 cm」叠加红色尺寸标注（外框 + 中心大字 + 底部宽度），桌面版与网页版一致。
- **示例图不再缺底部时间行**：检测阶段把紧贴标签的小字条带（≤20pt 间距）并入标签框——标签底部的打印日期行会完整保留；离标签较远的行间说明文字（如 Printing Quantity）仍然排除。桌面版 `detect_stickers` 与网页版 `detectStickers` 同步修复。
- **Excel 图片不再重叠**：按图片像素宽度计算列间距（桌面版列宽≈96px、网页版≈64px），空白标签图与各款式示例图依次排列互不覆盖，也不压住左侧表格；同时修复 openpyxl 字符串锚点忽略显示尺寸的问题（改用显式 OneCellAnchor + ext）。
- **网页版样式改版**：参考第一版《文档处理工具》风格——渐变蓝头部、编号卡片、横向页签、芯片式文件名展示。

## v1.4 更新回顾（多文件批量 + 合并输出）

- **Excel 合并转化**：支持**多选** Excel/CSV/XLS（每个款式一个文件），转换后**合并导出为一个 .xlsx**。
  - 不再展示原始宽表内容；每个款式一个竖排价格牌区块（标题 + 表头 + 主体 + 空白吊牌 + 合计），末尾追加所有款式的「总计」行。
  - 保留**标签尺寸**与**参考空白标签图**输入；未选参考图时，空白吊牌默认使用内置图 `assets/default_blank_tag.png`（4.5×9cm 样式）。
  - 可**多选价格签 PDF**：按款号匹配（文件名含款号，未匹配到时按顺序兜底配对），把每个款式 PDF 里的**第一枚价格牌渲染为示例图**，贴在空白标签右侧。
- **PDF 价格签重排**：支持**多选** PDF，合并输出**一个总 PDF**——按款式顺序（文件顺序）往后排，款式内**先按颜色、再按尺码**（XS→S→M→L→XL→XXL）。
  - 防切边默认 **3mm**（右侧提示「默认3，不建议小于3」）。
- **最终产物**：1 个汇总重排 PDF + 1 个汇总 Excel。

## 网页版（单文件 HTML）

`价格牌处理工具_网页版.html`（网页版2） —— 双击用浏览器打开即可，无需安装 Python/依赖，**文件全部在本地浏览器处理、不上传**。

| 页签 | 能力 |
|---|---|
| ① Excel 合并转化 | 多选 Excel/CSV/XLS → 页面内预览合并大表（连续主体 + 空白吊牌 + 合计 + 右侧空白标签图与每款式价格牌示例图）→ 导出合并 .xlsx |
| ② PDF 价格签重排 | 多选 PDF → 解析检测 → 合并输出**一个总 PDF**（款式 → 颜色 → 尺码排序） |

- 依赖通过 CDN 加载（SheetJS / ExcelJS / pdf.js / pdf-lib / JSZip），jsDelivr 失败自动切 unpkg，**首次打开需联网**。
- 默认空白标签图已内嵌（base64），无需额外资源。
- PDF 示例图按 **PO 号**与款式匹配（如 `2008600_PriceStickerApparel.pdf` ↔ G1945AX），其次按款号，最后按顺序兜底。
- 已修复：pdf-lib `embedPdf` 不传页码索引时只嵌入第 1 页，多页 PDF 重排第 2 页起会报 `Cannot read properties of undefined (reading 'h')`——现已显式传入全部页码索引。

> 桌面 Python 版与网页版算法同源：检测阈值 `DETECT_SCALE=3 / PAD_PX=30 / COL_GAP=90 / ROW_GAP=180`，生产数 `ceil(订单数 × (1+上浮%))`。

## 目录结构

```
kaifa/
├── main.py                # GUI 入口（双击运行 run.py 启动）
├── run.py                 # 启动快捷方式（等价于 python main.py）
├── core/                  # 核心业务模块（独立可测，与 GUI 解耦）
│   ├── __init__.py
│   ├── units.py           # 单位换算（cm/mm ↔ pt）
│   ├── pdf_rearrange.py   # PDF 检测 + 合并 + 重排，v1.3 起每枚严格居中且尺寸一致
│   ├── excel_pricecard.py # Excel → 竖排价格牌，v1.3 起导出与预览共用 SheetModel
│   ├── sheet_view.py      # Canvas 电子表格预览组件（Excel 预览「所见即所得」）
│   └── label_image.py     # 带尺寸标注的标签参考图（替代 Canvas）
├── tests/                 # 回归测试（用 samples 验证核心算法）
│   ├── test_pdf_rearrange.py
│   └── test_excel_pricecard.py
├── build_exe.py           # PyInstaller 打包脚本
├── requirements.txt       # Python 依赖
└── README.md              # 本文档
```

---

## v1.3 本次更新

- **PDF 价格签重排**：每枚标签以自身墨迹中心为圆心、统一标称尺寸裁切，缩放恒为 1.0，实现「完全居中 + 尺寸完全一致」。源页边缘越界时只绘制页面内部分，居中关系不变。
- **Excel 预览**：新增 Canvas 电子表格预览，与最终导出 `.xlsx` 共用同一 `SheetModel`，做到「预览即所得」；支持列标、行号、合并单元格、底色填充、标签参考图。
- **界面风格**：改为极简高级风格——左侧导航、顶栏、内容卡片、发丝级灰线、近黑主按钮、无 emoji，整体更简洁直接。

---

## 快速开始（开发环境运行）

### 前置条件

- Python 3.10+（[python.org](https://www.python.org/downloads/) 官方安装版，**自带 tkinter**）
- Windows / macOS / Linux 均可

> **Anaconda、Homebrew、`python -m venv` 创建的虚拟环境也都自带 tkinter**。
> 如果 `import tkinter` 报错，说明当前 Python 安装不完整，请重新安装官方版本（确保勾选 "tcl/tk and IDLE"）。

### 1. 安装依赖

```bash
cd kaifa
python -m pip install -r requirements.txt
```

> 推荐 Python 3.10+。本机已验证 3.13.12 与 3.9.0 均可用。

### 2. 启动 GUI

**最简单**：直接双击 `启动.bat`（自动定位 Python、装依赖、启动 GUI）。

或手动：
```bash
python run.py
```

或直接：

```bash
python main.py
```

启动后左侧有两个导航项：

| 导航项 | 功能 |
|---|---|
| ① Excel 生产数量统计转化 | 选择宽表 xls/xlsx/csv → 配置上浮比例和标签尺寸 → 预览 → 导出 .xlsx |
| ② PDF 价格签重排 | 选择多枚/页 PDF → 解析并检测 → 输出单枚/页或多枚/页重排 PDF |

### 3. 跑回归测试（验证核心算法与原 HTML 等价）

```bash
python -m tests.test_pdf_rearrange
python -m tests.test_excel_pricecard
```

两个测试均使用 `samples/` 中的真实文件：
- `samples/input.pdf`（多枚横排 PDF） → 检测 → 重排 → 与 `示例输出_价格签重排_4.5x9cm.pdf` 对照
- `samples/output.xls`（宽表） → 解析 → 转换 → 与 `示例输出_G1945AX-特殊价格牌.xlsx` 对照

---

## 打包 EXE

### 方法 1：双击 `build_exe.bat`

```bat
cd kaifa
build_exe.bat
```

会自动安装依赖 + 打包 EXE + 生成 NSIS 安装包 + 打开输出目录。产物：

| 文件 | 大小 | 说明 |
|---|---|---|
| `dist\价格牌处理工具.exe` | ~124 MB | 单文件 EXE，双击直接运行 |
| `dist\价格牌处理工具_安装程序_v1.3.exe` | ~123 MB | NSIS 制作的专业安装包，分发给最终用户 |

### 方法 2：手动

```bash
python -m pip install -r requirements.txt
python build_exe.py            # 仅打包 EXE
"C:\Program Files (x86)\NSIS\makensis.exe" installer.nsi   # 仅生成安装包
```

> **首次打包需 1-3 分钟**（PyInstaller 解包依赖）。后续改动后再次打包通常 30s 内完成。

> **EXE 体积**：单文件模式约 124 MB（PyMuPDF 较大）。若嫌大可改用 `--onedir` 模式（多文件，启动更快）。

### NSIS 安装包特性

- **Modern UI 2** 向导：欢迎页 → 许可协议 → 安装路径 → 安装进度 → 完成页（与腾讯会议/微信同款风格）
- **许可协议页**：显示 `LICENSE.txt`
- **完成页可勾选**：立即运行 + 创建桌面快捷方式
- **写入注册表**：控制面板"程序和功能"中可见，显示名称、版本、发布者、卸载命令、图标
- **快捷方式**：开始菜单（含"卸载"项）+ 桌面
- **静默卸载**：`Uninstall.exe /S` 完全清理（注册表 + 快捷方式 + 安装目录）
- **压缩**：LZMA /SOLID，安装包体积比 EXE 还小

> **NSIS 安装**：从 https://nsis.sourceforge.io/Download 下载 3.x（已含 makensis.exe）。本项目 `installer.nsi` 已设置 `Unicode true`、UTF-8 BOM、LZMA /SOLID，无需额外配置。

---

## 调试指南

### 业务逻辑与 GUI 解耦

所有算法都在 `core/`，可直接在 IDE 里单步调试：

```python
from core.pdf_rearrange import detect_stickers, rearrange_pdf

# 单步检测
import fitz
doc = fitz.open("samples/input.pdf")
boxes = detect_stickers(doc[0])   # 第 1 页检测结果
print(boxes)

# 单步重排
rearrange_pdf(
    input_path="samples/input.pdf",
    output_path="samples/_debug.pdf",
    per_page=1, margin_mm=0, padding_mm=4,
)
```

修改 `core/` 任意文件后，重启 GUI 即生效（无需重新打包 EXE）。

### 修改检测参数

打开 `core/pdf_rearrange.py`，顶部有 4 个常量：

```python
DETECT_SCALE = 3     # 检测放大倍率；越大越精细，但越慢
PAD_PX = 30          # 检测框外扩像素；约 3.5mm，防止切边同时避免页面过大
COL_GAP = 90         # 横向合并容差（像素）
ROW_GAP = 180        # 纵向合并容差（像素）
```

按实际 PDF 内容微调即可。

**坐标系注意**：PyMuPDF 的 PDF 坐标（`fitz.Rect`、`show_pdf_page` 等）与 Pillow 渲染图一致，原点都在左上角，y 轴向下增长。这与 pdf-lib（HTML 版）的"y 从底部向上"不同，因此移植时**不能**照搬 `page_h - y` 的写法。

### 修改 Excel 输出样式

打开 `core/excel_pricecard.py`，修改：
- `FONT_NAME = "宋体"` → 字体
- `FILL_TOTAL = "FFFFFF00"` → 合计行底色
- `FILL_S = "FFCCFFCC"` → S 码行底色
- `widths = [12, 8, 9, 9, 10, 10, 18, 16]` → 列宽
- `start_row = 10` / `start_col = 4` → 竖排表起始位置
- `img.width = 240` / `img.height = 480` → 嵌入参考图尺寸

---

**输出等价性**：
- PDF 重排：清晰度由源 PDF 决定（不重渲染），所以两版输出文件**视觉一致**
- Excel 价格牌：格式（合并、配色、列宽、嵌入图）按相同规则生成

### 已验证的运行结果

在 `samples/input.pdf` 上运行 `python -m tests.test_pdf_rearrange`：

| 指标 | 值 |
|---|---|
| 输入 | 2 页 / 12 枚横排标签 |
| 检测 | 12 枚（每页 6 枚，100% 检出） |
| 输出 | 12 页 / 单枚 |
| 输出页面尺寸 | **4.50 × 7.34 cm**（按检测到的实际标签尺寸自动确定） |
| 清晰度 | 与源 PDF 完全一致（每页含 246 字符原始文字） |

在用户提供文件 `2008600_PriceStickerApparel.pdf` 上验证：
- 修复前：顶部条码/数字"8"被截断，输出 5.21 × 8.04 cm
- 修复后：内容完整，输出 **4.50 × 7.34 cm**（接近参考 PDF 4.54 × 7.54 cm）

---

## 已知坑（2026-09 验证时发现）

### 1. PyMuPDF 1.28 的 Page 引用问题
`out.new_page()` 返回的 Page 在被存入 list 后**会失去 parent 引用**，导致后续 `show_pdf_page` 报 `'NoneType' object has no attribute 'is_pdf'`。

**解决**：在循环内每次通过 `out[i]` 重新取 Page，不要缓存引用。本项目 `core/pdf_rearrange.py` 已规避。

### 2. 旧 `.xls` 格式需要 xlrd 1.2.0
Python 生态中 `xlrd 2.0+` 已移除 `.xls` 支持，但 `.xls` 在国内仓库仍常见（导出工具默认格式）。本项目已通过按后缀自动选引擎支持 `.xls`：

- `.csv` → Python csv
- `.xls` → `xlrd==1.2.0`（**必须固定 1.2.0**，这是最后一个支持 `.xls` 的版本）
- `.xlsx` → openpyxl

`requirements.txt` 已固定 `xlrd==1.2.0`，`启动.bat` 会自动装。如果手动 `pip install` 后报 `ModuleNotFoundError: No module named 'xlrd'`，说明装到了别的 Python 环境里，参见下面的 Q&A。

### 3. PyMuPDF 坐标系与 pdf-lib 相反（本项目已踩的坑）
PyMuPDF 的 `fitz.Rect`、`show_pdf_page`、文本 bbox 等 API，原点都在页面**左上角**，y 轴向**下**增长。而原 HTML 版使用的 pdf-lib 原点在左下角，y 轴向上。

如果移植时直接照搬 HTML/JS 的 `page_h - y` 写法，会导致裁切框整体下移，**顶部内容被截断**。本项目 `core/pdf_rearrange.py` 已采用正确的 y 坐标方向。

### 4. PyMuPDF 的 `fitz` 别名已 deprecated
1.24+ 起 `import fitz` 仍能用但会出 warning，未来会移除。新写法是 `import pymupdf`。

本项目保留 `import fitz` 以兼容 PyMuPDF 1.23~1.28（社区主流版本）。

### 5. show_pdf_page 在 1.24+ 改为静态方法
旧：`page.show_pdf_page(rect, docsrc, pno, clip=...)`
新：`fitz.Page.show_pdf_page(page, rect, docsrc, pno, clip=...)`

本项目已采用新写法。

### 6. 标签参考图叠加冗余标注
**现象**：用户上传的参考图本身已带尺寸标注，Python 版 `make_annotated_from_bytes` 又在图上叠加红色外框、底部箭头、`4.5 × 9.0 cm` 大字，导致标注重复、外框过大、内容被遮挡。

**修复**（`core/label_image.py`）：
- 用户上传参考图时，仅做等比缩放并居中，**不再叠加任何额外标注**。
- 无参考图时仍用 `make_label_diagram` 生成标准尺寸示意图。
- 缩放策略由 `max()` 改为 `min()`，保证原图内容**完整不被裁切**。

---

## 常见问题

### Q: 报错 `ModuleNotFoundError: No module named 'fitz'`
说明 `pymupdf` 没装在当前 Python 环境里。原因通常是：
1. **多 Python 版本冲突**：Windows 上 `python` 可能指向某个 Python（managed/Anaconda/官方），而 `pip install` 装在了另一个。
2. **`pip install` 因网络问题静默失败**：默认源访问慢。

**快速排查**：
```cmd
python -c "import sys; print(sys.executable)"
python -c "import fitz; print(fitz.__doc__[:50])"
```
两个命令输出的 Python 路径应一致，否则 `import fitz` 必然失败。

**解决**：
```cmd
cd "C:\Users\Ruth\Desktop\价格牌处理\kaifa"
python -m pip install -i https://pypi.tuna.tsinghua.edu.cn/simple --timeout 120 -r requirements.txt
```
（用国内源 + 长 timeout，安装成功率显著提升）

**最省事**：直接双击 `启动.bat`，它会自动定位 Python 并装齐依赖。

### Q: 报错 `openpyxl does not support the old .xls file format`
旧 `.xls` 需要 `xlrd 1.2.0`（最后一个支持 `.xls` 的版本，2.0+ 已移除）。

**解决**：
```cmd
python -m pip install -i https://pypi.tuna.tsinghua.edu.cn/simple --timeout 120 "xlrd==1.2.0"
```

装好后本项目会自动按文件后缀选引擎：`.csv` / `.xls` / `.xlsx` 都支持。`requirements.txt` 已固定此版本。

如果你的输入都是 `.xlsx`，也可以跳过这一步，把 `core/excel_pricecard.py` 的 xlrd 分支删掉以减小打包体积。

### Q: EXE 启动很慢？
PyInstaller 单文件模式需要解压到临时目录，首次启动 5-10s。改用 `--onedir` 模式可秒开：

修改 `build_exe.py` 的 args：
```python
# "--onefile",
"--onedir",
```

### Q: 打包时被杀软误报？
在 `build_exe.py` 已加 `--noupx`（不压缩），多数情况可避免。仍误报请给 EXE 加数字签名或换 `--onedir`。

### Q: Tkinter 界面在高分屏模糊？
Windows 下可在 `main.py` 顶部加：
```python
from ctypes import windll
windll.shcore.SetProcessDpiAwareness(1)
```

### Q: 中文 PDF 检测失败？
PyMuPDF 默认支持中文路径与中文文本。如遇乱码，检查源 PDF 字体是否内嵌。

---

## 依赖版本（已验证）

| 包 | 版本 | 用途 |
|---|---|---|
| Python | 3.10+ | 运行 |
| PyMuPDF | ≥1.23.0 | PDF 处理 |
| openpyxl | ≥3.1.2 | Excel 处理（.xlsx） |
| xlrd | **1.2.0** | Excel 处理（.xls 旧格式，必须固定此版本） |
| Pillow | ≥10.0.0 | 图像处理 |
| pyinstaller | ≥6.0 | 打包 EXE |

---

## License

内部使用工具，未指定开源协议。# biaoqianchuli
# biaoqianchuli
# biaoqianchuli
