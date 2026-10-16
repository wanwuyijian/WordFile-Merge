# Word 合并工具（WordMergeTool）

把**不同格式**的文件合并成**一个 Word 文档**：Word、PowerPoint、Excel、PDF、图片、Markdown、纯文本……
一个都不少地搬进同一份 `.docx`，并且**全程不修改任何源文件**。

![版本](https://img.shields.io/badge/version-1.0.0-blue) ![平台](https://img.shields.io/badge/platform-Windows%2010%20%7C%2011-lightgrey)

---

## 它解决什么问题

手上有一堆格式各异的材料——一份 Word 报告、几张现场照片、一个 PDF 合同、一个 Excel 明细表、
一份 PPT 汇报——要合成**一份**能打印、能归档、能发给别人的 Word 文档。
手工复制粘贴会丢格式、图片会跑版、宽表格会超出纸面。

这个工具就是干这件事的，并且对**排版**有硬性保证：任何内容都不出页面。

---

## 快速开始

### 方式一：直接用打包好的 exe（推荐）

1. 把 `WordMergeTool.exe` 拷到任意 Windows 10 / 11 电脑（**无需安装 Python、Office、任何依赖**）
2. 双击运行
3. 把文件拖进窗口，或点「添加文件 / 添加文件夹」
4. **拖动缩略图**调整合并顺序
5. 点「开始合并」，选择输出位置

### 方式二：从源码运行

```powershell
# 安装依赖
python -m pip install python-docx python-pptx openpyxl pymupdf pillow tkinterdnd2

# 启动图形界面
cd <项目目录>
$env:PYTHONPATH = "$PWD\src"
python -m wmtool
```

### 方式三：命令行（适合批处理）

```powershell
# 合并整个目录
WordMergeTool-CLI.exe --input "D:\资料" --output "D:\合并结果.docx"

# 按指定顺序合并多个文件
WordMergeTool-CLI.exe --files 封面.docx 正文.pdf 附件.xlsx --output out.docx

# 列出所有支持的格式
WordMergeTool-CLI.exe --list-formats

# 端到端自检（自动造样例并验证）
WordMergeTool-CLI.exe --selftest "D:\tmp\wmtool-test"
```

---

## 支持的格式

| 类别 | 扩展名 | 合并方式 |
| --- | --- | --- |
| 图片 | `.jpg .jpeg .png .bmp .gif .webp .tif .tiff .ico` | **一图一页**，等比缩放，**不做 OCR** |
| PDF | `.pdf` | **一页对一页**，逐页渲染，**不做 OCR** |
| Word | `.docx .docm` | 重放正文：段落格式、表格、图片、横向节 |
| Excel | `.xlsx .xlsm .csv .tsv` | 每个工作表一张表，宽表自动转横版 |
| PowerPoint | `.pptx .pptm .ppsx` | 一页幻灯片对一页 |
| 文本 | `.txt .md .log .json .xml .html .yaml .ini` | 结构化解析（Markdown 走完整语法） |
| 代码 | `.py .js .ts .java .c .cpp .go .sql .ps1 .bat` … | 等宽字体保留缩进 |
| 旧格式 | `.doc .ppt .xls .wps .et .dps` | 需本机有 Office 或 LibreOffice |

---

## 排版规则（这是本工具的核心）

需求里最硬的两条是「**不超出 Word 页面**」和「**不损失源文件内容**」。
本工具用一套统一的几何约束来同时满足它们：

### 1. 图片与 PDF：一页一个，绝不出框

* 按页面**可用内容框**等比缩放，长边贴边、短边留白
* **只缩不放**：小图保持原尺寸，不会糊成大马赛克
* 一图一页、一页 PDF 一页 Word，中间硬分页

### 2. 横向 / 纵向页面自动切换

宽内容会自动落到**横向页**上，窄内容留在纵向页：

| 内容 | 处理 |
| --- | --- |
| Excel 多列宽表 | 纵向放不下 → 自动切横版；仍放不下 → 缩小字号（下限 5.5pt） |
| 16:9 / 4:3 幻灯片 | 一律横向，否则会被压成一条 |
| 横向 PDF 页 | 该页所在节切横向，逐页自适应（横竖混排的 PDF 也能还原） |
| 源 Word 里的横向节 | 按源文档的节方向还原 |

每个源文件处理完都会**自动切回纵向**，不会把横向"传染"给下一个文件。

### 3. 表格：按可用宽度压缩

* 等比缩放列宽到页面可用宽度内，并固定表格布局（`tblLayout=fixed`），
  Word 不再自作主张重排
* 压缩比过大时同步缩小字号，避免文字挤成竖排
* 首行自动设为跨页重复的表头

### 4. 内容不丢

* 超长表格按行拆到多页续排，**不截断任何一列或一行**
* 单元格文本保留全文（不做截断显示）
* 每个源文件都带标题、书签；文档开头有目录域（Word 里按 `F9` 更新页码）

> 关于两条需求的取舍需要说明：**PDF 与幻灯片合并进来后是位图**，其中的文字不能被 Word 搜索或选中。
> 这是「不 OCR」+「一页对一页」两条约束叠加的必然结果。
> 幻灯片额外提供「可编辑文字大纲」段落（hybrid 模式），保证文字仍可复制检索。

---

## 源文件保护

**任何源文件都不会被修改**，这由三道机制保证：

1. **只读打开**：Word/PDF/Excel 一律以只读方式解析，从不调用保存
2. **副本转换**：旧格式（`.doc/.ppt/.xls`）先复制到临时目录，在副本上转换；
   Office COM 打开时传 `ReadOnly=True`，且永不 `Save`
3. **指纹校验**：合并前后对每个源文件算 SHA-256 并比对，结果写进日志和合并报告。
   如果某个文件在合并期间被其它程序改动，会在报告里明确指出来

输出采用「先写临时文件再原子替换」，中途崩溃不会留下半个看似正常的 docx。

---

## 日志与崩溃日志

全部日志都写在 **exe 同级目录的 `logs\`** 下（源码模式则在项目目录下），
所以把 exe 拷到别的电脑，日志跟着走，排障不用远程连过去。

| 文件 | 内容 |
| --- | --- |
| `logs\app-YYYYMMDD.log` | 全量运行日志（含毫秒时间戳、级别、线程名），单文件 5MB 滚动，保留 5 份 |
| `logs\error.log` | 只记 ERROR 及以上，快速扫一眼出过什么错 |
| `logs\crash-YYYYMMDD-HHMMSS.txt` | **崩溃日志**：完整异常栈 + 环境信息 + 最近 400 行日志 |

崩溃日志通过三个钩子捕获，覆盖面是全的：

* `sys.excepthook` —— 主线程未捕获异常
* `threading.excepthook` —— 子线程未捕获异常（转换就是在子线程跑的）
* `tkinter.report_callback_exception` —— 界面事件回调里的异常（这类异常默认只打印到控制台，很容易漏）

程序内还挂了「日志窗口」，界面上的日志和文件日志是同一份内容。

---

## 打包成 exe

```powershell
# 一键打包（生成 dist\WordMergeTool.exe 与 dist\WordMergeTool-CLI.exe）
powershell -ExecutionPolicy Bypass -File .\build.ps1 -Mode onefile
```

打包要点（脚本已处理）：

* 单文件模式，**目标机无需 Python**
* 显式排除 `matplotlib`、`numpy`、`pandas`、`scipy` —— PyInstaller 常把它们从
  PyMuPDF / Pillow 的依赖里误拖进来，体积会白白涨几十 MB
* 隐藏导入 `tkinterdnd2`（拖拽功能，PyInstaller 静态分析看不到）
* 随包带上 `tkinterdnd2` 的原生库（`tkdnd`）

产物自检：把 exe 拷到另一台同架构的 Windows 上双击，能开界面即可。
真正的"可移植性"取决于目标机是否有 **VC++ 运行库**（Windows 10/11 一般自带）。

---

## 项目结构

```
word合并工具/
├── src/wmtool/
│   ├── __init__.py          版本号唯一来源（1.0.0）
│   ├── __main__.py          python -m wmtool 入口
│   ├── cli.py               命令行入口
│   ├── gui.py               图形界面
│   ├── engine.py            合并编排（分页、标题、书签、指纹校验、原子落盘）
│   ├── geometry.py          排版度量（EMU、页框、等比缩放）
│   ├── docx_util.py         docx 底层封装（横向节、书签、目录域、表格边框）
│   ├── converters/          各格式转换器
│   │   ├── base.py          转换器契约
│   │   ├── image.py         图片 → 一图一页
│   │   ├── pdf.py           PDF → 一页对一页
│   │   ├── docx.py          Word 正文重放
│   │   ├── xlsx.py          Excel → 自适应横竖版表格
│   │   ├── pptx.py          PPT → 三种模式
│   │   ├── text.py          文本 / Markdown / 代码
│   │   └── legacy.py        旧格式（COM / LibreOffice）
│   ├── mdparse.py           Markdown 解析（无第三方依赖）
│   ├── verify.py            产物校验（不出页面 + 不丢内容）
│   ├── selftest.py          端到端自检
│   ├── logging_setup.py     日志与崩溃日志
│   ├── settings.py          用户设置持久化
│   └── paths.py             路径与运行模式
├── build.ps1                一键打包
├── requirements.txt
└── CHANGELOG.md
```

---

## 自检与回归

```powershell
python -m wmtool.cli --selftest ".\selftest_out"
```

自检会：**造样例 → 合并 → 校验**，覆盖全部格式，并断言：

* 源文件指纹未变（证明没改源文件）
* 产物能被 Word 解析，段落/表格/图片/节数量达标
* **没有任何表格或图片超出所在节的可用宽高**（按元素实际归属的节逐块判定）
* 关键内容关键词全部命中（证明内容没丢）

---

## 常见问题

**Q：为什么合并后的 PDF 文字选不中？**
A：因为你要求「不用 OCR」且「一页对一页」。要在 Word 里可编辑，就必须识别文字（OCR）或重建排版，
两者都会改变原始呈现。当前策略是保真优先。

**Q：合并很慢是正常的吗？**
A：PPT 渲染是瓶颈。工具按 PowerPoint COM → LibreOffice → 内置渲染器 的顺序找可用通道；
本机装了 PowerPoint 时最快也最保真。内置渲染器是零依赖兜底，复杂特效会简化。

**Q：`.xls` 文件被跳过了？**
A：纯 Python 没有可靠的 `.xls` 读取器。装 Microsoft Office 或 LibreOffice 后即可支持；
或先用 Office 另存为 `.xlsx`。

**Q：日志在哪？**
A：exe 同级的 `logs\` 目录。出错时优先把 `crash-*.txt` 发回来。

---

## 版本

当前 **1.0.0**。版本号唯一事实来源是 `src/wmtool/__init__.py` 里的 `__version__`，
变更记录见 [CHANGELOG.md](CHANGELOG.md)。
