# -*- coding: utf-8 -*-
"""端到端自检：造样例 -> 合并 -> 校验产物。

``python -m wmtool.cli --selftest <目录>`` 会跑这个模块。
自检的价值在于：**不依赖用户手上的文件**，也能证明整条链路是通的，
并且能验证"源文件没被改"和"内容没丢"这两条硬要求。
"""

from __future__ import annotations

import datetime as _dt
import hashlib
import io
import sys
from pathlib import Path

from . import __app_name_cn__, __version__
from .engine import MergeEngine, _sha256
from .logging_setup import get_logger, install_crash_handlers, setup_logging
from .settings import Settings

log = get_logger("selftest")


# --------------------------------------------------------------------------- #
# 造样例
# --------------------------------------------------------------------------- #
def make_samples(root: Path) -> list[Path]:
    """在 ``root/源文件`` 里造出各种格式的样例，返回文件列表（自然序）。"""
    src_dir = root / "源文件"
    src_dir.mkdir(parents=True, exist_ok=True)
    made: list[Path] = []

    made.append(_make_docx(src_dir))
    made.append(_make_xlsx(src_dir))
    made.append(_make_pptx(src_dir))
    made.append(_make_pdf(src_dir))
    made += _make_images(src_dir)
    made.append(_make_markdown(src_dir))
    made.append(_make_text_gbk(src_dir))
    made.append(_make_csv(src_dir))
    return sorted(made, key=lambda p: p.name)


def _make_docx(src_dir: Path) -> Path:
    """样例 Word：含标题、正文、表格、图片、横向节。"""
    from docx import Document as D
    from docx.enum.section import WD_ORIENT
    from docx.shared import Pt

    path = src_dir / "01-样例文档.docx"
    doc = D()
    doc.add_heading("这是一份样例 Word 文档", level=1)
    doc.add_paragraph("用于验证合并工具能正确搬运正文、格式与表格。")
    doc.add_paragraph("第二段：包含**粗体**（源里是真粗体）、以及一个长句子来测试换行与折行行为是否正常。")

    t = doc.add_table(rows=3, cols=3)
    for r, row in enumerate([["项目", "数量", "金额"], ["A 产品", "12", "1,200.00"], ["B 产品", "7", "980.50"]]):
        for c, val in enumerate(row):
            t.cell(r, c).text = val

    # 内嵌一张小图
    from PIL import Image

    buf = io.BytesIO()
    Image.new("RGB", (400, 240), (200, 220, 255)).save(buf, format="PNG")
    buf.seek(0)
    doc.add_picture(buf, width=Pt(200))

    # 追加一个横向节，验证方向还原
    sec = doc.add_section()
    sec.orientation = WD_ORIENT.LANDSCAPE
    from docx.shared import Inches

    sec.page_width, sec.page_height = sec.page_height, sec.page_width
    doc.add_paragraph("这一段位于源文档的【横向页】里，合并后应当仍是横向页。")
    doc.save(str(path))
    return path


def _make_xlsx(src_dir: Path) -> Path:
    """样例 Excel：一个宽表（逼出横向页）+ 一个窄表。"""
    from openpyxl import Workbook

    path = src_dir / "02-样例表格.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = "月度明细"
    headers = [
        "序号", "日期", "部门", "项目名称", "负责人", "数量", "单价", "金额",
        "税率", "含税金额", "备注说明文字列", "状态", "审批人", "归档编号",
    ]
    ws.append(headers)
    base = _dt.datetime(2025, 1, 1)
    for i in range(1, 31):
        ws.append([
            i, base + _dt.timedelta(days=i - 1), f"部门{i % 5 + 1}", f"项目-{i:03d}",
            f"负责人{i % 7 + 1}", i * 3, 12.5 * (i % 9 + 1), i * 3 * 12.5,
            0.06, i * 3 * 12.5 * 1.06, "这是一段较长的备注文字，用来测试列宽估算与换行",
            "已完成" if i % 2 else "进行中", f"审批人{i % 4 + 1}", f"ARC-{1000 + i}",
        ])
    ws2 = wb.create_sheet("汇总")
    ws2.append(["部门", "合计金额"])
    for d in range(1, 6):
        ws2.append([f"部门{d}", d * 1234.56])
    wb.save(str(path))
    return path


def _make_pptx(src_dir: Path) -> Path:
    """样例 PPT：含标题页 + 内容页（文本框/形状/表格）。"""
    from pptx import Presentation
    from pptx.util import Inches, Pt

    path = src_dir / "03-样例演示.pptx"
    prs = Presentation()
    blank = prs.slide_layouts[6]

    s1 = prs.slides.add_slide(blank)
    tb = s1.shapes.add_textbox(Inches(0.8), Inches(2.2), Inches(9), Inches(1.5))
    p = tb.text_frame.paragraphs[0]
    run = p.add_run()
    run.text = "样例演示文稿"
    run.font.size = Pt(40)

    s2 = prs.slides.add_slide(blank)
    tb2 = s2.shapes.add_textbox(Inches(0.6), Inches(0.5), Inches(9), Inches(2.5))
    tf = tb2.text_frame
    tf.text = "要点一：演示如何合并幻灯片"
    for line in ("要点二：横向版式要保住", "要点三：文字仍可复制"):
        para = tf.add_paragraph()
        para.text = line
    rows, cols = 3, 2
    shape = s2.shapes.add_table(rows, cols, Inches(0.6), Inches(3.4), Inches(8), Inches(1.8))
    table = shape.table
    for r, row in enumerate([["指标", "数值"], ["覆盖率", "98%"], ["缺陷数", "3"]]):
        for c, val in enumerate(row):
            table.cell(r, c).text = val

    prs.save(str(path))
    return path


def _make_pdf(src_dir: Path) -> Path:
    """样例 PDF：一页纵向 + 一页横向（验证方向自适应）。"""
    import pymupdf

    path = src_dir / "04-样例文档.pdf"
    doc = pymupdf.open()
    p1 = doc.new_page(width=595, height=842)  # A4 纵向
    p1.insert_text((60, 90), "Sample PDF - Portrait Page", fontsize=20)
    p1.insert_text((60, 130), "This page tests one-page-to-one-page merging.", fontsize=11)
    p1.draw_rect(pymupdf.Rect(60, 170, 530, 300), color=(0.2, 0.4, 0.8), width=1.5)

    p2 = doc.new_page(width=842, height=595)  # A4 横向
    p2.insert_text((60, 90), "Sample PDF - Landscape Page", fontsize=20)
    p2.draw_line(pymupdf.Point(60, 140), pymupdf.Point(780, 140))
    doc.save(str(path))
    doc.close()
    return path


def _make_images(src_dir: Path) -> list[Path]:
    """样例图片：一张竖图、一张超宽图、一张带透明通道的 PNG。"""
    from PIL import Image, ImageDraw

    out = []

    portrait = src_dir / "05-样例照片.jpg"
    im = Image.new("RGB", (900, 1600), (245, 240, 230))
    d = ImageDraw.Draw(im)
    d.rectangle([60, 60, 840, 1540], outline=(120, 90, 60), width=6)
    d.text((90, 90), "PORTRAIT SAMPLE 900x1600", fill=(60, 40, 20))
    im.save(portrait, quality=90)
    out.append(portrait)

    wide = src_dir / "06-样例宽图.png"
    im2 = Image.new("RGB", (3000, 900), (235, 245, 250))
    d2 = ImageDraw.Draw(im2)
    d2.rectangle([40, 40, 2960, 860], outline=(30, 90, 140), width=8)
    d2.text((80, 80), "WIDE SAMPLE 3000x900 (should be scaled to fit one page)", fill=(20, 60, 100))
    im2.save(wide)
    out.append(wide)

    alpha = src_dir / "07-样例透明图.png"
    im3 = Image.new("RGBA", (600, 600), (0, 0, 0, 0))
    d3 = ImageDraw.Draw(im3)
    d3.ellipse([50, 50, 550, 550], fill=(80, 170, 220, 255))
    d3.text((180, 290), "RGBA SAMPLE", fill=(255, 255, 255, 255))
    im3.save(alpha)
    out.append(alpha)

    return out


def _make_markdown(src_dir: Path) -> Path:
    """样例 Markdown：标题/列表/表格/代码块/引用全都有。"""
    path = src_dir / "08-样例笔记.md"
    path.write_text(
        """# 样例笔记

这是一段普通段落，包含 **粗体**、*斜体* 和 `行内代码`。

## 二级标题

- 列表项一
- 列表项二
  - 嵌套项

1. 有序一
2. 有序二

> 这是一段引用文字。

| 名称 | 说明 | 数量 |
| --- | --- | --- |
| 甲 | 第一种情况 | 10 |
| 乙 | 第二种情况 | 20 |

```python
def hello(name):
    return f"你好，{name}"
```
""",
        encoding="utf-8",
    )
    return path


def _make_text_gbk(src_dir: Path) -> Path:
    """样例 GBK 文本（验证编码自动识别）。"""
    path = src_dir / "09-样例GBK文本.txt"
    path.write_bytes("这是 GBK 编码的纯文本文件。\n第二行：用于验证编码自动识别。\n".encode("gbk"))
    return path


def _make_csv(src_dir: Path) -> Path:
    """样例 CSV。"""
    path = src_dir / "10-样例数据.csv"
    path.write_text(
        "姓名,部门,金额\n张三,研发部,1200.50\n李四,市场部,980.00\n王五,财务部,1560.75\n",
        encoding="utf-8-sig",
    )
    return path


# --------------------------------------------------------------------------- #
# 跑自检
# --------------------------------------------------------------------------- #
def run_selftest(root: Path) -> int:
    """执行完整自检并返回退出码（0 = 全通过）。"""
    setup_logging()
    install_crash_handlers()

    root = Path(root).resolve()
    root.mkdir(parents=True, exist_ok=True)
    print(f"{__app_name_cn__} v{__version__} 端到端自检")
    print(f"工作目录: {root}")
    print("-" * 66)

    # 1. 造样例
    print("[1/4] 生成各格式样例 ...")
    files = make_samples(root)
    for f in files:
        print(f"      {f.name:<28} {f.stat().st_size / 1024:>8.1f} KB")
    print(f"      共 {len(files)} 个样例文件")

    # 2. 记录源文件指纹
    print("[2/4] 记录源文件指纹 ...")
    before = {f: _sha256(f) for f in files}

    # 3. 合并
    print("[3/4] 执行合并 ...")
    settings = Settings()
    settings.timestamp_output = False
    engine = MergeEngine(settings)
    output = root / "合并结果.docx"
    report = engine.merge(
        files,
        output,
        progress=lambda i, t, m: print(f"\r      [{i}/{t}] {m[:52]:<52}", end="", flush=True),
    )
    print("\r" + " " * 70 + "\r", end="")

    print(f"      结果: {report.summary_line()}")
    print(f"      输出: {report.output_path}（{report.output_bytes / 1024:.1f} KB）")

    # 4. 校验
    print("[4/4] 校验 ...")
    failures: list[str] = []

    # 4.1 源文件未变
    changed = [f.name for f in files if _sha256(f) != before[f]]
    if changed:
        failures.append(f"源文件被修改了: {changed}")
    else:
        print(f"      [通过] 源文件指纹未变（{len(files)} 个文件）")

    # 4.2 输出存在且是合法 docx
    if not output.is_file() or output.stat().st_size == 0:
        failures.append("输出文件不存在或为空")
    else:
        try:
            from docx import Document as D

            merged = D(str(output))
            n_para = len(merged.paragraphs)
            n_tbl = len(merged.tables)
            n_pics = len(merged.inline_shapes)
            n_sect = len(merged.sections)
            print(
                f"      [通过] 输出可被 Word 解析：{n_para} 段 / {n_tbl} 表 / "
                f"{n_pics} 图 / {n_sect} 节"
            )
            if n_para < 20:
                failures.append(f"段落数偏少（{n_para}），内容可能丢失")
            if n_tbl < 3:
                failures.append(f"表格数偏少（{n_tbl}），Excel/Word/PPT 表格可能丢失")
            if n_pics < 6:
                failures.append(f"图片数偏少（{n_pics}），图片/PDF/PPT 渲染可能丢失")
            if n_sect < 2:
                failures.append("只有一个节，横向页切换可能失效")
            else:
                landscape = sum(1 for s in merged.sections if s.page_width > s.page_height)
                print(f"      [通过] 横向节 {landscape} 个（Excel/PPT/横向 PDF 用）")
                if landscape < 1:
                    failures.append("没有横向节，宽表/幻灯片/横向 PDF 可能被压扁")
        except Exception as exc:
            failures.append(f"输出无法被 python-docx 解析: {exc}")

    # 4.3 失败文件
    if report.failed:
        for o in report.failed:
            failures.append(f"{o.name} 处理失败: {o.error}")

    # 4.4 关键提示词是否出现在正文里（内容没丢的粗验证）
    if output.is_file():
        try:
            from .verify import verify_document

            vr = verify_document(
                output,
                must_contain=[
                    "样例文档", "月度明细", "要点一", "样例笔记", "GBK 编码", "张三",
                ],
                min_images=6,
                min_tables=3,
            )
            print(vr.report())
            if not vr.ok:
                for i in vr.issues:
                    if i.severity == "error":
                        failures.append(f"{i.kind}: {i.detail}")
        except Exception as exc:
            failures.append(f"产物校验失败: {exc}")

    print("-" * 66)
    if failures:
        print("自检【失败】：")
        for f in failures:
            print(f"  ✗ {f}")
        return 1

    print("自检【全部通过】")
    print(f"请打开查看: {output}")
    return 0


if __name__ == "__main__":  # pragma: no cover
    target = Path(sys.argv[1]) if len(sys.argv) > 1 else Path.cwd() / "selftest"
    sys.exit(run_selftest(target))
