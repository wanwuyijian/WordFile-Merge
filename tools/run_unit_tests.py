# -*- coding: utf-8 -*-
"""单元测试：几何、方向决策、Markdown 解析、日志与崩溃日志。

跑法::

    python tools/run_unit_tests.py

刻意不引 pytest —— 这个项目要能在一台只有 Python 的干净机器上自证正确，
多一个测试框架就多一个依赖。
"""

from __future__ import annotations

import sys
import tempfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

PASSED = 0
FAILED: list[str] = []


def check(cond: bool, label: str, detail: str = "") -> None:
    """断言并计数。"""
    global PASSED
    if cond:
        PASSED += 1
        print(f"  [OK]   {label}" + (f"  ({detail})" if detail else ""))
    else:
        FAILED.append(label)
        print(f"  [FAIL] {label}" + (f"  ({detail})" if detail else ""))


def test_geometry() -> None:
    """几何：等比缩放与"只缩不放"。"""
    print("\n[geometry] 等比缩放")
    from wmtool.geometry import fit_size

    # 宽图：按宽度缩，高度按比例
    w, h = fit_size(2000, 1000, 1000, 1000)
    check(w == 1000 and h == 500, "宽图按宽度贴合", f"{w}x{h}")

    # 高图：按高度缩
    w, h = fit_size(1000, 3000, 1000, 1000)
    check(w == 333 and h == 1000, "高图按高度贴合", f"{w}x{h}")

    # 小图：不放大
    w, h = fit_size(100, 50, 1000, 1000)
    check(w == 100 and h == 50, "小图不放大（避免糊）", f"{w}x{h}")

    # 非法尺寸不能返回 0（python-docx 收到 0 会抛异常）
    w, h = fit_size(0, 100, 1000, 1000)
    check(w >= 1 and h >= 1, "非法尺寸不返回 0", f"{w}x{h}")
    w, h = fit_size(-5, -5, 1000, 1000)
    check(w >= 1 and h >= 1, "负数尺寸不返回 0", f"{w}x{h}")

    # 极端长宽比
    w, h = fit_size(20000, 300, 1000, 1000)
    check(w == 1000 and h == 15, "极端横幅正确贴宽", f"{w}x{h}")


def test_orientation() -> None:
    """Excel 方向决策：小表纵向、宽表横向、超宽表必横向。"""
    print("\n[xlsx] 横竖方向决策")
    from wmtool.converters.base import ConvertContext
    from wmtool.converters.xlsx import FONT_MAX_PT, FONT_MIN_PT, XlsxConverter
    from wmtool.geometry import LANDSCAPE_BOX, PORTRAIT_BOX
    from wmtool.settings import Settings

    ctx = ConvertContext(
        doc=None, settings=Settings(), src=None, index=1, total=1,
        box=PORTRAIT_BOX, landscape_box=LANDSCAPE_BOX,
    )

    def decide(ncols: int):  # type: ignore[no-untyped-def]
        rows = [[f"列{i:02d}" for i in range(ncols)]]
        natural = sum(XlsxConverter._estimate_widths(rows))
        return XlsxConverter._choose_orientation(ctx, natural)

    box, font = decide(3)
    check(not box.landscape, "3 列表保持纵向")
    check(font == FONT_MAX_PT, "3 列表不缩字号", f"{font}pt")

    box, font = decide(8)
    check(not box.landscape, "8 列表保持纵向")

    box, font = decide(14)
    check(box.landscape, "14 列表转横向")
    check(font > FONT_MIN_PT, "14 列表在横向下字号仍高于下限", f"{font:.1f}pt")

    # 关键回归：超宽表在横竖两边都触底时，必须选横向
    box, font = decide(40)
    check(box.landscape, "40 列表转横向（两边都触底时选更宽的）", f"{font:.1f}pt")

    box_l, font_l = decide(40)
    ctx2 = ctx
    natural40 = sum(XlsxConverter._estimate_widths([[f"列{i:02d}" for i in range(40)]]))
    p_font = max(FONT_MIN_PT, FONT_MAX_PT * PORTRAIT_BOX.usable_width / natural40)
    l_font = max(FONT_MIN_PT, FONT_MAX_PT * LANDSCAPE_BOX.usable_width / natural40)
    check(
        l_font >= p_font,
        "横向给出的字号不低于纵向（选横向是严格的更优解）",
        f"横{l_font:.1f} vs 纵{p_font:.1f}",
    )


def test_mdparse() -> None:
    """Markdown 解析：各类块的识别。"""
    print("\n[mdparse] Markdown 解析")
    from wmtool.mdparse import parse_inline, parse_markdown

    md = """# 标题一

正文段落。

## 标题二

- 项目 A
- 项目 B

1. 第一
2. 第二

> 引用内容

```python
print("hi")
```

| 列A | 列B |
| --- | --- |
| 1 | 2 |

---
"""
    blocks = parse_markdown(md)
    kinds = [b.kind for b in blocks]
    check(kinds.count("heading") == 2, "识别到 2 个标题", str(kinds))
    check("code" in kinds, "识别到代码块")
    check("quote" in kinds, "识别到引用")
    check("table" in kinds, "识别到表格")
    check("rule" in kinds, "识别到分隔线")
    check(kinds.count("list") == 2, "识别到 2 个列表")

    tbl = next(b for b in blocks if b.kind == "table")
    # 分隔行（| --- | --- |）是语法不是数据，所以 1 表头 + 1 数据行 = 2 行
    check(len(tbl.rows) == 2, "表格 2 行（表头 + 1 数据行，分隔行不计入）", str(tbl.rows))
    check(tbl.rows[0] == ["列A", "列B"], "表头解析正确", str(tbl.rows[0]))

    code = next(b for b in blocks if b.kind == "code")
    check("print" in code.text and code.language == "python", "代码块内容与语言正确")

    # 行内格式
    inline = parse_inline("普通 **粗** 和 *斜* 与 `代码` 以及 [链接](http://x.com)")
    check(any(p.bold for p in inline), "行内粗体")
    check(any(p.italic for p in inline), "行内斜体")
    check(any(p.code for p in inline), "行内代码")
    check(any(p.link for p in inline), "行内链接")

    # 空输入不能炸
    check(parse_markdown("") == [], "空输入返回空列表")


def test_logging_and_crash() -> None:
    """日志与崩溃日志：必须真的落盘，且崩溃报告含栈与环境信息。"""
    print("\n[logging] 日志与崩溃日志")
    import os

    tmp = Path(tempfile.mkdtemp(prefix="wmtool-logtest-"))
    os.environ["WORDMERGE_LOG_DIR"] = str(tmp)
    try:
        # 重新导入以套用新的日志目录
        for mod in [m for m in list(sys.modules) if m.startswith("wmtool")]:
            del sys.modules[mod]

        from wmtool.logging_setup import (
            get_logger,
            setup_logging,
            write_crash_report,
        )

        log_file = setup_logging()
        check(log_file.is_file(), "主日志文件已创建", log_file.name)

        get_logger("test").info("这是一条测试日志 %d", 42)
        check("这是一条测试日志 42" in log_file.read_text(encoding="utf-8"),
              "日志内容已写入文件")

        # 崩溃报告
        try:
            raise ValueError("故意抛出的测试异常")
        except ValueError:
            exc_type, exc_val, exc_tb = sys.exc_info()
            crash = write_crash_report(exc_type, exc_val, exc_tb, where="单元测试")

        check(crash is not None and crash.is_file(), "崩溃报告已创建")
        content = crash.read_text(encoding="utf-8") if crash else ""
        check("故意抛出的测试异常" in content, "崩溃报告含异常信息")
        check("Traceback" in content, "崩溃报告含调用栈")
        check("Python:" in content and "操作系统:" in content, "崩溃报告含环境信息")
        check("最近日志" in content, "崩溃报告含最近日志片段")
        check("崩溃报告" in content, "崩溃报告有明确标题")
    finally:
        os.environ.pop("WORDMERGE_LOG_DIR", None)
        import shutil

        shutil.rmtree(tmp, ignore_errors=True)


def test_converters_registry() -> None:
    """转换器注册表：常见扩展名都要有归属。"""
    print("\n[converters] 格式注册")
    from wmtool.converters import get_converter, is_supported, supported_suffixes

    for ext in (".docx", ".pptx", ".xlsx", ".pdf", ".jpg", ".png", ".jpeg",
                ".txt", ".md", ".csv", ".doc", ".ppt", ".xls"):
        check(is_supported(Path("x" + ext)), f"{ext} 已注册")

    for ext in (".xyz", ".exe", ".dll"):
        check(not is_supported(Path("x" + ext)), f"{ext} 不被支持（应拒绝）")

    conv = get_converter(Path("a.jpg"))
    check(conv is not None and conv.name == "image", "jpg 走图片转换器")
    conv = get_converter(Path("a.pdf"))
    check(conv is not None and conv.name == "pdf", "pdf 走 PDF 转换器")

    check(len(supported_suffixes()) > 20, "支持的扩展名数量充足",
          f"{len(supported_suffixes())} 种")


def test_docx_util() -> None:
    """docx 底层工具：横向节、书签、目录域、表格边框。"""
    print("\n[docx_util] Word 底层封装")
    from docx import Document as D

    from wmtool.docx_util import (
        add_bookmark,
        add_toc_field,
        new_section,
        page_break,
        set_document_defaults,
        set_table_layout_fixed,
        style_data_table,
    )
    from wmtool.geometry import LANDSCAPE_BOX, PORTRAIT_BOX

    doc = D()
    set_document_defaults(doc, PORTRAIT_BOX)
    check(len(doc.sections) == 1, "初始只有 1 个节")

    page_break(doc)
    new_section(doc, LANDSCAPE_BOX)
    check(len(doc.sections) == 2, "新增节后共 2 个节")
    s = doc.sections[-1]
    check(s.page_width > s.page_height, "新节确实是横向",
          f"{s.page_width}x{s.page_height}")

    add_bookmark(doc, "测试 书签 name!")
    add_toc_field(doc)
    xml = doc.element.body.xml
    check("w:bookmarkStart" in xml, "书签已写入")
    check("TOC" in xml, "目录域已写入")
    check("updateFields" in doc.settings.element.xml, "已设置打开时更新域")

    t = doc.add_table(rows=2, cols=2)
    style_data_table(t)
    set_table_layout_fixed(t)
    check("tblBorders" in t._tbl.xml, "表格边框已写入")
    check("tblLayout" in t._tbl.xml, "固定布局已写入")

    # 存盘再读回，确认产物合法
    tmp = Path(tempfile.mkdtemp(prefix="wmtool-docxtest-"))
    try:
        out = tmp / "t.docx"
        doc.save(str(out))
        reread = D(str(out))
        check(len(reread.sections) == 2, "存盘读回后节数保持")
        check(len(reread.tables) == 1, "存盘读回后表格保持")
    finally:
        import shutil

        shutil.rmtree(tmp, ignore_errors=True)


def main() -> int:
    """跑全部单元测试。"""
    print("=" * 62)
    print(" Word 合并工具 —— 单元测试")
    print("=" * 62)

    for fn in (
        test_geometry,
        test_orientation,
        test_mdparse,
        test_docx_util,
        test_converters_registry,
        test_logging_and_crash,
    ):
        try:
            fn()
        except Exception as exc:
            import traceback

            traceback.print_exc()
            FAILED.append(f"{fn.__name__} 抛异常: {exc}")

    print("\n" + "=" * 62)
    print(f" 通过 {PASSED} 项，失败 {len(FAILED)} 项")
    if FAILED:
        for f in FAILED:
            print(f"  ✗ {f}")
        return 1
    print(" 单元测试【全部通过】")
    return 0


if __name__ == "__main__":
    sys.exit(main())
