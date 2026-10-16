# -*- coding: utf-8 -*-
"""纯文本 / Markdown / 代码 -> Word。

* ``.md/.markdown``：走 :mod:`wmtool.mdparse` 的结构化解析（标题/列表/表格/代码块/引用）
* ``.txt/.log/.json/.xml/.html`` 等：按行写成段落，自动识别编码
* 代码文件：等宽字体 + 保留缩进

编码按 utf-8-sig -> utf-8 -> gbk/gb18030 -> latin-1 的顺序尝试，
四轮都失败才报错——国内环境 GBK 文本很常见。
"""

from __future__ import annotations

from pathlib import Path

from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Emu, Pt, RGBColor

from ..docx_util import (
    add_caption,
    add_heading,
    add_hyperlink,
    _set_run_east_asian,
    set_cell_text,
    set_table_layout_fixed,
    style_data_table,
)
from ..logging_setup import get_logger
from ..mdparse import parse_inline, parse_markdown
from .base import ConvertContext, ConvertResult

log = get_logger("conv.text")

SUPPORTED_SUFFIXES = {
    ".txt", ".text", ".log", ".md", ".markdown", ".csv.txt",
    ".json", ".xml", ".html", ".htm", ".yaml", ".yml", ".ini", ".cfg", ".conf",
    ".py", ".js", ".ts", ".java", ".c", ".cpp", ".h", ".cs", ".go", ".rs",
    ".sql", ".sh", ".ps1", ".bat", ".cmd", ".css", ".vue", ".php", ".rb",
}

MARKDOWN_SUFFIXES = {".md", ".markdown"}
CODE_SUFFIXES = {
    ".py", ".js", ".ts", ".java", ".c", ".cpp", ".h", ".cs", ".go", ".rs",
    ".sql", ".sh", ".ps1", ".bat", ".cmd", ".css", ".vue", ".php", ".rb",
    ".json", ".xml", ".yaml", ".yml", ".ini", ".cfg", ".conf",
}
HTML_SUFFIXES = {".html", ".htm"}

#: 编码尝试顺序
ENCODINGS = ("utf-8-sig", "utf-8", "gbk", "gb18030", "big5", "latin-1")

MONO_FONT = "Consolas"


class TextConverter:
    """文本类文件 -> Word。"""

    name = "text"
    display = "文本"

    def convert(self, ctx: ConvertContext) -> ConvertResult:
        result = ConvertResult(converter=self.name)
        src = ctx.src
        suffix = src.suffix.lower()

        text, encoding = self._read_text(src)
        if encoding not in ("utf-8-sig", "utf-8"):
            result.warn(f"文件按 {encoding} 解码（非 UTF-8）")

        if suffix in MARKDOWN_SUFFIXES:
            self._write_markdown(ctx, text)
            result.summary = f"Markdown 文档（{len(text.splitlines())} 行）"
        elif suffix in HTML_SUFFIXES:
            plain = self._strip_html(text)
            self._write_plain(ctx, plain, mono=False)
            result.summary = f"HTML 文档（已剥离标签，{len(plain.splitlines())} 行）"
        elif suffix in CODE_SUFFIXES:
            self._write_plain(ctx, text, mono=True)
            result.summary = f"代码/结构化文本（{len(text.splitlines())} 行）"
        else:
            self._write_plain(ctx, text, mono=False)
            result.summary = f"文本（{len(text.splitlines())} 行）"

        if not text.strip():
            result.warn("文件内容为空")
        ctx.report(1.0, "文本已写入")
        return result

    # ------------------------------------------------------------------ #
    @staticmethod
    def _read_text(src: Path) -> tuple[str, str]:
        """按候选编码读文本。"""
        raw = src.read_bytes()
        if not raw:
            return "", "utf-8"
        for enc in ENCODINGS:
            try:
                return raw.decode(enc), enc
            except UnicodeDecodeError:
                continue
        # 兜底：忽略错误强解，绝不因为编码问题丢整个文件
        return raw.decode("utf-8", errors="replace"), "utf-8(replace)"

    @staticmethod
    def _strip_html(text: str) -> str:
        """极简 HTML 去标签（不引 bs4，够用即可）。"""
        import html as _html
        import re

        text = re.sub(r"(?is)<(script|style)[^>]*>.*?</\1>", "", text)
        text = re.sub(r"(?i)<br\s*/?>", "\n", text)
        text = re.sub(r"(?i)</(p|div|tr|li|h[1-6])>", "\n", text)
        text = re.sub(r"<[^>]+>", "", text)
        text = _html.unescape(text)
        return re.sub(r"\n{3,}", "\n\n", text)

    # ------------------------------------------------------------------ #
    def _write_plain(self, ctx: ConvertContext, text: str, *, mono: bool) -> None:
        """按行写段落（保留空行结构）。"""
        lines = text.replace("\r\n", "\n").replace("\r", "\n").split("\n")
        for line in lines:
            p = ctx.doc.add_paragraph()
            p.paragraph_format.space_after = Pt(0)
            if not line.strip():
                continue
            run = p.add_run(line)
            if mono:
                run.font.name = MONO_FONT
                run.font.size = Pt(8.5)
                _set_run_east_asian(run, MONO_FONT)
            else:
                run.font.size = Pt(10.5)

    def _write_markdown(self, ctx: ConvertContext, text: str) -> None:
        """把 Markdown 渲染成 Word 结构。"""
        blocks = parse_markdown(text)
        usable = ctx.box.usable_width

        for block in blocks:
            if block.kind == "heading":
                # 源文档的一级标题在合并文档里降一级，避免抢封面标题的层级
                add_heading(
                    ctx.doc,
                    self._inline_text(block.text),
                    level=min(4, block.level + 1),
                    size_pt=max(11.0, 17.0 - block.level * 1.6),
                )
            elif block.kind == "paragraph":
                p = ctx.doc.add_paragraph()
                self._write_inline(p, block.text)
            elif block.kind == "quote":
                p = ctx.doc.add_paragraph()
                p.paragraph_format.left_indent = Emu(int(0.6 * 360000))
                run = p.add_run("▎ " + self._inline_text(block.text))
                run.italic = True
                run.font.color.rgb = RGBColor(0x55, 0x55, 0x55)
            elif block.kind == "rule":
                p = ctx.doc.add_paragraph()
                run = p.add_run("─" * 40)
                run.font.color.rgb = RGBColor(0xBB, 0xBB, 0xBB)
            elif block.kind == "code":
                self._write_code(ctx, block.text)
            elif block.kind == "list":
                self._write_list(ctx, block)
            elif block.kind == "table":
                self._write_md_table(ctx, block.rows, usable)

    def _write_inline(self, p, text: str) -> None:  # type: ignore[no-untyped-def]
        """写一行带行内格式的文本。"""
        for piece in parse_inline(text):
            if piece.link:
                add_hyperlink(p, piece.link, piece.text)
                continue
            run = p.add_run(piece.text)
            run.bold = piece.bold
            run.italic = piece.italic
            if piece.code:
                run.font.name = MONO_FONT
                run.font.size = Pt(9.5)
                run.font.color.rgb = RGBColor(0xB0, 0x30, 0x60)
                _set_run_east_asian(run, MONO_FONT)

    @staticmethod
    def _inline_text(text: str) -> str:
        """去掉行内标记，取纯文字。"""
        return "".join(p.text for p in parse_inline(text))

    def _write_code(self, ctx: ConvertContext, code: str) -> None:
        """代码块：等宽 + 灰底（用单元格做底色的方式太绕，这里用缩进+灰字）。"""
        for line in code.split("\n"):
            p = ctx.doc.add_paragraph()
            p.paragraph_format.space_after = Pt(0)
            p.paragraph_format.left_indent = Emu(int(0.5 * 360000))
            run = p.add_run(line if line else " ")
            run.font.name = MONO_FONT
            run.font.size = Pt(8.5)
            run.font.color.rgb = RGBColor(0x33, 0x33, 0x33)
            _set_run_east_asian(run, MONO_FONT)
        ctx.doc.add_paragraph()

    def _write_list(self, ctx: ConvertContext, block) -> None:  # type: ignore[no-untyped-def]
        """列表：手动加圆点/序号 + 缩进（不用 Word 的 numbering，避免编号串到别的文档）。"""
        for i, (indent, item) in enumerate(block.items, start=1):
            p = ctx.doc.add_paragraph()
            p.paragraph_format.left_indent = Emu(int((0.45 + 0.45 * indent) * 360000))
            p.paragraph_format.space_after = Pt(2)
            bullet = f"{i}. " if block.ordered else ("○ " if indent else "• ")
            run = p.add_run(bullet)
            run.font.size = Pt(10.5)
            self._write_inline(p, item)

    def _write_md_table(self, ctx: ConvertContext, rows: list[list[str]], usable: int) -> None:
        """Markdown 表格 -> Word 表格（等宽列 + 首行表头）。"""
        if not rows:
            return
        n_cols = max(len(r) for r in rows)
        each = max(1, usable // n_cols)
        font_pt = 9.0 if n_cols <= 5 else 7.0

        table = ctx.doc.add_table(rows=len(rows), cols=n_cols)
        style_data_table(table, font_pt=font_pt)
        set_table_layout_fixed(table)
        for r_idx, row in enumerate(rows):
            for c_idx in range(n_cols):
                text = row[c_idx] if c_idx < len(row) else ""
                set_cell_text(
                    table.cell(r_idx, c_idx),
                    self._inline_text(text),
                    bold=(r_idx == 0),
                    size_pt=font_pt,
                    align=WD_ALIGN_PARAGRAPH.LEFT,
                )
        for c_idx in range(n_cols):
            for row in table.rows:
                try:
                    row.cells[c_idx].width = Emu(each)
                except Exception:
                    pass
        ctx.doc.add_paragraph()
