# -*- coding: utf-8 -*-
"""Word (.docx) 合并器：把源文档的**正文流**重放进目标文档。

不直接 zip 拼接（那样两个文档的 section/样式/编号会互相打架），
而是走 python-docx 在**同一个文档对象**里重放块级元素：

* 段落：连同 run 级别的粗体/斜体/下划线/字号/颜色/中文字体一起搬
* 表格：按目标页面宽度等比缩放列宽（源表更宽就压缩 + 降字号），保证不出页
* 图片：从源文档的关系里取出二进制，按目标可用框等比缩放后重新嵌入
* 分节：源文档的横向节会用目标文档的横向节还原（关键，见需求第 4 条）

**不动源文件**：python-docx 打开 .docx 只读 lxml 树，全程不 save 源路径。
"""

from __future__ import annotations

import io
from pathlib import Path
from typing import Iterable

from docx import Document as DocxOpen
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.shared import Emu, Pt
from docx.table import Table
from docx.text.paragraph import Paragraph

from ..docx_util import (
    add_caption,
    add_picture_centered,
    set_cell_text,
    set_row_height_pt,
    set_table_layout_fixed,
    repeat_header_row,
    style_data_table,
)
from ..geometry import fit_size
from ..logging_setup import get_logger
from .base import ConvertContext, ConvertResult

log = get_logger("conv.docx")

SUPPORTED_SUFFIXES = {".docx", ".docm"}

#: 最少保留的字号，再小就没法读了；到这一步说明表格确实太宽
MIN_FONT_PT = 6.0
DEFAULT_FONT_PT = 10.5


class DocxConverter:
    """Word 文档合并器（重放式，非 zip 拼接）。"""

    name = "docx"
    display = "Word 文档"

    def convert(self, ctx: ConvertContext) -> ConvertResult:
        result = ConvertResult(converter=self.name)
        src = ctx.src

        try:
            # 只读打开：不改、不存源文件
            source = DocxOpen(str(src))
        except Exception as exc:
            raise ValueError(f"Word 文档打开失败: {exc}") from exc

        body = source.element.body
        blocks = [child for child in body.iterchildren() if isinstance(child.tag, str)]
        # 记录源文档里非默认方向的分节（页码/方向信息在 sectPr 里）
        src_landscape_sections = self._landscape_section_indexes(source)

        para_count = 0
        table_count = 0
        image_count = 0
        heading_count = 0
        index = 0

        for child in blocks:
            tag = child.tag
            if tag == qn("w:p"):
                try:
                    para = Paragraph(child, source)
                except Exception:
                    continue
                # 段前分页、横向节的处理交给"节标记"循环
                if self._is_empty(para):
                    continue
                is_heading = self._is_heading(para)
                n_img, n_tbl = self._copy_paragraph(para, ctx, result)
                image_count += n_img
                table_count += n_tbl
                para_count += 1
                if is_heading:
                    heading_count += 1
            elif tag == qn("w:tbl"):
                try:
                    table = Table(child, source)
                except Exception:
                    continue
                if self._copy_table(table, ctx, result):
                    table_count += 1
            elif tag == qn("w:sectPr"):
                # 源文档的分节符：遇到横向节就切到横向（还原原始排版意图）
                landscape = self._sectpr_is_landscape(child)
                if landscape:
                    from ..docx_util import new_section

                    ctx.doc.add_page_break()
                    new_section(ctx.doc, ctx.landscape_box)
                    result.landscape = True
            index += 1
            ctx.report(min(0.95, (index + 1) / max(1, len(blocks))), "重放 Word 内容")

        result.image_count = image_count
        result.table_count = table_count
        parts = [f"{para_count} 段"]
        if table_count:
            parts.append(f"{table_count} 表")
        if image_count:
            parts.append(f"{image_count} 图")
        result.summary = "Word 正文：" + " · ".join(parts)
        if src_landscape_sections:
            result.warn("源文档含横向页面，已尽量还原为横向节")
        ctx.report(1.0, "Word 内容已重放")
        return result

    # ------------------------------------------------------------------ #
    # 段落
    # ------------------------------------------------------------------ #
    def _copy_paragraph(
        self, para: Paragraph, ctx: ConvertContext, result: ConvertResult
    ) -> tuple[int, int]:
        """复制一个段落（含图片 run）。返回 (图片数, 0)。"""
        # 纯图片段落：单独走图片路径，保证等比缩放
        images = [r for r in para.runs if self._run_has_image(r)]
        if images and not self._paragraph_text(para).strip():
            n = 0
            for run in images:
                if self._copy_image_run(run, ctx):
                    n += 1
            return n, 0

        target = ctx.doc.add_paragraph()
        self._copy_paragraph_format(para, target)

        # 段落里既有文字又有图片：文字照搬，图片按行内图缩放插入
        for run in para.runs:
            if self._run_has_image(run):
                if self._copy_image_run(run, ctx, paragraph=target):
                    result.image_count += 1
                continue
            new_run = target.add_run(run.text)
            self._copy_run_format(run, new_run)
        return 0, 0

    @staticmethod
    def _paragraph_text(para: Paragraph) -> str:
        return "".join(r.text for r in para.runs)

    @staticmethod
    def _is_empty(para: Paragraph) -> bool:
        """空段落（既没字也没图）——丢掉，避免合并后一堆空行。"""
        if para.text.strip():
            return False
        for run in para.runs:
            if run._element.findall(".//" + qn("w:drawing")) or run._element.findall(
                ".//" + qn("w:pict")
            ):
                return False
        return True

    @staticmethod
    def _is_heading(para: Paragraph) -> bool:
        try:
            name = (para.style.name or "").lower()
        except Exception:
            return False
        return "heading" in name or "标题" in name

    def _copy_paragraph_format(self, src: Paragraph, dst: Paragraph) -> None:
        """复制段落级格式：对齐、缩进、段间距、大纲级别。"""
        try:
            spf = src.paragraph_format
            dpf = dst.paragraph_format
            if spf.alignment is not None:
                dpf.alignment = spf.alignment
            for attr in ("left_indent", "right_indent", "first_line_indent"):
                val = getattr(spf, attr, None)
                if val is not None:
                    try:
                        setattr(dpf, attr, val)
                    except Exception:
                        pass
            for attr in ("space_before", "space_after", "line_spacing"):
                val = getattr(spf, attr, None)
                if val is not None:
                    try:
                        setattr(dpf, attr, val)
                    except Exception:
                        pass
            if spf.page_break_before:
                dpf.page_break_before = True
        except Exception as exc:
            log.debug("段落格式复制失败: %s", exc)

        # 保留大纲级别（没有它，插入 TOC 时标题进不了目录）
        try:
            src_ppr = src._p.pPr
            if src_ppr is not None:
                outline = src_ppr.find(qn("w:outlineLvl"))
                if outline is not None:
                    dst_ppr = dst._p.get_or_add_pPr()
                    dst_ppr.append(outline.__deepcopy__(outline))
        except Exception:
            pass

    def _copy_run_format(self, src_run, dst_run) -> None:  # type: ignore[no-untyped-def]
        """复制 run 级格式：粗斜下划线、字号、颜色、字体名。"""
        try:
            dst_run.bold = src_run.bold
            dst_run.italic = src_run.italic
            dst_run.underline = src_run.underline
            if src_run.font.size is not None:
                dst_run.font.size = src_run.font.size
            if src_run.font.color is not None and src_run.font.color.rgb is not None:
                dst_run.font.color.rgb = src_run.font.color.rgb
            if src_run.font.name:
                dst_run.font.name = src_run.font.name
            # 中文字体
            src_rpr = src_run._element.find(qn("w:rPr"))
            if src_rpr is not None:
                rfonts = src_rpr.find(qn("w:rFonts"))
                if rfonts is not None:
                    east = rfonts.get(qn("w:eastAsia"))
                    if east:
                        dst_rpr = dst_run._element.get_or_add_rPr()
                        dst_fonts = dst_rpr.find(qn("w:rFonts"))
                        if dst_fonts is None:
                            from docx.oxml import OxmlElement

                            dst_fonts = OxmlElement("w:rFonts")
                            dst_rpr.append(dst_fonts)
                        dst_fonts.set(qn("w:eastAsia"), east)
        except Exception as exc:
            log.debug("run 格式复制失败: %s", exc)

    # ------------------------------------------------------------------ #
    # 图片
    # ------------------------------------------------------------------ #
    @staticmethod
    def _run_has_image(run) -> bool:  # type: ignore[no-untyped-def]
        el = run._element
        return bool(el.findall(".//" + qn("w:drawing")) or el.findall(".//" + qn("w:pict")))

    def _copy_image_run(self, run, ctx: ConvertContext, paragraph=None) -> bool:  # type: ignore[no-untyped-def]
        """把源文档里的图片取出来，等比缩放后插进目标文档。"""
        try:
            blips = run._element.findall(".//" + qn("a:blip"))
            if not blips:
                return False
            shown = False
            for blip in blips:
                embed = blip.get(qn("r:embed")) or blip.get(qn("r:link"))
                if not embed:
                    continue
                part = run.part.related_parts.get(embed)
                if part is None:
                    continue
                blob = part.blob
                # 原图尺寸：优先取 docx 里记录的显示尺寸，取不到就按像素算
                w_emu, h_emu = self._drawing_extent(run)
                if w_emu is None or h_emu is None:
                    from PIL import Image

                    with Image.open(io.BytesIO(blob)) as im:
                        px_w, px_h = im.size
                    w_emu = int(px_w * 914400 / 96)
                    h_emu = int(px_h * 914400 / 96)

                target_w, target_h = fit_size(
                    w_emu, h_emu, ctx.box.usable_width, ctx.box.usable_height
                )
                if paragraph is not None:
                    pic_run = paragraph.add_run()
                    pic_run.add_picture(io.BytesIO(blob), width=Emu(target_w), height=Emu(target_h))
                else:
                    add_picture_centered(
                        ctx.doc, io.BytesIO(blob), target_w, target_h, alt="源文档图片"
                    )
                shown = True
            return shown
        except Exception as exc:
            log.warning("图片复制失败（已跳过）: %s", exc)
            return False

    @staticmethod
    def _drawing_extent(run) -> tuple[int | None, int | None]:  # type: ignore[no-untyped-def]
        """从 drawing 的 wp:extent 里取显示尺寸（EMU）。"""
        try:
            extents = run._element.findall(".//" + qn("wp:extent"))
            if extents:
                cx = extents[0].get("cx")
                cy = extents[0].get("cy")
                if cx and cy:
                    return int(cx), int(cy)
        except Exception:
            pass
        return None, None

    # ------------------------------------------------------------------ #
    # 表格
    # ------------------------------------------------------------------ #
    def _copy_table(self, table: Table, ctx: ConvertContext, result: ConvertResult) -> bool:
        """复制表格：按目标可用宽度等比缩放列宽，必要时降字号。"""
        rows = table.rows
        cols = table.columns
        n_rows, n_cols = len(rows), len(cols)
        if n_rows == 0 or n_cols == 0:
            return False

        # 源表列宽
        src_widths = []
        for col in cols:
            try:
                w = col.width
            except Exception:
                w = None
            src_widths.append(int(w) if w else 0)
        if not any(src_widths):
            total = ctx.box.usable_width
            src_widths = [total // n_cols] * n_cols

        src_total = sum(src_widths)
        max_width = ctx.box.usable_width
        scale = min(1.0, max_width / src_total) if src_total else 1.0
        target_widths = [max(1, int(w * scale)) for w in src_widths]

        # 列太宽 -> 字号等比缩小，避免文字挤成竖排
        font_pt = DEFAULT_FONT_PT
        if scale < 1.0:
            font_pt = max(MIN_FONT_PT, DEFAULT_FONT_PT * max(0.55, scale))
            if scale < 0.75:
                result.warn(
                    f"表格 {n_cols} 列总宽超出页面，已压缩至 {int(scale * 100)}% 并缩小字号到 {font_pt:.1f}pt"
                )

        new_table = ctx.doc.add_table(rows=n_rows, cols=n_cols)
        style_data_table(new_table, font_pt=font_pt)
        set_table_layout_fixed(new_table)

        for r_idx, row in enumerate(rows):
            for c_idx, cell in enumerate(row.cells):
                if c_idx >= n_cols:
                    continue
                text = self._cell_text(cell)
                try:
                    target_cell = new_table.cell(r_idx, c_idx)
                except Exception:
                    continue
                set_cell_text(target_cell, text, bold=(r_idx == 0), size_pt=font_pt)
            # 行高照搬（有设定时）
            try:
                if row.height is not None:
                    set_row_height_pt(new_table.rows[r_idx], max(12.0, row.height.pt))
            except Exception:
                pass

        # 落列宽（先设 cell.width 再设 column.width，Word 才认）
        for c_idx, w in enumerate(target_widths):
            for row in new_table.rows:
                try:
                    row.cells[c_idx].width = Emu(w)
                except Exception:
                    pass
            try:
                new_table.columns[c_idx].width = Emu(w)
            except Exception:
                pass

        repeat_header_row(new_table)
        ctx.doc.add_paragraph()  # 表后留一个空段，避免表格和下一段粘在一起
        return True

    @staticmethod
    def _cell_text(cell) -> str:  # type: ignore[no-untyped-def]
        """取单元格文本，段间用换行连接。"""
        try:
            parts = [p.text for p in cell.paragraphs]
        except Exception:
            return ""
        return "\n".join(p for p in parts if p is not None).strip("\n")

    # ------------------------------------------------------------------ #
    # 分节方向
    # ------------------------------------------------------------------ #
    @staticmethod
    def _landscape_section_indexes(source) -> list[int]:  # type: ignore[no-untyped-def]
        """找出源文档里横向的 sectPr 序号。"""
        out = []
        try:
            for i, sect_pr in enumerate(source.element.body.iter(qn("w:sectPr"))):
                if DocxConverter._sectpr_is_landscape(sect_pr):
                    out.append(i)
        except Exception:
            pass
        return out

    @staticmethod
    def _sectpr_is_landscape(sect_pr) -> bool:  # type: ignore[no-untyped-def]
        """判断 sectPr 是否为横向页（比较 pgSz 的 w/h）。"""
        try:
            pg_sz = sect_pr.find(qn("w:pgSz"))
            if pg_sz is None:
                return False
            w = int(pg_sz.get(qn("w:w")) or 0)
            h = int(pg_sz.get(qn("w:h")) or 0)
            orient = pg_sz.get(qn("w:orient"))
            if orient == "landscape":
                return True
            return w > h > 0
        except Exception:
            return False


def iter_block_items(parent) -> Iterable[object]:  # type: ignore[no-untyped-def]
    """按文档顺序遍历段落和表格（调试/测试用）。"""
    for child in parent.element.body.iterchildren():
        if child.tag == qn("w:p"):
            yield Paragraph(child, parent)
        elif child.tag == qn("w:tbl"):
            yield Table(child, parent)
