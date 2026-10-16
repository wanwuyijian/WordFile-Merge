# -*- coding: utf-8 -*-
"""PDF 转换器：**一页 PDF = Word 里的一页**，不 OCR。

做法是最朴素也最可靠的一种——把 PDF 每页渲染成位图贴进 Word：

* 不解析文字层、不做 OCR，所以扫描件和文字版 PDF 走同一条路、结果一致；
* 页面方向自适应：横向的 PDF 页自动放到 Word 的横向节里，
  纵向的放纵向节，因此""一页对一页""且**不变形**；
* 渲染 DPI 可调，默认 200（清晰度和体积的平衡点），超大页面自动降档。

代价要说清楚：合并出来的 Word 里 PDF 部分是**图**，不能选中文字。
这是"不 OCR + 一页对一页"的必然结果。
"""

from __future__ import annotations

import io
from pathlib import Path

import pymupdf  # PyMuPDF
from PIL import Image

from ..docx_util import add_caption, add_picture_centered, new_section
from ..geometry import fit_size
from ..logging_setup import get_logger
from .base import ConvertContext, ConvertResult

log = get_logger("conv.pdf")

SUPPORTED_SUFFIXES = {".pdf"}

#: 渲染 DPI：200 在 A4 上约 1650×2340 像素，打印清晰、体积可控
DEFAULT_DPI = 200
MIN_DPI = 110
#: 单页渲染后的像素上限（超过则自动降 DPI，防止超大图纸把内存打爆）
MAX_PAGE_PIXELS = 40_000_000


class PdfConverter:
    """把 PDF 逐页渲染并嵌入 Word。"""

    name = "pdf"
    display = "PDF"

    def convert(self, ctx: ConvertContext) -> ConvertResult:
        result = ConvertResult(converter=self.name)
        src = ctx.src

        try:
            doc = pymupdf.open(str(src))
        except Exception as exc:
            raise ValueError(f"PDF 打开失败（可能已损坏或加密）: {exc}") from exc

        try:
            if doc.needs_pass:
                raise ValueError("PDF 已加密，需要密码，暂不支持（源文件未被修改）")
            page_count = doc.page_count
            if page_count == 0:
                raise ValueError("PDF 没有任何页面")

            # 逐页判断方向，统计后决定这一节用纵向还是横向
            orientations: list[bool] = []  # True = 该页是横向
            for pno in range(page_count):
                rect = doc.load_page(pno).rect
                orientations.append(rect.width > rect.height * 1.05)

            # 横竖混排的 PDF：多数页是横的才整体用横向，否则一律纵向。
            # 注意——无论走哪条路，**图片尺寸必须按实际所在的节来算**，
            # 不能"先决定 box 再切节"，那样两者一旦不一致图片就会超页。
            landscape_pages = sum(orientations)
            use_landscape = landscape_pages / max(1, page_count) > 0.6
            result.landscape = use_landscape

            if use_landscape:
                _switch_section_if_needed(ctx, ctx.landscape_box)
                result.warn(
                    f"{page_count} 页中有 {landscape_pages} 页为横向，已按横向排版以保证一页对一页"
                )

            embedded = 0
            for pno in range(page_count):
                page = doc.load_page(pno)
                is_land = bool(orientations[pno])
                img = self._render_page(page, is_land)

                # 该页是横的就放横向节，是竖的就放纵向节——
                # 一页一张图，方向和页都跟着这一页走，两边永远一致。
                target_box = ctx.landscape_box if is_land else ctx.box
                _switch_section_if_needed(ctx, target_box)

                px_w, px_h = img.size
                emu_per_px = 914400 / 96.0
                nat_w, nat_h = px_w * emu_per_px, px_h * emu_per_px
                target_w, target_h = fit_size(
                    nat_w, nat_h, target_box.usable_width, target_box.usable_height
                )

                buf = self._encode(img)
                add_picture_centered(ctx.doc, buf, target_w, target_h,
                                     alt=f"{src.name} 第 {pno + 1} 页")
                embedded += 1

                # 每页之间分页——这就是"一页占一页"的落点
                if pno < page_count - 1:
                    ctx.doc.add_page_break()

                frac = (pno + 1) / page_count
                ctx.report(frac, f"PDF 渲染中 {pno + 1}/{page_count} 页")

            result.image_count = embedded
            result.summary = f"PDF {page_count} 页（逐页位图，未 OCR）"
            if embedded != page_count:
                result.warn(f"期望嵌入 {page_count} 页，实际 {embedded} 页")
        finally:
            try:
                doc.close()
            except Exception:
                pass

        return result

    # ------------------------------------------------------------------ #
    @staticmethod
    def _render_page(page, landscape: bool) -> Image.Image:  # type: ignore[no-untyped-def]
        """把一页 PDF 渲染成 PIL 图（自动按像素上限降 DPI）。

        ``landscape`` 是**目标页**的方向；PDF 页自带的方向与目标不一致时旋转 90°，
        这样一页内容既不会变形，也不会被裁掉。
        """
        dpi = DEFAULT_DPI
        rect = page.rect
        # 预估像素量，超上限就按比例降 DPI（下限 MIN_DPI）
        est = (rect.width / 72 * dpi) * (rect.height / 72 * dpi)
        if est > MAX_PAGE_PIXELS:
            scale = (MAX_PAGE_PIXELS / est) ** 0.5
            dpi = max(MIN_DPI, int(dpi * scale))

        matrix = pymupdf.Matrix(dpi / 72, dpi / 72)
        pix = page.get_pixmap(matrix=matrix, alpha=False)
        img = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)

        page_landscape = rect.width > rect.height * 1.05
        if page_landscape != landscape:
            img = img.rotate(90, expand=True)
        return img

    @staticmethod
    def _encode(img: Image.Image) -> io.BytesIO:
        """PDF 渲染出来是连续色调，用 JPEG 能显著压体积；线稿类仍可接受。"""
        buf = io.BytesIO()
        img.convert("RGB").save(buf, format="JPEG", quality=88, optimize=True)
        buf.seek(0)
        return buf


def pdf_page_count(path: Path) -> int | None:
    """快速取页数（GUI 预览用）。"""
    try:
        with pymupdf.open(str(path)) as doc:
            return doc.page_count
    except Exception:
        return None


def _switch_section_if_needed(ctx: ConvertContext, box) -> None:  # type: ignore[no-untyped-def]
    """确保文档当前节的方向与 ``box`` 一致，不一致就新开一个节。

    这是"图片尺寸"与"页面方向"保持同步的唯一开关：先切节、再按同一个 box 算图，
    两者就不可能打架。
    """
    sections = ctx.doc.sections
    if not sections:
        return
    current = sections[-1]
    cur_landscape = int(current.page_width or 0) > int(current.page_height or 0)
    want_landscape = box.width > box.height
    if cur_landscape == want_landscape:
        return
    ctx.doc.add_page_break()
    new_section(ctx.doc, box)
