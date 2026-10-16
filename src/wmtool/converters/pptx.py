# -*- coding: utf-8 -*-
"""PowerPoint (.pptx) -> Word。

三种模式，默认走 **hybrid**（保真优先 + 文本可编辑兜底）：

``slide``
    整页幻灯片渲染成一张图，一页对一页，**版式 100% 不丢**，但文字不可选。
    这是渲染服务不可用时的默认模式（纯 Python，无外部依赖）。
``outline``
    提取标题/正文/表格/备注，重排成 Word 版式。
    文字可编辑、可搜索，但**复杂版式会变样**。
``hybrid``
    页面截图保真 + 紧随其后附一份可编辑的文字大纲（含备注）。
    既不丢版式也不丢可复制内容，代价是页数翻倍——这是默认值。

渲染通道（前两个需要外部程序，缺失时自动降级到内置渲染）：
1. Microsoft PowerPoint COM（装 Office 的主机上，最保真）
2. LibreOffice ``soffice --convert-to pdf``（便携场景）
3. **内置渲染器**（PyMuPDF 直接按形状画，零依赖，装不装 Office 都能跑）
"""

from __future__ import annotations

import io
import math
import os
import shutil
import subprocess
import tempfile
from pathlib import Path

import pymupdf
from PIL import Image

from ..docx_util import (
    add_caption,
    add_heading,
    add_picture_centered,
    new_section,
    set_cell_text,
    set_table_layout_fixed,
    style_data_table,
)
from ..geometry import fit_size
from ..logging_setup import get_logger
from .base import ConvertContext, ConvertResult

log = get_logger("conv.pptx")

SUPPORTED_SUFFIXES = {".pptx", ".pptm", ".ppsx"}

#: 渲染 DPI（slide 模式）
RENDER_DPI = 144
#: 单页像素上限
MAX_PAGE_PIXELS = 30_000_000


# --------------------------------------------------------------------------- #
# 渲染通道
# --------------------------------------------------------------------------- #
def _find_soffice() -> str | None:
    """找 LibreOffice / OpenOffice 的可执行文件。"""
    for name in ("soffice", "soffice.exe", "libreoffice"):
        found = shutil.which(name)
        if found:
            return found
    candidates = [
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
        r"C:\Program Files\OpenOffice 4\program\soffice.exe",
    ]
    for c in candidates:
        if Path(c).is_file():
            return c
    return None


def _word_installed() -> bool:
    """本机是否装了 Microsoft Word（只查注册表/文件，不起进程）。"""
    for p in (
        r"C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE",
        r"C:\Program Files (x86)\Microsoft Office\root\Office16\WINWORD.EXE",
        r"C:\Program Files\Microsoft Office\Office16\WINWORD.EXE",
    ):
        if Path(p).is_file():
            return True
    try:
        import winreg  # type: ignore

        with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, r"Word.Application\CurVer"):
            return True
    except Exception:
        return False


def _powerpoint_installed() -> bool:
    """本机是否装了 PowerPoint。"""
    for p in (
        r"C:\Program Files\Microsoft Office\root\Office16\POWERPNT.EXE",
        r"C:\Program Files (x86)\Microsoft Office\root\Office16\POWERPNT.EXE",
    ):
        if Path(p).is_file():
            return True
    try:
        import winreg  # type: ignore

        with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, r"PowerPoint.Application\CurHelp"):
            return True
    except Exception:
        return False


def render_via_powerpoint(src: Path, out_dir: Path) -> list[Path] | None:
    """用 PowerPoint COM 把 pptx 逐页导出成 PNG（最保真）。失败返回 None。"""
    if os.name != "nt":
        return None
    try:
        import pythoncom  # type: ignore
        import win32com.client  # type: ignore
    except Exception:
        return None

    app = None
    pres = None
    try:
        pythoncom.CoInitialize()
        app = win32com.client.DispatchEx("PowerPoint.Application")
        # 只读打开：WithWindow=False 时 PowerPoint 仍会常驻，故最后必须 Quit
        pres = app.Presentations.Open(str(src), ReadOnly=True, Untitled=False, WithWindow=False)
        out_dir.mkdir(parents=True, exist_ok=True)
        pres.Export(str(out_dir), "PNG", 1600, 900)
        if pres.Saved is False:
            pres.Saved = True  # 只读打开，绝不写回源文件
        files = sorted(out_dir.glob("*.PNG")) + sorted(out_dir.glob("*.png"))
        files = sorted(set(files), key=lambda p: p.name)
        return files or None
    except Exception as exc:
        log.warning("PowerPoint 渲染不可用: %s", exc)
        return None
    finally:
        for obj in (pres, app):
            try:
                if obj is not None:
                    obj.Close() if obj is pres else obj.Quit()
            except Exception:
                pass
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass


def render_via_libreoffice(src: Path, out_dir: Path) -> list[Path] | None:
    """用 LibreOffice 把 pptx 转成 PDF，再交给 PyMuPDF 渲染。"""
    soffice = _find_soffice()
    if not soffice:
        return None
    try:
        out_dir.mkdir(parents=True, exist_ok=True)
        subprocess.run(
            [
                soffice,
                "--headless",
                "--norestore",
                "--convert-to",
                "pdf",
                "--outdir",
                str(out_dir),
                str(src),
            ],
            check=True,
            timeout=300,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        )
        pdfs = list(out_dir.glob("*.pdf"))
        if not pdfs:
            return None
        docs = pymupdf.open(str(pdfs[0]))
        pages = []
        for i in range(docs.page_count):
            out = out_dir / f"slide-{i + 1:04d}.png"
            pix = docs.load_page(i).get_pixmap(dpi=RENDER_DPI, alpha=False)
            pix.save(str(out))
            pages.append(out)
        docs.close()
        return pages or None
    except Exception as exc:
        log.warning("LibreOffice 渲染不可用: %s", exc)
        return None
    finally:
        _cleanup_soffice_lock(out_dir)


def _cleanup_soffice_lock(out_dir: Path) -> None:
    """清掉 LibreOffice 可能留下的锁目录（.~lock.*）。"""
    try:
        for p in out_dir.glob(".~lock.*"):
            p.unlink(missing_ok=True)
    except Exception:
        pass


# --------------------------------------------------------------------------- #
# 内置渲染器（零依赖兜底）
# --------------------------------------------------------------------------- #
def render_builtin(src: Path, result: ConvertResult) -> list[bytes]:
    """内置渲染：用 python-pptx 读形状 + PyMuPDF 画图，不需要装 Office。

    支持：背景色、矩形/圆角矩形/椭圆、图片、文本框（含自动换行与字号），
    表格画成网格 + 文字。渐变/阴影/艺术字等特效不还原（会在结果里给出警告）。
    """
    from pptx import Presentation
    from pptx.util import Emu as PptxEmu

    prs = Presentation(str(src))
    slide_w = int(prs.slide_width or 12192000)
    slide_h = int(prs.slide_height or 6858000)

    # 目标像素尺寸：按 144 DPI 换算，但受单页像素上限约束
    scale = RENDER_DPI / 914400.0
    px_w = max(320, int(slide_w * scale))
    px_h = max(180, int(slide_h * scale))
    if px_w * px_h > MAX_PAGE_PIXELS:
        k = math.sqrt(MAX_PAGE_PIXELS / (px_w * px_h))
        px_w, px_h = int(px_w * k), int(px_h * k)

    out: list[bytes] = []
    unsupported = set()

    for s_idx, slide in enumerate(prs.slides):
        # 画布缩放：EMU(ppt) -> 像素
        kx = px_w / slide_w
        ky = px_h / slide_h
        page = pymupdf.open()
        page_obj = page.new_page(width=px_w, height=px_h)
        _draw_background(page_obj, slide, px_w, px_h)

        for shape in slide.shapes:
            try:
                _draw_shape(page_obj, shape, kx, ky, unsupported)
            except Exception as exc:
                log.debug("内置渲染：第 %d 页某形状跳过 (%s)", s_idx + 1, exc)

        pix = page_obj.get_pixmap(alpha=False)
        out.append(pix.tobytes("png"))
        page.close()

    if unsupported:
        result.warn(
            "内置渲染器未还原的特效: " + "、".join(sorted(unsupported))
            + "（装 LibreOffice 或 PowerPoint 可提升保真度）"
        )
    return out


def _draw_background(page_obj, slide, px_w: int, px_h: int) -> None:  # type: ignore[no-untyped-def]
    """铺背景色（取不到就用白）。"""
    color = (1, 1, 1)
    try:
        bg = slide.background.fill
        if bg.type is not None and str(bg.type) != "MSO_FILL_TYPE.BACKGROUND (5)":
            rgb = bg.fore_color.rgb
            color = (rgb[0] / 255, rgb[1] / 255, rgb[2] / 255)
    except Exception:
        pass
    page_obj.draw_rect(pymupdf.Rect(0, 0, px_w, px_h), color=color, fill=color, width=0)


def _shape_rgb(shape) -> tuple[float, float, float] | None:  # type: ignore[no-untyped-def]
    """取形状填充色。"""
    try:
        fill = shape.fill
        if fill.type is None or "BACKGROUND" in str(fill.type):
            return None
        rgb = fill.fore_color.rgb
        return (rgb[0] / 255, rgb[1] / 255, rgb[2] / 255)
    except Exception:
        return None


def _font_size_px(run_or_para, kx: float, ky: float) -> float:  # type: ignore[no-untyped-def]
    """字号：pt -> 像素（按页面缩放比换算）。"""
    pt = None
    try:
        pt = run_or_para.font.size.pt  # type: ignore[union-attr]
    except Exception:
        pt = None
    if not pt:
        pt = 18.0
    return max(4.0, pt * RENDER_DPI / 72.0)


def _draw_shape(page_obj, shape, kx: float, ky: float, unsupported: set) -> None:  # type: ignore[no-untyped-def]
    """画一个形状。"""
    from pptx.enum.shapes import MSO_SHAPE_TYPE

    left = float(shape.left or 0) * kx
    top = float(shape.top or 0) * ky
    width = max(1.0, float(shape.width or 0) * kx)
    height = max(1.0, float(shape.height or 0) * ky)
    rect = pymupdf.Rect(left, top, left + width, top + height)

    stype = shape.shape_type

    if stype == MSO_SHAPE_TYPE.PICTURE:
        try:
            blob = shape.image.blob
            page_obj.insert_image(rect, stream=blob, keep_proportion=True)
        except Exception as exc:
            log.debug("内置渲染：图片插入失败 %s", exc)
        return

    if stype == MSO_SHAPE_TYPE.GROUP:
        for sub in shape.shapes:
            try:
                _draw_shape(page_obj, sub, kx, ky, unsupported)
            except Exception:
                pass
        return

    if getattr(shape, "has_table", False):
        _draw_table(page_obj, shape, rect, kx, ky)
        return

    # 普通形状：填充 + 边框
    fill = _shape_rgb(shape)
    stroke = None
    try:
        line = shape.line
        if line.fill.type is not None and "BACKGROUND" not in str(line.fill.type):
            rgb = line.color.rgb
            stroke = (rgb[0] / 255, rgb[1] / 255, rgb[2] / 255)
    except Exception:
        pass

    try:
        auto = shape.auto_shape_type
        if str(auto).startswith("OVAL") or str(auto).startswith("MSO_SHAPE.OVAL"):
            page_obj.draw_oval(rect, color=stroke, fill=fill, width=1)
        else:
            radius = 0.0
            if "ROUNDED" in str(auto):
                radius = min(width, height) * 0.12
            if radius > 0:
                page_obj.draw_rect(rect, color=stroke, fill=fill, width=1, radius=radius)
            else:
                page_obj.draw_rect(rect, color=stroke, fill=fill, width=1)
    except Exception:
        page_obj.draw_rect(rect, color=stroke, fill=fill, width=1)

    # 文本
    if getattr(shape, "has_text_frame", False) and shape.text_frame.text.strip():
        _draw_text_frame(page_obj, shape.text_frame, rect, kx, ky)


def _draw_text_frame(page_obj, tf, rect, kx: float, ky: float) -> None:  # type: ignore[no-untyped-def]
    """把文本框内容画进矩形（按段落逐行排，超宽自动换行）。"""
    # 文本框内边距（约 0.1 英寸）
    pad = 9.0
    avail_w = max(8.0, rect.width - pad * 2)
    y = rect.y0 + pad
    line_h_default = 16.0

    for para in tf.paragraphs:
        text = "".join(r.text for r in para.runs)
        if not text.strip():
            y += line_h_default * 0.6
            continue

        size_px = line_h_default
        color = (0, 0, 0)
        bold = False
        align = "left"
        try:
            if para.runs:
                size_px = _font_size_px(para.runs[0], kx, ky)
                bold = bool(para.runs[0].font.bold)
                rgb = para.runs[0].font.color
                if rgb is not None and rgb.type is not None:
                    c = rgb.rgb
                    color = (c[0] / 255, c[1] / 255, c[2] / 255)
        except Exception:
            pass
        try:
            align = str(para.alignment or "").lower()
        except Exception:
            pass

        line_h = size_px * 1.25
        # 简单换行：按可用宽度估算每行能放几个字符（中文按 1 个字号宽）
        for line in _wrap_text(text, avail_w, size_px):
            if y + line_h > rect.y1 + line_h and rect.height > 0:
                break  # 超出文本框不再画，保持和 PPT 一致的可视区域
            x = rect.x0 + pad
            if "center" in align:
                tw = _text_width(line, size_px)
                x = rect.x0 + max(0, (rect.width - tw) / 2)
            elif "right" in align:
                tw = _text_width(line, size_px)
                x = rect.x1 - pad - tw
            try:
                page_obj.insert_text(
                    pymupdf.Point(x, y + size_px),
                    line,
                    fontsize=size_px,
                    fontname="china-s",  # PyMuPDF 内置简体中文字体
                    color=color,
                    render_mode=0 if not bold else 2,
                )
            except Exception:
                page_obj.insert_text(
                    pymupdf.Point(x, y + size_px), line, fontsize=size_px, color=color
                )
            y += line_h
        y += size_px * 0.35


def _wrap_text(text: str, avail_w: float, size_px: float) -> list[str]:
    """按可用宽度粗略折行（中英文混排够用）。"""
    lines: list[str] = []
    for raw in text.split("\n"):
        if not raw:
            lines.append("")
            continue
        cur = ""
        for ch in raw:
            w = _text_width(cur + ch, size_px)
            if w > avail_w and cur:
                lines.append(cur)
                cur = ch
            else:
                cur += ch
        if cur:
            lines.append(cur)
    return lines


def _text_width(text: str, size_px: float) -> float:
    """估算文本像素宽：ASCII 约 0.55 字号宽，中日韩按 1.0。"""
    total = 0.0
    for ch in text:
        total += size_px * (1.0 if ord(ch) > 0x2E80 else 0.55)
    return total


def _draw_table(page_obj, shape, rect, kx: float, ky: float) -> None:  # type: ignore[no-untyped-def]
    """把 PPT 表格画成网格 + 文字。"""
    table = shape.table
    n_rows = len(table.rows)
    n_cols = len(table.columns)
    if n_rows == 0 or n_cols == 0:
        return

    # 列宽/行高按 PPT 里的实际值换算
    col_w = [max(4.0, float(c.width or 0) * kx) for c in table.columns]
    total_w = sum(col_w) or rect.width
    if abs(total_w - rect.width) > 1:
        k = rect.width / total_w
        col_w = [w * k for w in col_w]
    row_h = [max(4.0, float(r.height or 0) * ky) for r in table.rows]
    total_h = sum(row_h) or rect.height
    if total_h > 0 and abs(total_h - rect.height) > 1:
        k = rect.height / total_h
        row_h = [h * k for h in row_h]

    y = rect.y0
    for r_idx in range(n_rows):
        x = rect.x0
        for c_idx in range(n_cols):
            cell_rect = pymupdf.Rect(x, y, x + col_w[c_idx], y + row_h[r_idx])
            try:
                fill = None
                cell = table.cell(r_idx, c_idx)
                if cell.fill.type is not None and "BACKGROUND" not in str(cell.fill.type):
                    rgb = cell.fill.fore_color.rgb
                    fill = (rgb[0] / 255, rgb[1] / 255, rgb[2] / 255)
            except Exception:
                fill = None
            page_obj.draw_rect(cell_rect, color=(0.4, 0.4, 0.4), fill=fill, width=0.6)

            text = table.cell(r_idx, c_idx).text.strip()
            if text:
                size = max(5.0, min(14.0, row_h[r_idx] * 0.42))
                lines = _wrap_text(text, col_w[c_idx] - 6, size)
                ty = y + size
                for line in lines[: max(1, int(row_h[r_idx] // (size * 1.15)))]:
                    page_obj.insert_text(
                        pymupdf.Point(x + 3, ty),
                        line,
                        fontsize=size,
                        fontname="china-s",
                        color=(0, 0, 0),
                    )
                    ty += size * 1.15
            x += col_w[c_idx]
        y += row_h[r_idx]


# --------------------------------------------------------------------------- #
# 转换器
# --------------------------------------------------------------------------- #
class PptxConverter:
    """PowerPoint -> Word（一页幻灯片 = 一页 Word）。"""

    name = "pptx"
    display = "PowerPoint"

    def __init__(self, mode: str = "hybrid"):
        self.mode = mode if mode in ("slide", "outline", "hybrid") else "hybrid"

    def convert(self, ctx: ConvertContext) -> ConvertResult:
        result = ConvertResult(converter=self.name)
        src = ctx.src

        try:
            from pptx import Presentation

            prs = Presentation(str(src))
            slide_count = len(prs.slides)
            slide_w = int(prs.slide_width or 12192000)
            slide_h = int(prs.slide_height or 6858000)
        except Exception as exc:
            raise ValueError(f"PPTX 打开失败: {exc}") from exc

        if slide_count == 0:
            raise ValueError("演示文稿没有任何幻灯片")

        # 幻灯片天然是横的（16:9 / 4:3），Word 也得用横向页，否则必缩成一条
        box = ctx.landscape_box
        result.landscape = True
        ctx.doc.add_page_break()
        new_section(ctx.doc, box)

        if self.mode in ("slide", "hybrid"):
            images = self._render_slides(src, result)
            if images:
                self._write_slide_images(ctx, images, box, result)
            else:
                result.warn("所有渲染通道都不可用，已降级为纯文字大纲（版式无法还原）")
                self.mode = "outline"

        if self.mode in ("outline", "hybrid"):
            if self.mode == "hybrid":
                ctx.doc.add_page_break()
                new_section(ctx.doc, box)
                add_heading(ctx.doc, "可编辑文字大纲（便于复制检索）", level=2, size_pt=13)
            self._write_outline(ctx, src, result)

        result.summary = f"PowerPoint {slide_count} 页（{self.mode} 模式）"
        ctx.report(1.0, "PowerPoint 已写入")

        # 用完横向，切回纵向，后续文件不受影响
        ctx.doc.add_page_break()
        new_section(ctx.doc, ctx.box)
        return result

    # ------------------------------------------------------------------ #
    def _render_slides(self, src: Path, result: ConvertResult) -> list[bytes]:
        """按 PowerPoint -> LibreOffice -> 内置渲染 的顺序取每页位图。"""
        workdir = Path(tempfile.mkdtemp(prefix="wmtool-ppt-"))
        try:
            # 通道 1：PowerPoint COM
            if _powerpoint_installed():
                files = render_via_powerpoint(src, workdir / "ppt")
                if files:
                    log.info("PPTX 渲染通道：PowerPoint（%d 页）", len(files))
                    return [f.read_bytes() for f in files]

            # 通道 2：LibreOffice
            files = render_via_libreoffice(src, workdir / "lo")
            if files:
                log.info("PPTX 渲染通道：LibreOffice（%d 页）", len(files))
                return [f.read_bytes() for f in files]

            # 通道 3：内置渲染（零依赖）
            log.info("PPTX 渲染通道：内置渲染器")
            return render_builtin(src, result)
        finally:
            shutil.rmtree(workdir, ignore_errors=True)

    def _write_slide_images(
        self, ctx: ConvertContext, images: list[bytes], box, result: ConvertResult
    ) -> None:  # type: ignore[no-untyped-def]
        """把每页幻灯片位图写进 Word（一页一张，等比缩放不出框）。"""
        for i, blob in enumerate(images):
            with Image.open(io.BytesIO(blob)) as im:
                px_w, px_h = im.size
            w_emu = int(px_w * 914400 / 96)
            h_emu = int(px_h * 914400 / 96)
            tw, th = fit_size(w_emu, h_emu, box.usable_width, box.usable_height)
            add_picture_centered(ctx.doc, io.BytesIO(blob), tw, th, alt=f"幻灯片 {i + 1}")
            result.image_count += 1
            if i < len(images) - 1:
                ctx.doc.add_page_break()
            frac = (i + 1) / max(1, len(images))
            ctx.report(frac * 0.8, f"幻灯片 {i + 1}/{len(images)} 已嵌入")

    def _write_outline(self, ctx: ConvertContext, src: Path, result: ConvertResult) -> None:
        """提取每页的标题/正文/表格/备注，重排成 Word 内容。"""
        from pptx import Presentation

        prs = Presentation(str(src))
        bad = set()
        for idx, slide in enumerate(prs.slides, start=1):
            title = self._slide_title(slide)
            add_heading(
                ctx.doc,
                f"第 {idx} 页：{title}" if title else f"第 {idx} 页",
                level=2,
                size_pt=12,
            )

            for shape in slide.shapes:
                try:
                    if getattr(shape, "has_table", False):
                        self._outline_table(ctx, shape.table)
                    elif getattr(shape, "has_text_frame", False):
                        self._outline_text(ctx, shape.text_frame)
                    elif shape.shape_type is not None and "PICTURE" in str(shape.shape_type):
                        pass  # 图已在 slide 模式里保真呈现，大纲不重复插图
                except Exception as exc:
                    bad.add(type(exc).__name__)

            notes = self._slide_notes(slide)
            if notes:
                p = ctx.doc.add_paragraph()
                run = p.add_run(f"备注：{notes}")
                from docx.shared import Pt, RGBColor

                run.font.size = Pt(9)
                run.font.color.rgb = RGBColor(0x60, 0x60, 0x60)

        if bad:
            result.warn("大纲提取时部分形状被跳过: " + "、".join(sorted(bad)))

    @staticmethod
    def _slide_title(slide) -> str:  # type: ignore[no-untyped-def]
        try:
            if slide.shapes.title is not None:
                return slide.shapes.title.text.strip().replace("\n", " ")
        except Exception:
            pass
        return ""

    @staticmethod
    def _slide_notes(slide) -> str:  # type: ignore[no-untyped-def]
        try:
            if slide.has_notes_slide:
                return slide.notes_slide.notes_text_frame.text.strip()
        except Exception:
            pass
        return ""

    def _outline_text(self, ctx: ConvertContext, tf) -> None:  # type: ignore[no-untyped-def]
        """把一个文本框写成段落（按级别缩进）。"""
        for para in tf.paragraphs:
            text = "".join(r.text for r in para.runs).strip()
            if not text:
                continue
            p = ctx.doc.add_paragraph()
            try:
                lvl = int(para.level or 0)
            except Exception:
                lvl = 0
            if lvl:
                p.paragraph_format.left_indent = EmuPt(18 * lvl)
            run = p.add_run(("• " if lvl else "") + text)
            try:
                if para.runs and para.runs[0].font.size:
                    run.font.size = para.runs[0].font.size
                if para.runs and para.runs[0].font.bold:
                    run.bold = True
            except Exception:
                pass

    def _outline_table(self, ctx: ConvertContext, table) -> None:  # type: ignore[no-untyped-def]
        """把 PPT 表格写成 Word 表格（按可用宽度压缩列宽）。"""
        n_rows = len(table.rows)
        n_cols = len(table.columns)
        if n_rows == 0 or n_cols == 0:
            return
        usable = ctx.box.usable_width
        each = max(1, usable // n_cols)
        font_pt = 9.0 if n_cols <= 6 else 7.0

        new_table = ctx.doc.add_table(rows=n_rows, cols=n_cols)
        style_data_table(new_table, font_pt=font_pt)
        set_table_layout_fixed(new_table)
        for r in range(n_rows):
            for c in range(n_cols):
                text = table.cell(r, c).text.strip()
                set_cell_text(new_table.cell(r, c), text, bold=(r == 0), size_pt=font_pt)
        for c in range(n_cols):
            for row in new_table.rows:
                try:
                    row.cells[c].width = Emu(each)
                except Exception:
                    pass
        ctx.doc.add_paragraph()


def EmuPt(pt: float):  # type: ignore[no-untyped-def]
    """小工具：磅 -> EMU（局部用，避免污染 geometry 的命名）。"""
    from docx.shared import Emu

    return Emu(int(pt * 12700))
