# -*- coding: utf-8 -*-
"""图片转换器：一张图片 = 独占一页。

要点（对应需求第 3 条）：
* **不做任何 OCR**，图片原样嵌入，不识别、不重排；
* 图片按页面可用框**等比缩放**，完整落进一页，不裁剪、不跨页；
* 小图不放大（避免糊成大马赛克）；
* 多帧图（GIF/多页 TIFF）只取第一帧，并在结果里给出警告。
"""

from __future__ import annotations

import io
from pathlib import Path

from PIL import Image, ImageOps

from ..docx_util import add_caption, add_picture_centered
from ..geometry import fit_size
from ..logging_setup import get_logger
from .base import ConvertContext, ConvertResult

log = get_logger("conv.image")

SUPPORTED_SUFFIXES = {
    ".jpg", ".jpeg", ".jpe", ".jfif", ".png", ".bmp", ".gif", ".webp", ".tif", ".tiff", ".ico",
}

#: 降采样阈值：最长边超过它就缩小，避免单张手机照把 docx 撑到几十 MB
MAX_LONG_EDGE_PX = 6000


class ImageConverter:
    """图片 -> 独占一页。"""

    name = "image"
    display = "图片"

    def convert(self, ctx: ConvertContext) -> ConvertResult:
        result = ConvertResult(converter=self.name)
        src = ctx.src
        box = ctx.box

        try:
            with Image.open(src) as im:
                frame_count = getattr(im, "n_frames", 1)
                if frame_count > 1:
                    result.warn(f"动图/多页图共 {frame_count} 帧，仅合并第 1 帧")
                im.seek(0)
                im = ImageOps.exif_transpose(im) or im
                px_w, px_h = im.size
                if px_w <= 0 or px_h <= 0:
                    raise ValueError("图片尺寸非法")

                # 超大图降采样（仅缩不放）
                longest = max(px_w, px_h)
                if longest > MAX_LONG_EDGE_PX:
                    scale = MAX_LONG_EDGE_PX / longest
                    im = im.resize(
                        (max(1, int(px_w * scale)), max(1, int(px_h * scale))),
                        Image.LANCZOS,
                    )
                    px_w, px_h = im.size
                    result.warn(
                        f"原图最长边 {longest}px，已降采样至 {px_w}×{px_h}px 以控制体积"
                    )

                # 像素 -> EMU（96 DPI 基准），再等比塞进可用框
                w_emu = int(px_w * 914400 / 96)
                h_emu = int(px_h * 914400 / 96)
                target_w, target_h = fit_size(
                    w_emu, h_emu, box.usable_width, box.usable_height
                )
                buf = self._to_stream(im)
                add_picture_centered(ctx.doc, buf, target_w, target_h, alt=src.name)
        except ValueError:
            raise
        except Exception as exc:
            raise ValueError(f"图片读取失败: {exc}") from exc

        result.image_count = 1
        result.summary = f"图片 1 张（{px_w}×{px_h}px）"
        add_caption(ctx.doc, f"图 {ctx.index}：{src.name}（{px_w}×{px_h}px）")
        ctx.report(1.0, "图片已嵌入")
        return result

    @staticmethod
    def _to_stream(im: Image.Image) -> io.BytesIO:
        """把 PIL 图编码成内存流。

        统一转成 PNG/JPEG 再嵌入，绕开源文件里千奇百怪的色彩配置，
        保证 Word 那边一定打得开。
        """
        buf = io.BytesIO()
        if im.mode in ("RGBA", "LA") or (im.mode == "P" and "transparency" in im.info):
            im.convert("RGBA").save(buf, format="PNG", optimize=True)
        else:
            im.convert("RGB").save(buf, format="JPEG", quality=92, optimize=True)
        buf.seek(0)
        return buf


def image_size(path: Path) -> tuple[int, int] | None:
    """快速取图片像素尺寸（GUI 预览用）。"""
    try:
        with Image.open(path) as im:
            return im.size
    except Exception:
        return None
