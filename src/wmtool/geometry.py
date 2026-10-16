# -*- coding: utf-8 -*-
"""排版度量：页边距、可用宽度、图片/表格缩放。

单位统一用 **EMU**（English Metric Unit，1 英寸 = 914400 EMU），
python-docx 内部也是这个单位，避免来回换算丢精度。

这一层存在的唯一目的：**让任何内容都不出页面**。
图片按可用框等比缩放、表格按可用宽度等比缩字号、超宽的 PDF 页自动转横版。
"""

from __future__ import annotations

from dataclasses import dataclass

EMU_PER_INCH = 914400
EMU_PER_CM = 360000
EMU_PER_PT = 12700
EMU_PER_TWIP = 635

#: 页面尺寸（EMU）：A4 纵向
A4_WIDTH = int(8.27 * EMU_PER_INCH)
A4_HEIGHT = int(11.69 * EMU_PER_INCH)


@dataclass(frozen=True)
class Margins:
    """页边距（EMU）。"""

    top: int = int(2.54 * EMU_PER_CM)
    bottom: int = int(2.54 * EMU_PER_CM)
    left: int = int(2.54 * EMU_PER_CM)
    right: int = int(2.54 * EMU_PER_CM)

    @classmethod
    def uniform(cls, cm: float) -> "Margins":
        v = int(cm * EMU_PER_CM)
        return cls(top=v, bottom=v, left=v, right=v)

    @property
    def horizontal(self) -> int:
        return self.left + self.right

    @property
    def vertical(self) -> int:
        return self.top + self.bottom


#: 默认页边距：2.0cm，比 Word 默认的 2.54cm 略窄，给内容让一点空间。
DEFAULT_MARGINS = Margins.uniform(2.0)


@dataclass(frozen=True)
class PageBox:
    """一个页面的可用内容框。"""

    width: int
    height: int
    margins: Margins = DEFAULT_MARGINS

    @property
    def usable_width(self) -> int:
        return max(1, self.width - self.margins.horizontal)

    @property
    def usable_height(self) -> int:
        return max(1, self.height - self.margins.vertical)

    @property
    def landscape(self) -> bool:
        return self.width > self.height

    def to_landscape(self) -> "PageBox":
        return PageBox(width=self.height, height=self.width, margins=self.margins)

    def to_portrait(self) -> "PageBox":
        return PageBox(width=min(self.width, self.height), height=max(self.width, self.height), margins=self.margins)


#: A4 纵向可用框
PORTRAIT_BOX = PageBox(width=A4_WIDTH, height=A4_HEIGHT, margins=DEFAULT_MARGINS)
#: A4 横向可用框
LANDSCAPE_BOX = PageBox(
    width=A4_HEIGHT, height=A4_WIDTH, margins=DEFAULT_MARGINS
)


def fit_size(src_w: float, src_h: float, box_w: int, box_h: int) -> tuple[int, int]:
    """把 ``src_w × src_h`` 等比缩放进 ``box_w × box_h``，返回 EMU 尺寸。

    * 只缩不放：小图保持原尺寸，避免糊成大马赛克。
    * 尺寸非法（<=0）时返回 1×1，绝不让调用方拿到 0 或负数去写进 docx
      （python-docx 收到 0 宽度会抛异常）。
    """
    if src_w <= 0 or src_h <= 0:
        return 1, 1
    scale = min(box_w / src_w, box_h / src_h, 1.0)
    w = max(1, int(round(src_w * scale)))
    h = max(1, int(round(src_h * scale)))
    return w, h


def fit_by_constraints(
    src_w: float, src_h: float, *, max_w: int, max_h: int
) -> tuple[int, int]:
    """同 :func:`fit_size`，但允许显式放大（用于已知目标框的场景）。"""
    if src_w <= 0 or src_h <= 0:
        return 1, 1
    scale = min(max_w / src_w, max_h / src_h)
    return max(1, int(round(src_w * scale))), max(1, int(round(src_h * scale)))


def pt_to_emu(pt: float) -> int:
    """磅 -> EMU。"""
    return int(round(pt * EMU_PER_PT))


def emu_to_pt(emu: float) -> float:
    """EMU -> 磅。"""
    return emu / EMU_PER_PT


def cm_to_emu(cm: float) -> int:
    """厘米 -> EMU。"""
    return int(round(cm * EMU_PER_CM))


def twip_to_emu(twip: float) -> int:
    """缇（1/20 磅，Excel 列宽单位）-> EMU。"""
    return int(round(twip * EMU_PER_TWIP))
