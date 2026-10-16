# -*- coding: utf-8 -*-
"""转换器注册表：按扩展名分发到具体转换器。

新增一种格式只要写一个带 ``convert(ctx)`` 的类，在这里登记即可。
"""

from __future__ import annotations

from pathlib import Path

from ..logging_setup import get_logger
from .base import ConvertContext, ConvertResult
from .docx import DocxConverter, SUPPORTED_SUFFIXES as DOCX_EXT
from .image import ImageConverter, SUPPORTED_SUFFIXES as IMAGE_EXT
from .legacy import LEGACY_TARGETS, is_legacy
from .pdf import PdfConverter, SUPPORTED_SUFFIXES as PDF_EXT
from .pptx import PptxConverter, SUPPORTED_SUFFIXES as PPTX_EXT
from .text import TextConverter, SUPPORTED_SUFFIXES as TEXT_EXT
from .xlsx import XlsxConverter, SUPPORTED_SUFFIXES as XLSX_EXT

log = get_logger("converters")

#: 扩展名 -> 转换器工厂
_REGISTRY: dict[str, object] = {}


def _register(suffixes, factory) -> None:  # type: ignore[no-untyped-def]
    for s in suffixes:
        _REGISTRY[s.lower()] = factory


def _build_registry() -> None:
    if _REGISTRY:
        return
    _register(IMAGE_EXT, lambda: ImageConverter())
    _register(PDF_EXT, lambda: PdfConverter())
    _register(DOCX_EXT, lambda: DocxConverter())
    _register(XLSX_EXT, lambda: XlsxConverter())
    _register(PPTX_EXT, lambda: PptxConverter(mode="hybrid"))
    _register(TEXT_EXT, lambda: TextConverter())
    _register(LEGACY_TARGETS.keys(), lambda: LegacyConverter())


class LegacyConverter:
    """旧格式的适配器：先转成现代格式，再交给对应转换器。"""

    name = "legacy"
    display = "旧格式（需外部转换器）"

    def convert(self, ctx: ConvertContext) -> ConvertResult:
        from .legacy import convert_legacy

        workdir = ctx.temp_subdir("legacy")
        new_path, channel = convert_legacy(ctx.src, workdir)
        inner = get_converter(new_path)
        if inner is None:
            raise ValueError(f"转换后的格式无法处理: {new_path.suffix}")

        sub_ctx = ConvertContext(
            doc=ctx.doc,
            settings=ctx.settings,
            src=new_path,
            index=ctx.index,
            total=ctx.total,
            timeout=ctx.timeout,
            box=ctx.box,
            landscape_box=ctx.landscape_box,
            temp_dir=ctx.temp_dir,
            progress=ctx.progress,
        )
        result = inner.convert(sub_ctx)
        result.warnings.insert(0, f"旧格式 {ctx.src.suffix} 经 {channel} 转换后合并（源文件未改动）")
        result.converter = f"{result.converter}(via {channel})"
        return result


def get_converter(path: Path):
    """按扩展名取转换器实例；不支持则返回 None。"""
    _build_registry()
    factory = _REGISTRY.get(path.suffix.lower())
    return factory() if factory else None  # type: ignore[operator]


def is_supported(path: Path) -> bool:
    """该文件是否可合并。"""
    _build_registry()
    return path.suffix.lower() in _REGISTRY


def supported_suffixes() -> list[str]:
    """所有支持的扩展名（升序）。"""
    _build_registry()
    return sorted(_REGISTRY.keys())


def supported_files_hint() -> str:
    """给文件对话框/提示用的通配符列表。"""
    exts = supported_suffixes()
    grouped = [
        ("Word 文档", [e for e in exts if e in DOCX_EXT or is_legacy_ext(e)]),
        ("PDF", [e for e in exts if e in PDF_EXT]),
        ("图片", [e for e in exts if e in IMAGE_EXT]),
        ("Excel", [e for e in exts if e in XLSX_EXT]),
        ("PowerPoint", [e for e in exts if e in PPTX_EXT]),
        ("文本", [e for e in exts if e in TEXT_EXT]),
    ]
    parts = []
    for label, items in grouped:
        if items:
            parts.append(f"{label} ({' '.join('*' + e for e in items)})")
    parts.append(f"全部支持的格式 ({' '.join('*' + e for e in exts)})")
    return "|".join(parts)


def is_legacy_ext(suffix: str) -> bool:
    """判断扩展名是否为旧格式。"""
    return suffix.lower() in LEGACY_TARGETS
