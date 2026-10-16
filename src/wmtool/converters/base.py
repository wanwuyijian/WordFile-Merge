# -*- coding: utf-8 -*-
"""转换器公共契约。

每个转换器接收：

* ``src``：源文件路径（**只读**，绝不写回）
* ``ctx``：:class:`ConvertContext`，携带输出文档、页面框、设置、工作区

并把自己的内容**追加**到 ``ctx.doc`` 末尾，返回 :class:`ConvertResult`
（含内容概要、是否横向、警告列表）。

铁律：转换器只能往 ``ctx.doc`` 写，**不得触碰 ``src``**。工作区里的临时文件
用 ``ctx.temp_file()`` 取，退出时由引擎统一清理。
"""

from __future__ import annotations

import tempfile
from dataclasses import dataclass, field
from pathlib import Path
from typing import TYPE_CHECKING

from docx.document import Document as DocxDocument

from ..geometry import LANDSCAPE_BOX, PORTRAIT_BOX, PageBox

if TYPE_CHECKING:  # pragma: no cover
    from ..settings import Settings


@dataclass
class ConvertResult:
    """单个文件转换后的自述。"""

    converter: str = ""
    #: 内容概要，写进文档目录/封面信息用
    summary: str = ""
    #: 建议的页面方向（转换器发现自己内容很宽时置 True）
    landscape: bool = False
    #: 实际写入的图片张数
    image_count: int = 0
    #: 实际写入的表格数
    table_count: int = 0
    #: 非致命警告，会写进日志和最终报告
    warnings: list[str] = field(default_factory=list)

    def warn(self, message: str) -> None:
        self.warnings.append(message)


@dataclass
class ConvertContext:
    """一次转换所需的全部上下文。"""

    doc: DocxDocument
    settings: "Settings"
    #: 当前源文件路径
    src: Path
    #: 当前源文件在任务里的显示序号（从 1 开始）
    index: int
    #: 当前源文件总数
    total: int
    #: 单文件时间预算（秒），<=0 表示不限
    timeout: float = 120.0
    #: 本次转换的页面框（纵向 A4）
    box: PageBox = PORTRAIT_BOX
    #: 横向页面框
    landscape_box: PageBox = LANDSCAPE_BOX
    #: 临时目录（由引擎创建与清理）
    temp_dir: Path | None = None
    #: 进度回调 (0.0~1.0, 说明文字)
    progress: object = None

    def temp_file(self, suffix: str = "") -> Path:
        """在工作区里取一个临时文件路径。"""
        base = self.temp_dir if self.temp_dir else Path(tempfile.gettempdir())
        base.mkdir(parents=True, exist_ok=True)
        fd, name = tempfile.mkstemp(suffix=suffix, dir=str(base))
        import os

        os.close(fd)
        return Path(name)

    def temp_subdir(self, name: str) -> Path:
        """在工作区里取一个临时子目录。"""
        base = self.temp_dir if self.temp_dir else Path(tempfile.gettempdir())
        d = base / name
        d.mkdir(parents=True, exist_ok=True)
        return d

    def report(self, fraction: float, message: str = "") -> None:
        """向引擎汇报进度（回调异常不得影响转换）。"""
        if self.progress is None:
            return
        try:
            self.progress(fraction, message)  # type: ignore[operator]
        except Exception:
            pass
