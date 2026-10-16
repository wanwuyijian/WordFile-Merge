# -*- coding: utf-8 -*-
"""用户设置（持久化到 %LOCALAPPDATA%\\WordMergeTool\\settings.json）。

设置项刻意保持很少：所有"聪明"的排版决策都由引擎自动做，
这里只放用户真正会想改的东西。
"""

from __future__ import annotations

import json
from dataclasses import asdict, dataclass, field, fields
from pathlib import Path
from typing import Any

from .logging_setup import get_logger
from .paths import settings_file

log = get_logger("settings")


@dataclass
class Settings:
    """可持久化的用户偏好。"""

    #: 是否为每个源文件插入"文档标题"（封面式标题 + 分页）
    include_titles: bool = True
    #: 是否在合并文档开头生成目录（Word 域，打开后按 F9 更新页码）
    include_toc: bool = True
    #: 是否给每个源文件加书签（Word 导航窗格可见）
    include_bookmarks: bool = True
    #: 输出文件名是否自动追加时间戳
    timestamp_output: bool = True
    #: 是否使用 Microsoft Word 引擎增强（本机装了 Word 且用户愿意用时的附加保真路径）
    use_word_engine: bool = True
    #: 上次使用的输入目录
    last_input_dir: str = ""
    #: 上次使用的输出目录
    last_output_dir: str = ""
    #: 窗口几何（"宽x高+X+Y"）
    window_geometry: str = ""
    #: 最近使用过的输入目录（最多 8 个）
    recent_dirs: list[str] = field(default_factory=list)

    # ------------------------------------------------------------------ #
    def to_dict(self) -> dict[str, Any]:
        return asdict(self)

    @classmethod
    def from_dict(cls, data: dict[str, Any]) -> "Settings":
        known = {f.name for f in fields(cls)}
        clean = {k: v for k, v in data.items() if k in known}
        try:
            return cls(**clean)
        except TypeError:
            log.warning("设置文件字段类型异常，回退默认值")
            return cls()

    def remember_dir(self, path: str, *, limit: int = 8) -> None:
        """记录一个最近使用目录（去重、置顶、限长）。"""
        if not path:
            return
        path = str(Path(path))
        items = [p for p in self.recent_dirs if str(p) != path]
        items.insert(0, path)
        self.recent_dirs = items[:limit]


def load_settings() -> Settings:
    """读取设置；文件不存在或损坏时回退默认值（并保留坏文件副本以便排查）。"""
    path = settings_file()
    if not path.is_file():
        return Settings()
    try:
        raw = path.read_text(encoding="utf-8")
        return Settings.from_dict(json.loads(raw))
    except Exception as exc:
        log.error("设置文件读取失败，使用默认设置: %s", exc)
        try:
            path.with_suffix(".json.bad").write_text(path.read_text(encoding="utf-8"), encoding="utf-8")
        except Exception:
            pass
        return Settings()


def save_settings(settings: Settings) -> None:
    """保存设置（原子写：先写临时文件再替换，避免断电写坏）。"""
    path = settings_file()
    try:
        tmp = path.with_suffix(".json.tmp")
        tmp.write_text(
            json.dumps(settings.to_dict(), ensure_ascii=False, indent=2), encoding="utf-8"
        )
        tmp.replace(path)
    except Exception as exc:
        log.error("设置保存失败: %s", exc)
