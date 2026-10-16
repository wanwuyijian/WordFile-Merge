# -*- coding: utf-8 -*-
"""路径与运行模式判定。

程序有两种运行形态：

* **源码模式**：直接 ``python -m wmtool``，日志写在项目目录下。
* **打包模式**：PyInstaller 打出的单文件 exe，日志写在 exe 同级的 ``logs`` 目录，
  这样用户把 exe 拷到 U 盘/别的电脑上，日志是跟着 exe 走的，方便排障。
"""

from __future__ import annotations

import os
import sys
from pathlib import Path


def is_frozen() -> bool:
    """是否运行在 PyInstaller 打出来的 exe 里。"""
    return bool(getattr(sys, "frozen", False))


def app_dir() -> Path:
    """程序所在的"家"目录：打包后是 exe 同级目录，源码模式是项目根目录。"""
    if is_frozen():
        return Path(sys.executable).resolve().parent
    # src/wmtool/paths.py -> 上溯三级到项目根
    return Path(__file__).resolve().parents[2]


def user_data_dir() -> Path:
    """用户级可写目录（%LOCALAPPDATA%\\WordMergeTool）。

    设置文件写这里，避免程序装在只读目录（如 Program Files）时写不进去。
    """
    base = os.environ.get("LOCALAPPDATA") or os.environ.get("APPDATA")
    if base:
        return Path(base) / "WordMergeTool"
    return Path.home() / ".word-merge-tool"


def logs_dir() -> Path:
    """日志目录。按优先级选：

    1. 环境变量 ``WORDMERGE_LOG_DIR``（用户显式指定，最优先）
    2. exe / 项目同级的 ``logs/``（**便携**：拷走 exe，日志跟着走）
    3. ``%LOCALAPPDATA%\\WordMergeTool\\logs``（上面写不进去时的退路，
       比如 exe 被放在只读的 Program Files 下）
    """
    override = os.environ.get("WORDMERGE_LOG_DIR")
    if override:
        d = Path(override)
        d.mkdir(parents=True, exist_ok=True)
        return d

    primary = app_dir() / "logs"
    try:
        primary.mkdir(parents=True, exist_ok=True)
        probe = primary / ".write-probe"
        probe.write_text("ok", encoding="utf-8")
        probe.unlink()
        return primary
    except OSError:
        fallback = user_data_dir() / "logs"
        fallback.mkdir(parents=True, exist_ok=True)
        return fallback


def settings_file() -> Path:
    """用户设置文件路径。"""
    d = user_data_dir()
    d.mkdir(parents=True, exist_ok=True)
    return d / "settings.json"


def resource_path(rel: str) -> Path:
    """取打包进 exe 的静态资源（PyInstaller 解包到 sys._MEIPASS）。"""
    base = getattr(sys, "_MEIPASS", None)
    if base:
        return Path(base) / rel
    return app_dir() / rel
