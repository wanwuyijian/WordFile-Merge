# -*- coding: utf-8 -*-
"""日志与崩溃日志。

三份产物，都落在 ``logs/`` 目录：

``app-YYYYMMDD.log``
    全量运行日志（UTF-8，带毫秒时间戳、级别、线程名）。
``crash-YYYYMMDD-HHMMSS.txt``
    崩溃报告。未捕获异常（主线程 / 子线程 / 顶层 excepthook）都会生成一份，
    内容是完整的异常栈 + 环境信息 + 最近 400 行日志，用户可以直接把这个文件发回来。
``error.log``
    只记 ERROR 及以上的一条流水，方便快速扫一眼"到底出过什么错"。
"""

from __future__ import annotations

import datetime as _dt
import logging
import os
import platform
import sys
import threading
import traceback
from logging.handlers import RotatingFileHandler
from pathlib import Path
from typing import Any, Callable

from . import __app_name__, __version__
from .paths import is_frozen, logs_dir

_LOG_FORMAT = "%(asctime)s.%(msecs)03d [%(levelname)-7s] [%(threadName)s] %(name)s: %(message)s"
_DATE_FORMAT = "%Y-%m-%d %H:%M:%S"

_LOGGER_NAME = "wmtool"
_configured = False
_log_file: Path | None = None
_crash_hook_installed = False

#: GUI 想在日志窗口里同步看到日志时，挂一个回调上来。
log_listeners: list[Callable[[str], None]] = []


class _ListenerHandler(logging.Handler):
    """把日志行转发给 GUI 的回调（回调内部异常必须吞掉，不能反过来搞崩日志）。"""

    def emit(self, record: logging.LogRecord) -> None:  # noqa: D102
        if not log_listeners:
            return
        try:
            line = self.format(record)
        except Exception:  # pragma: no cover - 格式化失败也要保住主流程
            return
        for cb in list(log_listeners):
            try:
                cb(line)
            except Exception:
                pass


def setup_logging(level: int = logging.INFO, *, console: bool = True) -> Path:
    """初始化日志系统，返回主日志文件路径。重复调用是安全的（幂等）。"""
    global _configured, _log_file
    if _configured and _log_file is not None:
        return _log_file

    directory = logs_dir()
    directory.mkdir(parents=True, exist_ok=True)
    stamp = _dt.datetime.now().strftime("%Y%m%d")
    _log_file = directory / f"app-{stamp}.log"

    root = logging.getLogger(_LOGGER_NAME)
    root.setLevel(logging.DEBUG)
    root.propagate = False
    for h in list(root.handlers):
        root.removeHandler(h)

    fmt = logging.Formatter(_LOG_FORMAT, datefmt=_DATE_FORMAT)

    fh = RotatingFileHandler(
        _log_file, maxBytes=5 * 1024 * 1024, backupCount=5, encoding="utf-8"
    )
    fh.setLevel(logging.DEBUG)
    fh.setFormatter(fmt)
    root.addHandler(fh)

    eh = RotatingFileHandler(
        directory / "error.log", maxBytes=2 * 1024 * 1024, backupCount=3, encoding="utf-8"
    )
    eh.setLevel(logging.ERROR)
    eh.setFormatter(fmt)
    root.addHandler(eh)

    lh = _ListenerHandler()
    lh.setLevel(logging.DEBUG)
    lh.setFormatter(fmt)
    root.addHandler(lh)

    if console:
        sh = logging.StreamHandler(sys.stderr)
        sh.setLevel(level)
        sh.setFormatter(fmt)
        root.addHandler(sh)

    _configured = True
    root.info("=" * 78)
    root.info(
        "%s v%s 启动 | python=%s | frozen=%s | pid=%d",
        __app_name__,
        __version__,
        platform.python_version(),
        is_frozen(),
        os.getpid(),
    )
    root.info("日志文件: %s", _log_file)
    return _log_file


def get_logger(name: str = "") -> logging.Logger:
    """取一个子 logger（``wmtool.<name>``）。"""
    if not _configured:
        setup_logging()
    suffix = f".{name}" if name else ""
    return logging.getLogger(f"{_LOGGER_NAME}{suffix}")


def current_log_file() -> Path | None:
    """当前日志文件路径（未初始化时为 None）。"""
    return _log_file


def _environment_report() -> str:
    """崩溃报告里附带的环境信息——排障时这些比栈还关键。"""
    lines = [
        f"程序: {__app_name__} v{__version__}",
        f"时间: {_dt.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
        f"Python: {platform.python_version()} ({platform.architecture()[0]})",
        f"打包运行: {is_frozen()}",
        f"可执行文件: {sys.executable}",
        f"工作目录: {os.getcwd()}",
        f"操作系统: {platform.platform()}",
        f"机器: {platform.node()}",
    ]
    if sys.version_info >= (3, 8):
        lines.append(f"frozen 属性: {getattr(sys, 'frozen', False)}")
    return "\n".join(lines)


def _tail_log(n: int = 400) -> str:
    """取主日志最后 n 行，塞进崩溃报告。"""
    path = _log_file
    if not path or not path.is_file():
        return "(无日志文件)"
    try:
        with path.open("r", encoding="utf-8", errors="replace") as f:
            tail = f.readlines()[-n:]
        return "".join(tail)
    except Exception as exc:  # pragma: no cover
        return f"(读取日志失败: {exc})"


def write_crash_report(exc_type: Any, exc_value: BaseException, exc_tb: Any, *, where: str) -> Path | None:
    """生成崩溃日志文件，返回文件路径。"""
    setup_logging()
    directory = logs_dir()
    directory.mkdir(parents=True, exist_ok=True)
    stamp = _dt.datetime.now().strftime("%Y%m%d-%H%M%S")
    path = directory / f"crash-{stamp}.txt"

    tb_text = "".join(traceback.format_exception(exc_type, exc_value, exc_tb))
    body = "\n".join(
        [
            "=" * 78,
            " 崩溃报告 / CRASH REPORT",
            "=" * 78,
            _environment_report(),
            f"发生位置: {where}",
            f"线程: {threading.current_thread().name}",
            "-" * 78,
            "异常栈:",
            tb_text.rstrip(),
            "-" * 78,
            "最近日志 (末尾 400 行):",
            _tail_log(400).rstrip(),
            "-" * 78,
            f"完整日志请查看: {_log_file}",
            "",
            "如果问题可复现，请把这个文件连同输入文件类型一起反馈。",
            "=" * 78,
        ]
    )
    try:
        path.write_text(body, encoding="utf-8")
    except Exception:
        return None

    log = logging.getLogger(f"{_LOGGER_NAME}.crash")
    log.error("已写出崩溃报告: %s", path)
    try:
        # 崩溃报告也复制一份到 stderr，命令行运行时不至于什么都看不到
        print(body, file=sys.stderr)
    except Exception:
        pass
    return path


def install_crash_handlers(*, reraise: bool = False) -> None:
    """安装全局崩溃钩子：主线程、子线程、以及 tkinter 回调里的异常。"""
    global _crash_hook_installed
    setup_logging()

    if not _crash_hook_installed:
        previous_hook = sys.excepthook

        def _hook(exc_type, exc_value, exc_tb):  # type: ignore[no-untyped-def]
            if issubclass(exc_type, KeyboardInterrupt):
                previous_hook(exc_type, exc_value, exc_tb)
                return
            write_crash_report(exc_type, exc_value, exc_tb, where="主线程未捕获异常")
            if reraise:
                previous_hook(exc_type, exc_value, exc_tb)

        sys.excepthook = _hook

        def _thread_hook(args):  # type: ignore[no-untyped-def]
            if issubclass(args.exc_type, SystemExit):
                return
            write_crash_report(
                args.exc_type,
                args.exc_value,
                args.exc_traceback,
                where=f"子线程未捕获异常: {args.thread.name if args.thread else '?'}",
            )

        threading.excepthook = _thread_hook  # type: ignore[assignment]
        _crash_hook_installed = True

    _install_tk_callback_hook()


def _install_tk_callback_hook() -> None:
    """tkinter 的 Callback 异常不会走 excepthook，得挂钩 report_callback_exception。"""
    try:
        import tkinter  # noqa: F401
    except Exception:
        return
    try:
        import tkinter as tk

        original = tk.Tk.report_callback_exception

        def _patched(self, exc, val, tb):  # type: ignore[no-untyped-def]
            write_crash_report(exc, val, tb, where="Tk 回调异常（界面事件处理）")
            try:
                original(self, exc, val, tb)
            except Exception:
                pass

        tk.Tk.report_callback_exception = _patched  # type: ignore[assignment]
    except Exception:
        pass


class CrashReporter:
    """上下文管理器：把一段逻辑里的异常记进崩溃日志，然后按需吞掉或重抛。

    用在"单个文件转换失败不能拖垮整批任务"这种地方。
    """

    def __init__(self, where: str, logger: logging.Logger | None = None, *, swallow: bool = True):
        self.where = where
        self.logger = logger or get_logger("crash")
        self.swallow = swallow

    def __enter__(self) -> "CrashReporter":
        return self

    def __exit__(self, exc_type, exc_value, exc_tb) -> bool:
        if exc_type is None:
            return False
        if issubclass(exc_type, KeyboardInterrupt):
            return False
        write_crash_report(exc_type, exc_value, exc_tb, where=self.where)
        if self.swallow:
            self.logger.error("【已记录崩溃日志】%s: %s", self.where, exc_value)
            return True
        return False


def register_listener(cb: Callable[[str], None]) -> None:
    """注册日志监听（GUI 日志窗口）。"""
    if cb not in log_listeners:
        log_listeners.append(cb)


def unregister_listener(cb: Callable[[str], None]) -> None:
    """注销日志监听。"""
    if cb in log_listeners:
        log_listeners.remove(cb)
