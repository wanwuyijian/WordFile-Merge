# -*- coding: utf-8 -*-
"""图形界面：选文件、拖拽排序、合并、看日志。

界面分三块（用 PanedWindow 分割，比例可调）：

* **上**：缩略图网格——待合并文件一览，**可以拖动缩略图调整顺序**，
  也支持从资源管理器直接拖文件/文件夹进来
* **右/下**：日志窗口 + 进度条 + 状态栏
* **底部**：设置勾选项与「开始合并」按钮

所有耗时工作都跑在后台线程，界面用队列 + ``after()`` 轮询更新——
tkinter 不是线程安全的，绝不在子线程里碰控件。
"""

from __future__ import annotations

import datetime as _dt
import os
import queue
import threading
import tkinter as tk
import traceback
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from PIL import Image, ImageDraw, ImageTk

from . import __app_name_cn__, __version__
from .converters import get_converter, supported_files_hint, supported_suffixes
from .engine import MergeEngine, MergeReport
from .logging_setup import (
    get_logger,
    install_crash_handlers,
    logs_dir,
    register_listener,
    setup_logging,
    unregister_listener,
    write_crash_report,
)
from .paths import app_dir
from .settings import Settings, load_settings, save_settings

log = get_logger("gui")

# 可选拖拽支持：没有 tkinterdnd2 也能用，只是不能拖文件进来
try:
    from tkinterdnd2 import DND_FILES, TkinterDnD

    HAS_DND = True
except Exception:  # pragma: no cover
    HAS_DND = False
    DND_FILES = None
    TkinterDnD = None

#: 缩略图尺寸（像素）
THUMB_W, THUMB_H = 132, 96
GRID_COLUMNS = 5

#: 颜色
COLOR_BG = "#F4F6F9"
COLOR_CARD = "#FFFFFF"
COLOR_SEL = "#2F6FED"
COLOR_TEXT = "#1F2733"
COLOR_MUTED = "#7A8699"


def _human_size(n: int) -> str:
    """字节数 -> 人类可读。"""
    size = float(max(0, n))
    for unit in ("B", "KB", "MB", "GB"):
        if size < 1024 or unit == "GB":
            return f"{size:.0f} {unit}" if unit == "B" else f"{size:.1f} {unit}"
        size /= 1024
    return f"{n} B"


# --------------------------------------------------------------------------- #
# 缩略图卡片
# --------------------------------------------------------------------------- #
class FileCard(tk.Frame):
    """一个文件卡片：缩略图 + 文件名 + 类型徽标。

    自带鼠标事件，配合 :class:`CardGrid` 实现拖拽排序。
    """

    def __init__(self, master, item: "FileItem", on_select, on_drag_start, on_drag_motion, on_drag_drop):
        super().__init__(master, bg=COLOR_CARD, highlightthickness=2,
                         highlightbackground="#DDE3EC", bd=0)
        self.item = item
        self.on_select = on_select
        self._on_drag_start = on_drag_start
        self._on_drag_motion = on_drag_motion
        self._on_drag_drop = on_drag_drop
        self._photo = None
        self.selected = False

        self.thumb = tk.Label(self, bg=COLOR_CARD, width=THUMB_W, height=THUMB_H)
        self.thumb.pack(padx=4, pady=(4, 2))

        name = item.path.name
        if len(name) > 20:
            name = name[:9] + "…" + name[-9:]
        self.name_label = tk.Label(
            self, text=name, bg=COLOR_CARD, fg=COLOR_TEXT,
            font=("微软雅黑", 8), wraplength=THUMB_W,
        )
        self.name_label.pack(padx=3)

        self.badge = tk.Label(
            self, text=item.kind_label, bg="#EAF0FB", fg="#3B5BA9",
            font=("微软雅黑", 7), padx=5, pady=1,
        )
        self.badge.pack(pady=(2, 4))

        for w in (self, self.thumb, self.name_label, self.badge):
            w.bind("<Button-1>", self._click)
            w.bind("<B1-Motion>", self._motion)
            w.bind("<ButtonRelease-1>", self._release)
            # 右键菜单里也可以用方向键移动顺序
            w.bind("<Button-3>", self._context)

    # ------------------------------------------------------------------ #
    def set_image(self, photo) -> None:  # type: ignore[no-untyped-def]
        self._photo = photo
        self.thumb.configure(image=photo, width=0, height=0)

    def set_selected(self, value: bool) -> None:
        self.selected = value
        self.configure(
            highlightbackground=COLOR_SEL if value else "#DDE3EC",
            bg="#F7FAFF" if value else COLOR_CARD,
        )

    def _click(self, event) -> None:  # type: ignore[no-untyped-def]
        self.on_select(self)
        self._on_drag_start(self, event)

    def _motion(self, event) -> None:  # type: ignore[no-untyped-def]
        self._on_drag_motion(self, event)

    def _release(self, event) -> None:  # type: ignore[no-untyped-def]
        self._on_drag_drop(self, event)

    def _context(self, event) -> None:  # type: ignore[no-untyped-def]
        self.on_select(self)
        menu = tk.Menu(self, tearoff=0)
        menu.add_command(label="上移", command=lambda: self._move(-1))
        menu.add_command(label="下移", command=lambda: self._move(1))
        menu.add_separator()
        menu.add_command(label="移除", command=lambda: self._grid_remove())
        menu.add_command(label="打开所在文件夹", command=self._open_folder)
        try:
            menu.tk_popup(event.x_root, event.y_root)
        finally:
            menu.grab_release()

    def _move(self, delta: int) -> None:
        grid = getattr(self.master, "owner", None)
        if grid is not None:
            grid.move(self, delta)

    def _grid_remove(self) -> None:
        grid = getattr(self.master, "owner", None)
        if grid is not None:
            grid.remove(self)

    def _open_folder(self) -> None:
        try:
            os.startfile(str(self.item.path.parent))  # noqa: S606
        except Exception as exc:
            log.warning("打开文件夹失败: %s", exc)


# --------------------------------------------------------------------------- #
# 文件条目
# --------------------------------------------------------------------------- #
class FileItem:
    """一个待合并文件。"""

    def __init__(self, path: Path):
        self.path = Path(path)
        self.thumb_image = None  # PIL 图或 None
        self.selected = False

    @property
    def kind_label(self) -> str:
        """类型徽标文字。"""
        conv = get_converter(self.path)
        if conv is None:
            return self.path.suffix.lstrip(".").upper() or "未知"
        return getattr(conv, "display", self.path.suffix.lstrip(".").upper())

    @property
    def size(self) -> int:
        try:
            return self.path.stat().st_size
        except OSError:
            return 0

    def make_thumbnail(self) -> Image.Image:
        """生成缩略图（PDF/PPT 取首页，图片取自身，其它画一个类型图标）。"""
        w, h = THUMB_W, THUMB_H
        suffix = self.path.suffix.lower()
        try:
            if suffix in (".png", ".jpg", ".jpeg", ".bmp", ".gif", ".webp", ".tif", ".tiff", ".ico"):
                with Image.open(self.path) as im:
                    im = im.convert("RGB")
                    im.thumbnail((w, h), Image.LANCZOS)
                    canvas = Image.new("RGB", (w, h), "#FFFFFF")
                    canvas.paste(im, ((w - im.width) // 2, (h - im.height) // 2))
                    return canvas

            if suffix == ".pdf":
                import pymupdf

                with pymupdf.open(str(self.path)) as doc:
                    if doc.page_count:
                        pix = doc.load_page(0).get_pixmap(dpi=48)
                        im = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
                        im.thumbnail((w, h), Image.LANCZOS)
                        canvas = Image.new("RGB", (w, h), "#FFFFFF")
                        canvas.paste(im, ((w - im.width) // 2, (h - im.height) // 2))
                        return canvas

            if suffix in (".pptx", ".pptm"):
                # 不在预览阶段启动 PowerPoint（太慢），只画图标
                return self._placeholder("PPT")
        except Exception as exc:
            log.debug("缩略图生成失败 %s: %s", self.path.name, exc)

        return self._placeholder(self.kind_label[:4].upper())

    def _placeholder(self, text: str) -> Image.Image:
        """画一个带文字的占位图。"""
        img = Image.new("RGB", (THUMB_W, THUMB_H), "#EEF2F8")
        d = ImageDraw.Draw(img)
        d.rectangle([1, 1, THUMB_W - 2, THUMB_H - 2], outline="#C9D4E4", width=1)
        tw = d.textlength(text)
        d.text(((THUMB_W - tw) / 2, THUMB_H / 2 - 6), text, fill="#5A6B85")
        return img


# --------------------------------------------------------------------------- #
# 卡片网格（拖拽排序）
# --------------------------------------------------------------------------- #
class CardGrid(tk.Frame):
    """缩略图网格，支持拖动卡片重排顺序。

    实现方式是"拖动时实时重排"：鼠标按下并移动超过阈值后，
    每次移动都算出指针下方的格子，直接把卡片换到那个位置。
    比画一个浮动残影简单得多，观感也够用。
    """

    def __init__(self, master, on_change=None, on_select=None):
        super().__init__(master, bg=COLOR_BG)
        self.owner = self
        self.on_change = on_change
        self.on_select_cb = on_select
        self.cards: list[FileCard] = []
        self.items: list[FileItem] = []
        self._drag_card: FileCard | None = None
        self._drag_started = False
        self._press_xy = (0, 0)

        self.canvas = tk.Canvas(self, bg=COLOR_BG, highlightthickness=0)
        self.scroll = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=self.scroll.set)
        self.canvas.pack(side="left", fill="both", expand=True)
        self.scroll.pack(side="right", fill="y")

        self.inner = tk.Frame(self.canvas, bg=COLOR_BG)
        self._win = self.canvas.create_window((0, 0), window=self.inner, anchor="nw")

        self.inner.bind("<Configure>", self._on_inner_configure)
        self.canvas.bind("<Configure>", self._on_canvas_configure)
        # 鼠标滚轮
        self.canvas.bind_all("<MouseWheel>", self._on_wheel)

        self._placeholder = tk.Label(
            self.inner, text="把文件拖到这里，或点上方「添加文件」",
            bg=COLOR_BG, fg=COLOR_MUTED, font=("微软雅黑", 11),
        )
        self._placeholder.grid(row=0, column=0, padx=24, pady=48, sticky="w")

    # ------------------------------------------------------------------ #
    def _on_inner_configure(self, _event=None) -> None:
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

    def _on_canvas_configure(self, event) -> None:  # type: ignore[no-untyped-def]
        self.canvas.itemconfigure(self._win, width=event.width)

    def _on_wheel(self, event) -> None:  # type: ignore[no-untyped-def]
        try:
            if self.canvas.winfo_exists():
                self.canvas.yview_scroll(int(-event.delta / 120), "units")
        except Exception:
            pass

    # ------------------------------------------------------------------ #
    def add_paths(self, paths: list[Path]) -> int:
        """添加文件；目录会被展开。返回实际新增数量。"""
        added: list[FileItem] = []
        existing = {str(it.path.resolve()) for it in self.items}
        for p in paths:
            p = Path(p)
            if p.is_dir():
                for f in sorted(p.rglob("*")):
                    if f.is_file() and self._acceptable(f) and str(f.resolve()) not in existing:
                        existing.add(str(f.resolve()))
                        added.append(FileItem(f))
            elif p.is_file() and self._acceptable(p):
                if str(p.resolve()) not in existing:
                    existing.add(str(p.resolve()))
                    added.append(FileItem(p))

        for item in added:
            self.items.append(item)
            card = FileCard(
                self.inner, item, self._select, self._drag_start,
                self._drag_motion, self._drag_drop,
            )
            self.cards.append(card)
            try:
                item.thumb_image = item.make_thumbnail()
                photo = ImageTk.PhotoImage(item.thumb_image)
                card.set_image(photo)
                card._photo_ref = photo  # 保持引用，否则图会被回收
            except Exception as exc:
                log.debug("缩略图显示失败: %s", exc)

        if added:
            self._relayout()
            self._notify()
        return len(added)

    @staticmethod
    def _acceptable(path: Path) -> bool:
        """是否是可合并的文件（排除临时文件与隐藏文件）。"""
        name = path.name
        if name.startswith("~$") or name.startswith("."):
            return False
        return path.suffix.lower() in set(supported_suffixes())

    def remove_selected(self) -> None:
        """移除选中的卡片。"""
        victims = [c for c in self.cards if c.selected]
        for card in victims:
            self._detach(card)
        if victims:
            self._relayout()
            self._notify()

    def clear(self) -> None:
        """清空。"""
        for card in list(self.cards):
            self._detach(card)
        self._relayout()
        self._notify()

    def _detach(self, card: FileCard) -> None:
        if card in self.cards:
            idx = self.cards.index(card)
            self.cards.pop(idx)
            self.items.pop(idx)
        card.destroy()

    def move(self, card: FileCard, delta: int) -> None:
        """把卡片按相对位置移动（右键菜单用）。"""
        if card not in self.cards:
            return
        i = self.cards.index(card)
        j = max(0, min(len(self.cards) - 1, i + delta))
        if i == j:
            return
        self.cards.insert(j, self.cards.pop(i))
        self.items.insert(j, self.items.pop(i))
        self._relayout()
        self._notify()

    def ordered_paths(self) -> list[Path]:
        """当前顺序下的文件路径。"""
        return [it.path for it in self.items]

    # ------------------------------------------------------------------ #
    def _select(self, card: FileCard) -> None:
        """点选卡片。按住 Ctrl 可多选，否则单选。"""
        ctrl = _ctrl_pressed()
        if not ctrl:
            for c in self.cards:
                c.set_selected(False)
            card.set_selected(True)
        else:
            card.set_selected(not card.selected)
        if self.on_select_cb:
            self.on_select_cb()

    def _drag_start(self, card: FileCard, event) -> None:  # type: ignore[no-untyped-def]
        self._drag_card = card
        self._drag_started = False
        self._press_xy = (event.x_root, event.y_root)

    def _drag_motion(self, card: FileCard, event) -> None:  # type: ignore[no-untyped-def]
        if self._drag_card is None or card is not self._drag_card:
            return
        dx = abs(event.x_root - self._press_xy[0])
        dy = abs(event.y_root - self._press_xy[1])
        if not self._drag_started and dx + dy < 8:
            return  # 还没到拖动阈值，可能只是点选
        self._drag_started = True
        self.configure(cursor="fleur")

        # 指针在窗口里的位置 -> 目标格子
        x = event.x_root - self.inner.winfo_rootx()
        y = event.y_root - self.inner.winfo_rooty()
        if not self.cards:
            return
        cw = self.cards[0].winfo_reqwidth() + 8
        ch = self.cards[0].winfo_reqheight() + 8
        col = max(0, min(self._columns() - 1, x // max(1, cw)))
        row = max(0, y // max(1, ch))
        target = row * self._columns() + col
        target = max(0, min(len(self.cards) - 1, target))
        cur = self.cards.index(card)
        if target != cur:
            self.cards.insert(target, self.cards.pop(cur))
            self.items.insert(target, self.items.pop(cur))
            self._relayout()

    def _drag_drop(self, card: FileCard, event) -> None:  # type: ignore[no-untyped-def]
        if self._drag_started:
            self.configure(cursor="")
            self._notify()
        self._drag_card = None
        self._drag_started = False

    def _columns(self) -> int:
        """按容器宽度算列数。"""
        try:
            width = self.canvas.winfo_width()
            card_w = THUMB_W + 34
            return max(1, width // card_w)
        except Exception:
            return GRID_COLUMNS

    def _relayout(self) -> None:
        """重排所有卡片，并显示/隐藏空提示。"""
        if not self.cards:
            self._placeholder.grid(row=0, column=0, padx=24, pady=48, sticky="w")
            self._on_inner_configure()
            return
        self._placeholder.grid_remove()
        cols = self._columns()
        for i, card in enumerate(self.cards):
            card.grid(
                row=i // cols, column=i % cols,
                padx=6, pady=6, sticky="n",
            )
        self._on_inner_configure()

    def _notify(self) -> None:
        if self.on_change:
            self.on_change()


def _ctrl_pressed() -> bool:
    """是否按着 Ctrl（用于多选）。"""
    try:
        import ctypes

        return bool(ctypes.windll.user32.GetAsyncKeyState(0x11) & 0x8000)
    except Exception:
        return False


# --------------------------------------------------------------------------- #
# 主窗口
# --------------------------------------------------------------------------- #
class App:
    """主窗口。"""

    def __init__(self) -> None:
        setup_logging()
        install_crash_handlers()
        self.settings: Settings = load_settings()
        self.report: MergeReport | None = None
        self.worker: threading.Thread | None = None
        self.engine: MergeEngine | None = None
        self.ui_queue: queue.Queue = queue.Queue()
        self._closing = False

        self.root = TkinterDnD.Tk() if HAS_DND else tk.Tk()
        self.root.title(f"{__app_name_cn__} v{__version__}")
        self.root.geometry(self.settings.window_geometry or "1120x760")
        self.root.minsize(900, 620)
        self.root.configure(bg=COLOR_BG)

        try:
            icon = Path(app_dir()) / "assets" / "app.ico"
            if icon.is_file():
                self.root.iconbitmap(str(icon))
        except Exception:
            pass

        self._build_style()
        self._build_widgets()
        self._setup_dnd()
        self._setup_dpi()

        register_listener(self._enqueue_log)
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)
        self.root.after(120, self._pump)
        log.info("界面已就绪（拖拽支持: %s）", "可用" if HAS_DND else "不可用")

    # ------------------------------------------------------------------ #
    def _build_style(self) -> None:
        style = ttk.Style()
        try:
            style.theme_use("vista")
        except Exception:
            pass
        style.configure("TButton", font=("微软雅黑", 9))
        style.configure("Accent.TButton", font=("微软雅黑", 10, "bold"))
        style.configure("TCheckbutton", background=COLOR_BG, font=("微软雅黑", 9))
        style.configure("TLabel", background=COLOR_BG, font=("微软雅黑", 9))
        style.configure("TFrame", background=COLOR_BG)
        style.configure("Treeview", font=("Consolas", 9), rowheight=20)
        style.configure("TLabelframe", background=COLOR_BG)
        style.configure("TLabelframe.Label", background=COLOR_BG, font=("微软雅黑", 9, "bold"))

    def _setup_dpi(self) -> None:
        """在高分屏上开启 DPI 感知，避免界面糊成一片。"""
        if os.name != "nt":
            return
        try:
            import ctypes

            ctypes.windll.shcore.SetProcessDpiAwareness(1)
        except Exception:
            pass

    def _build_widgets(self) -> None:
        # ---------------- 顶部工具条 ----------------
        toolbar = tk.Frame(self.root, bg=COLOR_BG)
        toolbar.pack(fill="x", padx=12, pady=(10, 6))

        title = tk.Label(
            toolbar, text=f"{__app_name_cn__}", bg=COLOR_BG, fg=COLOR_TEXT,
            font=("微软雅黑", 13, "bold"),
        )
        title.pack(side="left")
        tk.Label(
            toolbar, text=f"v{__version__}　把 Word / PPT / Excel / PDF / 图片合成一个 Word",
            bg=COLOR_BG, fg=COLOR_MUTED, font=("微软雅黑", 9),
        ).pack(side="left", padx=(10, 0))

        btns = tk.Frame(toolbar, bg=COLOR_BG)
        btns.pack(side="right")
        ttk.Button(btns, text="添加文件", command=self.add_files).pack(side="left", padx=3)
        ttk.Button(btns, text="添加文件夹", command=self.add_folder).pack(side="left", padx=3)
        ttk.Button(btns, text="移除选中", command=self.remove_selected).pack(side="left", padx=3)
        ttk.Button(btns, text="清空", command=self.clear_all).pack(side="left", padx=3)

        # ---------------- 中部：卡片区 + 日志区 ----------------
        paned = tk.PanedWindow(self.root, orient="vertical", bg=COLOR_BG,
                               sashwidth=6, sashrelief="flat")
        paned.pack(fill="both", expand=True, padx=12, pady=4)

        files_frame = ttk.Labelframe(paned, text=" 待合并文件（拖动缩略图可调整顺序） ")
        self.grid = CardGrid(files_frame, on_change=self._on_files_changed,
                             on_select=self._update_status)
        self.grid.pack(fill="both", expand=True, padx=6, pady=6)
        paned.add(files_frame, minsize=220, stretch="always")

        log_frame = ttk.Labelframe(paned, text=" 运行日志 ")
        self.log_text = tk.Text(
            log_frame, height=12, wrap="none", bg="#1E2430", fg="#D6DEE8",
            insertbackground="#D6DEE8", font=("Consolas", 9), bd=0,
            state="disabled",
        )
        log_scroll_y = ttk.Scrollbar(log_frame, orient="vertical", command=self.log_text.yview)
        log_scroll_x = ttk.Scrollbar(log_frame, orient="horizontal", command=self.log_text.xview)
        self.log_text.configure(yscrollcommand=log_scroll_y.set, xscrollcommand=log_scroll_x.set)
        self.log_text.grid(row=0, column=0, sticky="nsew")
        log_scroll_y.grid(row=0, column=1, sticky="ns")
        log_scroll_x.grid(row=1, column=0, sticky="ew")
        log_frame.rowconfigure(0, weight=1)
        log_frame.columnconfigure(0, weight=1)
        self.log_text.tag_configure("error", foreground="#FF9A9A")
        self.log_text.tag_configure("warning", foreground="#FFD479")
        self.log_text.tag_configure("info", foreground="#D6DEE8")
        paned.add(log_frame, minsize=140, stretch="always")

        # ---------------- 底部：选项 + 进度 + 按钮 ----------------
        bottom = tk.Frame(self.root, bg=COLOR_BG)
        bottom.pack(fill="x", padx=12, pady=(4, 12))

        opts = tk.Frame(bottom, bg=COLOR_BG)
        opts.pack(fill="x")
        self.var_titles = tk.BooleanVar(value=self.settings.include_titles)
        self.var_toc = tk.BooleanVar(value=self.settings.include_toc)
        self.var_bookmarks = tk.BooleanVar(value=self.settings.include_bookmarks)
        for text, var in (
            ("每个文件加标题", self.var_titles),
            ("生成目录（Word 里按 F9 更新页码）", self.var_toc),
            ("加书签（导航窗格可跳转）", self.var_bookmarks),
        ):
            ttk.Checkbutton(opts, text=text, variable=var).pack(side="left", padx=(0, 14))
        ttk.Button(opts, text="打开日志目录", command=self.open_logs).pack(side="right")

        progress_row = tk.Frame(bottom, bg=COLOR_BG)
        progress_row.pack(fill="x", pady=(8, 0))
        self.progress = ttk.Progressbar(progress_row, mode="determinate", maximum=100)
        self.progress.pack(side="left", fill="x", expand=True)
        self.btn_merge = ttk.Button(
            progress_row, text="开始合并", style="Accent.TButton", command=self.start_merge
        )
        self.btn_merge.pack(side="left", padx=(10, 0))
        self.btn_cancel = ttk.Button(
            progress_row, text="取消", command=self.cancel_merge, state="disabled"
        )
        self.btn_cancel.pack(side="left", padx=(6, 0))

        self.status = tk.Label(
            bottom, text="就绪", bg=COLOR_BG, fg=COLOR_MUTED,
            font=("微软雅黑", 9), anchor="w",
        )
        self.status.pack(fill="x", pady=(6, 0))

    def _setup_dnd(self) -> None:
        """注册拖放（文件与文件夹都可以拖进来）。"""
        if not HAS_DND:
            return
        try:
            for widget in (self.grid.canvas, self.root):
                widget.drop_target_register(DND_FILES)
                widget.dnd_bind("<<Drop>>", self._on_drop)
            log.info("已启用拖放：可直接把文件/文件夹拖进窗口")
        except Exception as exc:
            log.warning("拖放注册失败（不影响其它功能）: %s", exc)

    def _on_drop(self, event) -> None:  # type: ignore[no-untyped-def]
        """处理拖入的文件。"""
        try:
            raw = self.root.tk.splitlist(event.data)
        except Exception:
            raw = str(event.data).split()
        paths = [Path(p) for p in raw]
        n = self.grid.add_paths(paths)
        if n:
            log.info("拖入 %d 个文件", n)
        else:
            log.info("拖入的内容里没有可合并的文件")
            self.status.configure(text="拖入的内容里没有可合并的文件", fg="#C0392B")

    # ------------------------------------------------------------------ #
    # 按钮动作
    # ------------------------------------------------------------------ #
    def add_files(self) -> None:
        files = filedialog.askopenfilenames(
            title="选择要合并的文件",
            initialdir=self.settings.last_input_dir or str(Path.home()),
            filetypes=[
                ("Word 文档", "*.docx *.docm *.doc"),
                ("PDF", "*.pdf"),
                ("图片", "*.jpg *.jpeg *.png *.bmp *.gif *.webp *.tif *.tiff"),
                ("Excel", "*.xlsx *.xlsm *.csv *.tsv *.xls"),
                ("PowerPoint", "*.pptx *.pptm *.ppt"),
                ("文本", "*.txt *.md *.log *.json *.xml *.html"),
                ("全部支持的格式", " ".join("*" + s for s in supported_suffixes())),
                ("所有文件", "*.*"),
            ],
        )
        if files:
            n = self.grid.add_paths([Path(f) for f in files])
            self.settings.last_input_dir = str(Path(files[0]).parent)
            log.info("添加 %d 个文件", n)

    def add_folder(self) -> None:
        folder = filedialog.askdirectory(
            title="选择文件夹（会自动找出其中支持的文件）",
            initialdir=self.settings.last_input_dir or str(Path.home()),
        )
        if not folder:
            return
        n = self.grid.add_paths([Path(folder)])
        self.settings.remember_dir(folder)
        self.settings.last_input_dir = folder
        if n:
            log.info("从文件夹添加 %d 个文件: %s", n, folder)
        else:
            messagebox.showinfo("没有可合并的文件", "该文件夹里没有找到支持的格式。")

    def remove_selected(self) -> None:
        self.grid.remove_selected()

    def clear_all(self) -> None:
        if self.grid.cards and messagebox.askyesno("确认", "清空全部待合并文件？"):
            self.grid.clear()

    def open_logs(self) -> None:
        try:
            os.startfile(str(logs_dir()))  # noqa: S606
        except Exception as exc:
            messagebox.showerror("打不开", f"无法打开日志目录：{exc}")

    def _on_files_changed(self) -> None:
        self._update_status()

    def _update_status(self) -> None:
        n = len(self.grid.cards)
        total = sum(it.size for it in self.grid.items)
        sel = sum(1 for c in self.grid.cards if c.selected)
        text = f"共 {n} 个文件，合计 {_human_size(total)}"
        if sel:
            text += f"　|　已选中 {sel} 个"
        if n == 0:
            text = "就绪：请添加文件，或直接把文件拖进窗口"
        self.status.configure(text=text, fg=COLOR_MUTED)

    # ------------------------------------------------------------------ #
    # 合并
    # ------------------------------------------------------------------ #
    def start_merge(self) -> None:
        """校验输入并启动后台合并。"""
        if self.worker and self.worker.is_alive():
            return
        paths = self.grid.ordered_paths()
        if not paths:
            messagebox.showwarning("还没有文件", "请先添加要合并的文件。")
            return

        default_name = f"合并文档-{_dt.datetime.now():%Y%m%d-%H%M%S}.docx"
        init_dir = self.settings.last_output_dir or self.settings.last_input_dir or str(Path.home())
        out = filedialog.asksaveasfilename(
            title="保存合并结果",
            initialdir=init_dir,
            initialfile=default_name,
            defaultextension=".docx",
            filetypes=[("Word 文档", "*.docx")],
        )
        if not out:
            return
        output = Path(out)
        self.settings.last_output_dir = str(output.parent)
        self.settings.include_titles = self.var_titles.get()
        self.settings.include_toc = self.var_toc.get()
        self.settings.include_bookmarks = self.var_bookmarks.get()
        save_settings(self.settings)

        self.btn_merge.configure(state="disabled")
        self.btn_cancel.configure(state="normal")
        self.progress.configure(value=0)
        log.info("=" * 60)
        log.info("开始合并 %d 个文件 -> %s", len(paths), output)

        self.engine = MergeEngine(self.settings)
        self.worker = threading.Thread(
            target=self._merge_worker, args=(paths, output), name="merge", daemon=True
        )
        self.worker.start()

    def _merge_worker(self, paths: list[Path], output: Path) -> None:
        """后台线程：跑合并，结果丢进队列交给主线程处理。"""
        try:
            assert self.engine is not None
            report = self.engine.merge(
                paths, output, progress=self._on_progress_thread
            )
            self.ui_queue.put(("done", report))
        except Exception as exc:
            write_crash_report(type(exc), exc, exc.__traceback__, where="合并后台线程")
            self.ui_queue.put(("failed", exc))

    def _on_progress_thread(self, index: int, total: int, message: str) -> None:
        """进度回调（运行在后台线程）——只投队列，不碰控件。"""
        self.ui_queue.put(("progress", (index, total, message)))

    def cancel_merge(self) -> None:
        if self.engine:
            self.engine.stop.set()
            log.warning("用户请求取消，正在等待当前文件收尾 ...")
            self.btn_cancel.configure(state="disabled")

    # ------------------------------------------------------------------ #
    # 主线程轮询
    # ------------------------------------------------------------------ #
    def _pump(self) -> None:
        """定时把后台线程的消息搬到界面上（tkinter 只能主线程碰控件）。"""
        try:
            while True:
                kind, payload = self.ui_queue.get_nowait()
                if kind == "log":
                    self._append_log(payload)
                elif kind == "progress":
                    index, total, message = payload
                    pct = (index - 1) / max(1, total) * 100
                    self.progress.configure(value=pct)
                    self.status.configure(text=message, fg=COLOR_TEXT)
                elif kind == "done":
                    self._on_merge_done(payload)
                elif kind == "failed":
                    self._on_merge_failed(payload)
        except queue.Empty:
            pass
        if not self._closing:
            self.root.after(120, self._pump)

    def _enqueue_log(self, line: str) -> None:
        self.ui_queue.put(("log", line))

    def _append_log(self, line: str) -> None:
        """写日志窗口（保留最近 2000 行，防止内存无限长）。"""
        try:
            self.log_text.configure(state="normal")
            tag = "info"
            if "[ERROR" in line or "[CRITICAL" in line:
                tag = "error"
            elif "[WARNING" in line:
                tag = "warning"
            self.log_text.insert("end", line + "\n", tag)
            if int(self.log_text.index("end-1c").split(".")[0]) > 2000:
                self.log_text.delete("1.0", "500.0")
            self.log_text.see("end")
            self.log_text.configure(state="disabled")
        except Exception:
            pass

    # ------------------------------------------------------------------ #
    def _on_merge_done(self, report: MergeReport) -> None:
        self.btn_merge.configure(state="normal")
        self.btn_cancel.configure(state="disabled")
        self.progress.configure(value=100)
        self.report = report

        if report.error and not report.output_path:
            self._on_merge_failed(Exception(report.error))
            return

        summary = report.summary_line()
        self.status.configure(text=f"完成：{summary}", fg="#1E7A3C")
        log.info("合并结束：%s", summary)

        lines = [f"成功 {report.ok_count}/{len(report.outcomes)}"]
        if report.output_path:
            lines.append(f"输出：{report.output_path}")
            lines.append(f"大小：{_human_size(report.output_bytes)}")
        lines.append(f"耗时：{report.duration:.1f} 秒")

        if report.failed:
            lines.append("")
            lines.append("以下文件未能合并：")
            for o in report.failed[:10]:
                lines.append(f"  · {o.name}：{o.error}")
            if len(report.failed) > 10:
                lines.append(f"  … 还有 {len(report.failed) - 10} 个，见日志")

        if report.skipped:
            lines.append("")
            lines.append(f"跳过 {len(report.skipped)} 个（原因见日志）")

        text = "\n".join(lines)
        if report.failed:
            messagebox.showwarning("合并完成（部分失败）", text)
        else:
            if messagebox.askyesno("合并完成", text + "\n\n要现在打开这个文档吗？"):
                self._open_output(report.output_path)
        self._update_status()

    def _on_merge_failed(self, exc: Exception) -> None:
        self.btn_merge.configure(state="normal")
        self.btn_cancel.configure(state="disabled")
        self.progress.configure(value=0)
        self.status.configure(text=f"失败：{exc}", fg="#C0392B")
        log.error("合并失败: %s", exc)
        messagebox.showerror(
            "合并失败",
            f"{exc}\n\n详细原因已写入日志：\n{logs_dir()}",
        )

    @staticmethod
    def _open_output(path: Path | None) -> None:
        if not path:
            return
        try:
            os.startfile(str(path))  # noqa: S606
        except Exception as exc:
            log.warning("打开输出文件失败: %s", exc)

    # ------------------------------------------------------------------ #
    def _on_close(self) -> None:
        """关窗：能存就存，正在合并则先问一句。"""
        if self.worker and self.worker.is_alive():
            if not messagebox.askyesno("正在合并", "合并还在进行，确定要退出吗？"):
                return
            if self.engine:
                self.engine.stop.set()
        self._closing = True
        try:
            self.settings.window_geometry = self.root.winfo_geometry()
            save_settings(self.settings)
            unregister_listener(self._enqueue_log)
        except Exception:
            pass
        log.info("界面关闭")
        try:
            self.root.destroy()
        except Exception:
            pass

    def run(self) -> int:
        """进入事件循环。"""
        self.root.mainloop()
        return 0


def main() -> int:
    """GUI 入口。"""
    # 打包成 exe 后，冻结支持必须最先调用：任何一个被牵连的子进程
    # （PyInstaller 引导、multiprocessing）不加这句就会无限自我重启。
    try:
        import multiprocessing

        multiprocessing.freeze_support()
    except Exception:
        pass

    try:
        app = App()
        return app.run()
    except Exception as exc:
        write_crash_report(type(exc), exc, exc.__traceback__, where="GUI 启动")
        try:
            messagebox.showerror(
                "启动失败",
                f"{exc}\n\n崩溃日志已保存到：\n{logs_dir()}",
            )
        except Exception:
            print(f"启动失败: {exc}\n{traceback.format_exc()}")
        return 1


if __name__ == "__main__":  # pragma: no cover
    raise SystemExit(main())
