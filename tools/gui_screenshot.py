# -*- coding: utf-8 -*-
"""自动截图：启动 exe -> 等窗口出现 -> 截图 -> 关闭。

用来人工确认打包后的界面真的显示正常（不是白屏、不是控件错位）。
依赖 .NET 的 System.Drawing，只在 Windows 上有意义。
"""

from __future__ import annotations

import argparse
import subprocess
import sys
import time
from pathlib import Path


def find_window(pid_candidates: set[int], title_hint: str) -> int | None:
    """在给定 pid 集合（含子进程）里找一个有可见主窗口的进程。"""
    import ctypes
    from ctypes import wintypes

    user32 = ctypes.windll.user32
    result: list[int] = []

    @ctypes.WINFUNCTYPE(ctypes.c_bool, wintypes.HWND, wintypes.LPARAM)
    def _cb(hwnd, _lparam):  # type: ignore[no-untyped-def]
        if not user32.IsWindowVisible(hwnd):
            return True
        length = user32.GetWindowTextLengthW(hwnd)
        if length == 0:
            return True
        buf = ctypes.create_unicode_buffer(length + 1)
        user32.GetWindowTextW(hwnd, buf, length + 1)
        title = buf.value or ""
        if title_hint and title_hint in title:
            result.append(hwnd)
            return False
        return True

    user32.EnumWindows(_cb, 0)
    return result[0] if result else None


def screenshot(hwnd: int, out_path: Path) -> bool:
    """截取指定窗口（用 PrintWindow，窗口被遮挡也能截到）。"""
    import ctypes
    from ctypes import wintypes

    user32 = ctypes.windll.user32
    gdi32 = ctypes.windll.gdi32

    rect = wintypes.RECT()
    if not user32.GetWindowRect(hwnd, ctypes.byref(rect)):
        return False
    w = rect.right - rect.left
    h = rect.bottom - rect.top
    if w <= 0 or h <= 0:
        return False

    hdc_win = user32.GetWindowDC(hwnd)
    hdc_mem = gdi32.CreateCompatibleDC(hdc_win)
    hbitmap = gdi32.CreateCompatibleBitmap(hdc_win, w, h)
    gdi32.SelectObject(hdc_mem, hbitmap)

    # PW_RENDERFULLCONTENT = 2：让 DWM 把完整内容（含子控件）画进来
    user32.PrintWindow(hwnd, hdc_mem, 2)

    class BITMAPINFOHEADER(ctypes.Structure):
        _fields_ = [
            ("biSize", wintypes.DWORD), ("biWidth", wintypes.LONG),
            ("biHeight", wintypes.LONG), ("biPlanes", wintypes.WORD),
            ("biBitCount", wintypes.WORD), ("biCompression", wintypes.DWORD),
            ("biSizeImage", wintypes.DWORD), ("biXPelsPerMeter", wintypes.LONG),
            ("biYPelsPerMeter", wintypes.LONG), ("biClrUsed", wintypes.DWORD),
            ("biClrImportant", wintypes.DWORD),
        ]

    bmi = BITMAPINFOHEADER()
    bmi.biSize = ctypes.sizeof(BITMAPINFOHEADER)
    bmi.biWidth = w
    bmi.biHeight = -h  # 负数 = 自上而下
    bmi.biPlanes = 1
    bmi.biBitCount = 32
    bmi.biCompression = 0

    buf = ctypes.create_string_buffer(w * h * 4)
    gdi32.GetDIBits(hdc_mem, hbitmap, 0, h, buf, ctypes.byref(bmi), 0)

    from PIL import Image

    img = Image.frombuffer("RGBA", (w, h), buf, "raw", "BGRA", 0, 1)
    img.convert("RGB").save(out_path)

    gdi32.DeleteObject(hbitmap)
    gdi32.DeleteDC(hdc_mem)
    user32.ReleaseDC(hwnd, hdc_win)
    return True


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("exe", help="要启动的 exe 路径")
    ap.add_argument("--out", default="gui-screenshot.png")
    ap.add_argument("--title", default="Word 合并工具", help="窗口标题关键字")
    ap.add_argument("--timeout", type=float, default=60.0)
    ap.add_argument("--settle", type=float, default=4.0, help="窗口出现后再等几秒让界面画完")
    args = ap.parse_args()

    exe = Path(args.exe).resolve()
    if not exe.is_file():
        print(f"找不到 exe: {exe}", file=sys.stderr)
        return 2

    print(f"启动: {exe}")
    proc = subprocess.Popen([str(exe)], cwd=str(exe.parent))
    try:
        hwnd = None
        deadline = time.time() + args.timeout
        while time.time() < deadline:
            hwnd = find_window({proc.pid}, args.title)
            if hwnd:
                break
            if proc.poll() is not None:
                print(f"进程提前退出，退出码 {proc.returncode}", file=sys.stderr)
                return 3
            time.sleep(1)

        if not hwnd:
            print("超时：没找到目标窗口", file=sys.stderr)
            return 4

        print(f"找到窗口 hwnd={hwnd}，等待 {args.settle}s 让界面稳定 ...")
        time.sleep(args.settle)

        out = Path(args.out).resolve()
        if screenshot(hwnd, out):
            print(f"截图已保存: {out} ({out.stat().st_size / 1024:.1f} KB)")
            return 0
        print("截图失败", file=sys.stderr)
        return 5
    finally:
        try:
            proc.terminate()
            proc.wait(timeout=10)
        except Exception:
            try:
                proc.kill()
            except Exception:
                pass
        print("已关闭进程")


if __name__ == "__main__":
    sys.exit(main())
