# -*- coding: utf-8 -*-
"""PyInstaller 打包钩子：收集 tkinterdnd2 的原生库（tkdnd）。

不写这个钩子，PyInstaller 静态分析看不到 tkinterdnd2 需要的 .dll/.tcl，
打出来的 exe 一拖文件就报 "Unable to load tkdnd library"。
"""

from PyInstaller.utils.hooks import collect_data_files, collect_dynamic_libs

datas = collect_data_files("tkinterdnd2")
binaries = collect_dynamic_libs("tkinterdnd2")

# tkdnd 的原生库有时被放在 tkdnd 子目录里，两种都收一遍，确保不漏
try:
    datas += collect_data_files("tkinterdnd2", includes=["tkdnd/**/*"])
except Exception:
    pass
