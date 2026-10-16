# -*- coding: utf-8 -*-
"""旧格式（.doc / .ppt / .xls / .wps / .et / .dps）处理。

这些是二进制格式，纯 Python 库读不了，必须借外部程序。按下面的顺序尝试，
**任何一步都不修改源文件**（一律"复制到工作区 -> 在副本上转换"）：

1. Microsoft Word/PowerPoint/Excel COM（装 Office 的主机，质量最好）
2. LibreOffice ``soffice --convert-to``（免 Office，但要装 LO）
3. 都不行 -> 抛出带清晰指引的错误，由引擎记进报告（不阻断其它文件）

关键安全点：COM 打开时用 ``ReadOnly=True``，且**永不 Save**；
LibreOffice 用 ``--outdir`` 输出到临时目录，源目录不会被写。
"""

from __future__ import annotations

import os
import shutil
import subprocess
import tempfile
from pathlib import Path

from ..logging_setup import get_logger

log = get_logger("converters.legacy")

#: 旧格式 -> 需要转成的现代格式
LEGACY_TARGETS = {
    ".doc": ".docx",
    ".dot": ".docx",
    ".rtf": ".docx",
    ".wps": ".docx",
    ".ppt": ".pptx",
    ".pps": ".pptx",
    ".dps": ".pptx",
    ".xls": ".xlsx",
    ".et": ".xlsx",
}

SOFFICE_CANDIDATES = (
    r"C:\Program Files\LibreOffice\program\soffice.exe",
    r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
    r"C:\Program Files\OpenOffice 4\program\soffice.exe",
)

_WORD_PROGIDS = ("Word.Application",)
_EXCEL_PROGIDS = ("Excel.Application",)
_PPT_PROGIDS = ("PowerPoint.Application")


def is_legacy(suffix: str) -> bool:
    """该扩展名是否需要旧格式转换。"""
    return suffix.lower() in LEGACY_TARGETS


def find_soffice() -> str | None:
    """定位 LibreOffice/OpenOffice。"""
    found = shutil.which("soffice") or shutil.which("soffice.exe")
    if found:
        return found
    for c in SOFFICE_CANDIDATES:
        if Path(c).is_file():
            return c
    return None


def convert_legacy(src: Path, workdir: Path) -> tuple[Path, str]:
    """把旧格式转成现代格式，返回 (新文件路径, 使用的通道名)。

    转换在**工作区**里做：先把源文件复制过去再转，源文件绝不被写。
    """
    suffix = src.suffix.lower()
    target_ext = LEGACY_TARGETS.get(suffix)
    if target_ext is None:
        raise ValueError(f"不支持的旧格式: {suffix}")

    workdir.mkdir(parents=True, exist_ok=True)
    # 复制到工作区（源文件不改）
    safe_name = f"legacy-input-{abs(hash(str(src))) % 10**8}{suffix}"
    local = workdir / safe_name
    shutil.copy2(src, local)

    # 通道 1：Office COM
    converted = _convert_via_com(local, workdir, target_ext)
    if converted:
        return converted, "Office COM"

    # 通道 2：LibreOffice
    converted = _convert_via_soffice(local, workdir, target_ext)
    if converted:
        return converted, "LibreOffice"

    raise ValueError(
        f"无法转换旧格式 {suffix}：本机既没有可用的 Microsoft Office，也没有 LibreOffice。\n"
        f"  解决办法（二选一）：\n"
        f"    A. 在本机/目标机安装 Microsoft Office；\n"
        f"    B. 安装 LibreOffice（免费，装完本工具自动识别）。\n"
        f"  另外也可以先用 Office 把该文件另存为 {target_ext} 再合并。"
    )


# --------------------------------------------------------------------------- #
def _convert_via_com(src: Path, workdir: Path, target_ext: str) -> Path | None:
    """用 Office COM 转换（只读打开、另存到工作区、绝不写回源）。"""
    if os.name != "nt":
        return None
    try:
        import pythoncom  # type: ignore
        import win32com.client  # type: ignore
    except Exception:
        return None

    suffix = src.suffix.lower()
    out = workdir / (src.stem + target_ext)
    app = None
    doc = None
    try:
        pythoncom.CoInitialize()
        if suffix in (".doc", ".dot", ".rtf", ".wps"):
            app = win32com.client.DispatchEx("Word.Application")
            app.Visible = False
            app.DisplayAlerts = 0
            doc = app.Documents.Open(str(src), ReadOnly=True, AddToRecentFiles=False)
            # FileFormat=16 -> wdFormatXMLDocument (.docx)
            doc.SaveAs2(str(out), FileFormat=16)
        elif suffix in (".xls", ".et"):
            app = win32com.client.DispatchEx("Excel.Application")
            app.Visible = False
            app.DisplayAlerts = False
            doc = app.Workbooks.Open(str(src), ReadOnly=True, UpdateLinks=0)
            # FileFormat=51 -> xlOpenXMLWorkbook (.xlsx)
            doc.SaveAs(str(out), FileFormat=51)
        elif suffix in (".ppt", ".pps", ".dps"):
            app = win32com.client.DispatchEx("PowerPoint.Application")
            doc = app.Presentations.Open(str(src), ReadOnly=True, Untitled=False, WithWindow=False)
            # ppSaveAsOpenXMLPresentation = 24
            doc.SaveAs(str(out), 24)
        else:
            return None

        if out.is_file() and out.stat().st_size > 0:
            log.info("旧格式转换成功（Office COM）: %s -> %s", src.name, out.name)
            return out
        return None
    except Exception as exc:
        log.warning("Office COM 转换失败（%s）: %s", suffix, exc)
        return None
    finally:
        # 关文档、退程序；只读打开，源文件不会变
        try:
            if doc is not None:
                if suffix in (".doc", ".dot", ".rtf", ".wps"):
                    doc.Close(SaveChanges=0)
                elif suffix in (".xls", ".et"):
                    doc.Close(SaveChanges=False)
                else:
                    doc.Close()
        except Exception:
            pass
        try:
            if app is not None:
                app.Quit()
        except Exception:
            pass
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass


def _convert_via_soffice(src: Path, workdir: Path, target_ext: str) -> Path | None:
    """用 LibreOffice 转换（输出到工作区，源目录不受影响）。"""
    soffice = find_soffice()
    if not soffice:
        return None
    fmt = {
        ".docx": "docx:MS Word 2007 XML",
        ".xlsx": "xlsx:Calc MS Excel 2007 XML",
        ".pptx": "pptx:Impress MS PowerPoint 2007 XML",
    }.get(target_ext, target_ext.lstrip("."))

    outdir = workdir / "soffice-out"
    outdir.mkdir(parents=True, exist_ok=True)
    try:
        env = dict(os.environ)
        env["HOME"] = str(workdir)  # 隔离用户配置，避免污染真实 LibreOffice 配置
        subprocess.run(
            [
                soffice,
                "-env:UserInstallation=file:///" + str(workdir / "lo-profile").replace("\\", "/"),
                "--headless",
                "--norestore",
                "--norecentdocs",
                "--convert-to",
                fmt,
                "--outdir",
                str(outdir),
                str(src),
            ],
            check=True,
            timeout=300,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            env=env,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        )
        produced = list(outdir.glob(f"*{target_ext}"))
        if produced:
            log.info("旧格式转换成功（LibreOffice）: %s -> %s", src.name, produced[0].name)
            return produced[0]
        return None
    except Exception as exc:
        log.warning("LibreOffice 转换失败: %s", exc)
        return None
    finally:
        try:
            for p in outdir.glob(".~lock.*"):
                p.unlink(missing_ok=True)
        except Exception:
            pass


def legacy_help_text() -> str:
    """给用户的排障提示（出现在报告和日志里）。"""
    soffice = find_soffice()
    return (
        "旧格式（.doc/.ppt/.xls）需要外部转换器。\n"
        f"  当前 LibreOffice: {soffice or '未找到'}\n"
        "  安装 LibreOffice 即可支持这些格式；或先用 Office 另存为新格式。"
    )
