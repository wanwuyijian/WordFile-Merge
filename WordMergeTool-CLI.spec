# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['packaging/entry_cli.py'],
    pathex=['src'],
    binaries=[],
    datas=[],
    hiddenimports=['tkinterdnd2', 'pymupdf', 'docx', 'pptx', 'openpyxl'],
    hookspath=['packaging/hooks'],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['matplotlib', 'numpy', 'pandas', 'scipy', 'PyQt5', 'PyQt6', 'PySide2', 'PySide6', 'IPython', 'notebook', 'pytest', 'setuptools', 'pip'],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='WordMergeTool-CLI',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=['H:/AIwork/FileProject/word合并工具/assets/app.ico'],
)
