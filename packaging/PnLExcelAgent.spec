# -*- mode: python ; coding: utf-8 -*-

from pathlib import Path


project_root = Path.cwd()

a = Analysis(
    [str(project_root / "app.py")],
    pathex=[str(project_root)],
    binaries=[],
    datas=[(str(project_root / "config.yaml"), ".")],
    hiddenimports=[
        "pythoncom",
        "pywintypes",
        "win32timezone",
        "win32com.client",
        "yaml",
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=["pytest"],
    noarchive=False,
    optimize=1,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name="P_L_Excel_Agent",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=True,
    disable_windowed_traceback=False,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=False,
    upx_exclude=[],
    name="P_L_Excel_Agent",
)
