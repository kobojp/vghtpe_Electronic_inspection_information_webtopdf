# -*- mode: python ; coding: utf-8 -*-
from pathlib import Path

project = Path(SPECPATH)
frontend = project / "frontend" / "dist"
if not (frontend / "index.html").is_file():
    raise RuntimeError("請先執行 npm --prefix frontend run build")

a = Analysis(
    [str(project / "main.py")],
    pathex=[str(project)],
    binaries=[(str(project / "wkhtmltopdf.exe"), ".")],
    datas=[(str(frontend), "frontend_dist"), (str(project / "app.ico"), ".")],
    hiddenimports=["webview.platforms.edgechromium", "uvicorn.logging", "uvicorn.loops.auto", "uvicorn.protocols.http.auto", "uvicorn.lifespan.on"],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=["tkinter", "PyQt5", "PyQt6", "PySide2", "PySide6"],
    noarchive=False,
)
pyz = PYZ(a.pure)
exe = EXE(
    pyz, a.scripts, a.binaries, a.datas, [],
    name="水電消防報表下載系統",
    debug=False,
    strip=False,
    upx=False,
    console=False,
    icon=str(project / "app.ico"),
)
