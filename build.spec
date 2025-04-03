# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=[('wkhtmltopdf.exe', '.')],  # 包含 wkhtmltopdf.exe
    datas=[
        ('data.json', '.'),  # 包含 data.json
        ('vghtpe.png', '.'),  # 包含圖示檔案
    ],
    hiddenimports=[
        'PIL',
        'PIL._imagingtk',
        'PIL._tkinter_finder',
        'PIL.Image',
        'PIL.ImageTk',
        'PIL.PngImagePlugin',
        'PIL.JpegImagePlugin',
        'PIL.GifImagePlugin',
        'PIL.IcoImagePlugin',
        'PIL.BmpImagePlugin',
        'PIL.TiffImagePlugin',
        'PIL.WebPImagePlugin',
        'PIL.ImageDraw',
        'PIL.ImageFont'
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(
    a.pure,
    a.zipped_data,
    cipher=block_cipher
)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='水電消防報表下載系統',  # 執行檔名稱
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,  # 改為 False 以隱藏命令列視窗
    disable_windowed_traceback=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='app.ico',
) 