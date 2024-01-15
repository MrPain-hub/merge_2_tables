# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['merge_for_xlsx_240115.py'],
    pathex=[],
    binaries=[],
    datas=[('.\picture', 'picture'),
           ],
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='merge_for_xlsx_240115',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    icon = '.\picture\my_icon.ico'
)
