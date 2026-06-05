# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['crear_arte_acad.py'],
    pathex=[],
    binaries=[],
    datas=[('LAYERS Y CAJETINES 1.dwg', '.')],
    hiddenimports=['win32com', 'win32com.client', 'pythoncom', 'win32api'],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
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
    name='AGP_CrearArte',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
