# -*- mode: python ; coding: utf-8 -*-
# PyInstaller spec — AGP Glass Suite
# Uso: pyinstaller AGP_Glass.spec
#
# El DWG del cajetín y el Excel se leen desde la red (\\192.168.2.37\...)
# No se incluyen en el .exe — cualquier actualización en la red aplica de una.

import os, sys

block_cipher = None

a = Analysis(
    ['agp_app.py'],
    pathex=['.'],
    binaries=[],
    datas=[
        # CustomTkinter necesita sus assets (temas, imágenes)
        ('.venv\\Lib\\site-packages\\customtkinter', 'customtkinter'),
    ],
    hiddenimports=[
        'db_app.asignaciones',
        'db_app.importar_excel',
        'crear_arte_acad',
        'autocad_ops',
        'pyodbc',
        'win32com.client',
        'win32com.server',
        'pythoncom',
        'pywintypes',
        'openpyxl',
        'openpyxl.cell._writer',
        'tkinter',
        'tkinter.ttk',
        'tkinter.messagebox',
        'tkinter.filedialog',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['matplotlib', 'numpy', 'pandas', 'scipy', 'PIL', 'flask', 'fastapi', 'uvicorn'],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='AGP_Glass',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=None,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='AGP_Glass',
)
