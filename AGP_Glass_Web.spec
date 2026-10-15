# -*- mode: python ; coding: utf-8 -*-
# PyInstaller spec — AGP Glass Suite (Artes) versión HTML / pywebview
#
# Uso (compilar FUERA de OneDrive el workpath, si no da "Acceso denegado"):
#   py -m PyInstaller AGP_Glass_Web.spec --noconfirm ^
#      --workpath "%TEMP%\agp_build" --distpath dist
#
# Sale en dist\AGP_Glass\AGP_Glass.exe (carpeta onedir, igual que la V7 que
# está en la red). El AGP_Glass.spec viejo (CustomTkinter) sigue intacto.

import os

a = Analysis(
    ['agp_web.py'],
    pathex=['.'],
    binaries=[],
    datas=[
        ('ui_web', 'ui_web'),
        ('INSTALAR_DRIVER.bat', '.'),
    ],
    hiddenimports=[
        'servicios',
        'db_app.asignaciones',
        'db_app.importar_excel',
        'db_app.logos_laser',
        'db_app.tracking',
        'db_app.auth',
        'db_app.conexiones_externas',
        'crear_arte_acad',
        'autocad_ops',
        'arte_maker',
        'agp_sso',
        'pymssql',
        'pymssql._pymssql',
        'pymssql._mssql',
        'win32com.client',
        'win32com.server',
        'win32clipboard',
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
    excludes=['matplotlib', 'numpy', 'pandas', 'scipy', 'PIL', 'flask', 'fastapi', 'uvicorn',
              'customtkinter'],
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data)

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
    icon='ui_web/assets/logo_agp.ico',
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

# INSTALAR_DRIVER.bat al lado del .exe (no solo dentro de _internal)
import shutil
_bat_src = os.path.join(SPECPATH, 'INSTALAR_DRIVER.bat')
_bat_dst = os.path.join(DISTPATH, 'AGP_Glass', 'INSTALAR_DRIVER.bat')
if os.path.isfile(_bat_src):
    shutil.copy2(_bat_src, _bat_dst)
