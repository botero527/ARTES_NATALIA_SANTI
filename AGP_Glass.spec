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
        # Script de instalación del driver ODBC (se copia al lado del .exe)
        ('INSTALAR_DRIVER.bat', '.'),
    ],
    hiddenimports=[
        'db_app.asignaciones',
        'db_app.importar_excel',
        'crear_arte_acad',
        'autocad_ops',
        'pymssql',
        'pymssql._pymssql',
        'pymssql._mssql',
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

# Copiar INSTALAR_DRIVER.bat al lado del .exe (no dentro de _internal)
import shutil
_bat_src = os.path.join(SPECPATH, 'INSTALAR_DRIVER.bat')
_bat_dst = os.path.join(DISTPATH, 'AGP_Glass', 'INSTALAR_DRIVER.bat')
if os.path.isfile(_bat_src):
    shutil.copy2(_bat_src, _bat_dst)
