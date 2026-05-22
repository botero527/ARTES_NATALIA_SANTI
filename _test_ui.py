# -*- coding: utf-8 -*-
"""Preview rápido del diálogo sin AutoCAD."""
import sys, os, types, datetime

win32com = types.ModuleType('win32com')
win32com.client = types.ModuleType('win32com.client')
pythoncom = types.ModuleType('pythoncom')
pythoncom.VT_ARRAY = 0; pythoncom.VT_R8 = 0; pythoncom.VT_DISPATCH = 0
sys.modules['win32com'] = win32com
sys.modules['win32com.client'] = win32com.client
sys.modules['pythoncom'] = pythoncom

_DIR = os.path.dirname(os.path.abspath(__file__))
g = {'__file__': os.path.join(_DIR, 'crear_arte_acad.py')}
exec(open(os.path.join(_DIR, 'crear_arte_acad.py'), encoding='utf-8').read(), g)

r = g['dialogo_cajetin']('1795-003-001-002')
print("Resultado:", r)
input("Enter para cerrar...")
