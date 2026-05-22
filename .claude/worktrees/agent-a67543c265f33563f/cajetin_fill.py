# -*- coding: utf-8 -*-
# Script SEPARADO — abre ventana para rellenar los textos del Cajetin 1
# Correr en Rhino con _RunPythonScript despues de crear el arte
import rhinoscriptsyntax as rs  # type: ignore

CAMPOS = [
    ("DIBUJO",      "Dibujo"),
    ("REVISADO",    "Revisado por"),
    ("FECHA",       "Fecha de emision"),
    ("VEHICULO",    "Vehiculo"),
    ("VISTA",       "Vista"),
    ("COD PLANO",   "Codigo plano"),
    ("MODELO",      "Modelo"),
    ("NAGS",        "Cod. NAGS"),
    ("VITRO",       "Vitro"),
    ("MALLA",       "Malla"),
    ("VERSION",     "Version"),
    ("PIEZA",       "Pieza"),
    ("MEDIDAS",     "Medidas"),
    ("ESCALA",      "Escala"),
]

import tkinter as tk
from tkinter import ttk

valores = {}
cancelado = [False]

ventana = tk.Tk()
ventana.title("Rellenar Cajetin 1")
ventana.resizable(False, False)
ventana.attributes("-topmost", True)

frame = ttk.Frame(ventana, padding=16)
frame.grid(row=0, column=0, sticky="nsew")

entries = {}
for i, (campo, etiqueta) in enumerate(CAMPOS):
    ttk.Label(frame, text=etiqueta + ":", anchor="e", width=18).grid(
        row=i, column=0, sticky="e", pady=4, padx=(0, 8)
    )
    ent = ttk.Entry(frame, width=36)
    ent.grid(row=i, column=1, sticky="w", pady=4)
    entries[campo] = ent

# Foco en el primer campo
list(entries.values())[0].focus_set()

def _aceptar(event=None):
    for campo, ent in entries.items():
        valores[campo] = ent.get().strip()
    ventana.destroy()

def _cancelar(event=None):
    cancelado[0] = True
    ventana.destroy()

ventana.bind("<Return>", _aceptar)
ventana.bind("<Escape>", _cancelar)

btn_frame = ttk.Frame(frame)
btn_frame.grid(row=len(CAMPOS), column=0, columnspan=2, pady=(12, 0))
ttk.Button(btn_frame, text="Aceptar",  command=_aceptar).pack(side="left", padx=8)
ttk.Button(btn_frame, text="Cancelar", command=_cancelar).pack(side="left", padx=8)

ventana.mainloop()

if cancelado[0] or not valores:
    print("Cancelado — sin cambios.")
else:
    all_layers = rs.LayerNames() or []
    for campo, texto in valores.items():
        if not texto:
            continue
        layer_buscar = "CAJETIN 1${} 1".format(campo)
        layer_found = None
        for ln in all_layers:
            if ln.upper().endswith(layer_buscar.upper()):
                layer_found = ln
                break
        if layer_found is None:
            print("  WARN: layer '{}' no encontrado".format(layer_buscar))
            continue
        n = 0
        for oid in (rs.ObjectsByLayer(layer_found) or []):
            try:
                if rs.IsText(oid):
                    rs.TextObjectText(oid, texto)
                    n += 1
            except Exception:
                pass
        print("  {} -> '{}' ({} texto(s))".format(campo, texto, n))
    print("=== Cajetin fill: LISTO ===")
