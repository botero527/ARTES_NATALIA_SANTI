#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
AGP Glass — App unificada
  · Crear Arte    (pipeline AutoCAD)
  · Consultar BD  (Azure SQL — vitros, mallas, vinilos, pasta plata)
Requiere: customtkinter, pyodbc, pywin32
"""

import os, sys, time, threading, subprocess, traceback

try:
    import customtkinter as ctk
    from customtkinter import CTkFont
except ImportError:
    import tkinter as tk
    tk.Tk().withdraw()
    import tkinter.messagebox as mb
    mb.showerror("Dependencia faltante",
                 "Falta customtkinter.\nEjecuta:  pip install customtkinter")
    sys.exit(1)

try:
    import pymssql as _pymssql
except ImportError:
    _pymssql = None

try:
    import pyodbc
except ImportError:
    pyodbc = None

try:
    import win32com.client, pythoncom
    _COM_OK = True
except ImportError:
    _COM_OK = False

try:
    from autocad_ops import AutoCADMotor
    _MOTOR_OK = True
except Exception:
    _MOTOR_OK = False

try:
    from crear_arte_acad import dialogo_cajetin, pipeline as _pipeline_acad, ErrorGuardadoArte
    _PIPELINE_OK = True
except Exception:
    _PIPELINE_OK = False
    class ErrorGuardadoArte(RuntimeError): pass

try:
    from db_app.asignaciones import (
        actualizar_ruta_arte as _actualizar_ruta_arte,
        anular_asignacion as _anular_asignacion,
        dialogo_separar as _dialogo_separar,
        limpiar_pendientes_huerfanos as _limpiar_pendientes_huerfanos,
    )
    _ASIGN_OK = True
except Exception:
    _ASIGN_OK = False

import re, math, shutil, json, datetime
import tkinter as _tk_root
from tkinter import filedialog, messagebox

_SENTINEL = object()  # valor centinela para campos virtuales no presentes

def _msgbox_topmost(tipo, titulo, mensaje):
    """Muestra un messagebox siempre al frente de todas las ventanas."""
    import tkinter as _tk
    try:
        root = _tk._default_root
    except Exception:
        root = None
    tmp = _tk.Toplevel(root)
    tmp.withdraw()
    tmp.attributes("-topmost", True)
    tmp.lift()
    tmp.focus_force()
    if tipo == "error":
        messagebox.showerror(titulo, mensaje, parent=tmp)
    elif tipo == "warning":
        messagebox.showwarning(titulo, mensaje, parent=tmp)
    else:
        messagebox.showinfo(titulo, mensaje, parent=tmp)
    try: tmp.destroy()
    except Exception: pass

# ══════════════════════════════════════════════════════════════════════════════
#  CONFIGURACIÓN
# ══════════════════════════════════════════════════════════════════════════════
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

CONN_AZURE = (
    "DRIVER={ODBC Driver 17 for SQL Server};"
    "SERVER=agpcolombia.database.windows.net,1433;"
    "DATABASE=AGP_Ingenieria;"
    "UID=DevIngenieria;"
    "PWD=HiJE068i0LQVrwA;"
    "Encrypt=yes;TrustServerCertificate=no;Connection Timeout=30;"
)

# Paleta
PAL = {
    "bg":        "#0F1117",
    "sidebar":   "#161B27",
    "card":      "#1C2333",
    "card2":     "#222A3E",
    "border":    "#2E3A55",
    "accent":    "#3B82F6",
    "accent2":   "#1D4ED8",
    "green":     "#10B981",
    "green2":    "#059669",
    "orange":    "#F59E0B",
    "red":       "#EF4444",
    "purple":    "#8B5CF6",
    "txt":       "#F1F5F9",
    "txt_mid":   "#94A3B8",
    "txt_dim":   "#475569",
    "log_bg":    "#0A0D14",
}

FONT = lambda s, w="normal": CTkFont(family="Segoe UI", size=s, weight=w)
MONO = lambda s=11: CTkFont(family="Consolas", size=s)

# ══════════════════════════════════════════════════════════════════════════════
#  BD — helpers
# ══════════════════════════════════════════════════════════════════════════════
# Wrapper para que pymssql acepte placeholders '?' igual que pyodbc
class _CursorWrap:
    def __init__(self, cur): self._c = cur
    def execute(self, sql, params=()):
        return self._c.execute(sql.replace("?", "%s"), params or ())
    def executemany(self, sql, seq):
        return self._c.executemany(sql.replace("?", "%s"), seq)
    def __getattr__(self, n): return getattr(self._c, n)

class _ConnWrap:
    def __init__(self, conn): self._c = conn
    def cursor(self): return _CursorWrap(self._c.cursor())
    def execute(self, sql, params=()):
        cur = self.cursor(); cur.execute(sql, params); return cur
    def commit(self):   self._c.commit()
    def rollback(self): self._c.rollback()
    def close(self):    self._c.close()
    def __enter__(self): return self
    def __exit__(self, *a): self._c.__exit__(*a)

def db_connect():
    if _pymssql is None:
        raise RuntimeError("pymssql no disponible — recompila el .exe")
    try:
        conn = _pymssql.connect(
            server="agpcolombia.database.windows.net",
            port=1433,
            user="DevIngenieria",
            password="HiJE068i0LQVrwA",
            database="AGP_Ingenieria",
            timeout=20,
            login_timeout=20,
            charset="UTF-8",
            tds_version="7.3",
        )
        return _ConnWrap(conn)
    except Exception as e:
        raise RuntimeError(
            f"No se pudo conectar a la base de datos.\n"
            f"Verifica que tienes acceso a internet o red interna.\n"
            f"Detalle: {e}"
        )

def db_query(sql, params=()):
    conn = db_connect()
    cur  = conn.cursor()
    cur.execute(sql, params)
    cols = [c[0] for c in cur.description]
    rows = [dict(zip(cols, r)) for r in cur.fetchall()]
    conn.close()
    return rows

# ══════════════════════════════════════════════════════════════════════════════
#  ARTE — helpers (portados de arte_maker.py)
# ══════════════════════════════════════════════════════════════════════════════
def _ruta_planos(ruta_dwg):
    dest = os.path.join(os.path.dirname(os.path.abspath(ruta_dwg)), "PLANOS")
    os.makedirs(dest, exist_ok=True)
    return dest

def _ruta_arte_salida(ruta_dwg, malla="", pieza="", nombre_archivo=""):
    # Siempre guarda en ARTES/BN/ — crea las carpetas si no existen
    artes = os.path.join(os.path.dirname(os.path.abspath(ruta_dwg)), "ARTES")
    dest  = os.path.join(artes, "BN")
    os.makedirs(dest, exist_ok=True)

    def _limpiar_nombre(s):
        # Reemplaza caracteres inválidos en nombres de archivo Windows
        for c in r'/\:*?"<>|':
            s = s.replace(c, "-")
        while "  " in s: s = s.replace("  ", " ")
        while "--" in s: s = s.replace("--", "-")
        return s.strip("- ")

    if nombre_archivo:
        base = nombre_archivo if nombre_archivo.lower().endswith(".dwg") else nombre_archivo + ".dwg"
        nombre = _limpiar_nombre(os.path.splitext(base)[0]) + ".dwg"
    else:
        # Mallas múltiples separadas por "/" o "," → unir con " - "
        mallas_str = " - ".join(
            m.strip() for m in malla.replace(",", "/").split("/") if m.strip()
        ) if malla.strip() else ""
        if mallas_str and pieza.strip():
            nombre = f"P {mallas_str} {pieza.strip()}.dwg"
        elif mallas_str:
            nombre = f"P {mallas_str}.dwg"
        elif pieza.strip():
            nombre = f"P {pieza.strip()}.dwg"
        else:
            nombre = "P " + os.path.splitext(os.path.basename(ruta_dwg))[0] + ".dwg"
        nombre = _limpiar_nombre(os.path.splitext(nombre)[0]) + ".dwg"

    return os.path.join(dest, nombre)

def _extraer_codigos(ruta):
    base   = os.path.splitext(os.path.basename(ruta))[0]
    grupos = re.findall(r'\d+', base)
    if not grupos: return []
    codigos, total = [], 0
    for g in reversed(grupos):
        if total + len(g) > 6: break
        codigos.insert(0, g); total += len(g)
    return codigos

def _buscar_artes(ruta, codigos):
    res = []
    for raiz, dirs, archivos in os.walk(ruta):
        dirs[:] = [d for d in dirs if not d.startswith(".")]
        partes = raiz.replace("\\", "/").upper().split("/")
        if not any(p == "ARTES" or p.startswith("ARTES") for p in partes): continue
        for arch in sorted(archivos):
            ext = os.path.splitext(arch)[1].lower()
            if ext not in (".dwg", ".3dm"): continue
            if not arch.upper().startswith("P"): continue
            nums = re.findall(r'\d+', os.path.splitext(arch)[0])
            if codigos and (len(nums) < len(codigos) or nums[-len(codigos):] != codigos): continue
            res.append({"version": os.path.relpath(raiz, ruta),
                        "archivo": arch,
                        "ruta_completa": os.path.join(raiz, arch)})
    return sorted(res, key=lambda x: (x["version"], x["archivo"]))

def _crear_arte_autocad(ruta_dwg, log_fn=None, valores_cajetin=None,
                        ruta_salida=None, perim_index=0, _com_ya_init=False,
                        compensar=False):
    """
    _com_ya_init=True → el llamador ya hizo CoInitialize; no llamar de nuevo.
    """
    if log_fn is None: log_fn = print
    if not _PIPELINE_OK: raise RuntimeError("Pipeline no disponible")
    if not _COM_OK: raise RuntimeError("pywin32 no disponible")
    if not _com_ya_init:
        pythoncom.CoInitialize()
    try:
        try:
            acad = win32com.client.GetActiveObject("AutoCAD.Application")
        except Exception:
            raise RuntimeError("AutoCAD no está abierto")
        log_fn(f"  Abriendo: {os.path.basename(ruta_dwg)}")
        doc = acad.Documents.Open(os.path.abspath(ruta_dwg), False, False)
        time.sleep(2.5)
        try: doc.Activate(); time.sleep(0.5)
        except Exception: pass
        # Esperar a que AutoCAD termine de cargar el documento
        for _ in range(30):
            try:
                nombre = doc.FullName
                if nombre and os.path.exists(nombre):
                    _ = doc.ModelSpace.Count  # fuerza carga completa
                    break
            except Exception:
                pass
            time.sleep(0.5)
        n = _pipeline_acad(doc, log_fn=log_fn, valores_cajetin=valores_cajetin,
                           ruta_salida=ruta_salida, perim_index=perim_index,
                           compensar=compensar)
        return n or 1
    finally:
        if not _com_ya_init:
            pythoncom.CoUninitialize()

# ══════════════════════════════════════════════════════════════════════════════
#  HELPER — copiar desde treeview
# ══════════════════════════════════════════════════════════════════════════════
def _setup_tree_copy(tree, toast_widget=None):
    """
    Estilo Excel: click recuerda la celda, Ctrl+C copia ese valor.
    Sin celda seleccionada, Ctrl+C copia toda la fila.
    """
    _sel = {"item": None, "col": None}   # celda actualmente recordada

    def _toast(msg):
        if toast_widget is None:
            return
        try:
            toast_widget.configure(text=msg, text_color="#22c55e")
            toast_widget.after(1800, lambda: toast_widget.configure(
                text="", text_color=PAL["txt_dim"]))
        except Exception:
            pass

    def _on_click(event):
        item = tree.identify_row(event.y)
        col  = tree.identify_column(event.x)
        if item and col:
            _sel["item"] = item
            _sel["col"]  = int(col.lstrip("#")) - 1

    def _on_ctrl_c(event):
        item = _sel["item"]
        col  = _sel["col"]
        # Si hay celda recordada úsala; si no, copia toda la fila seleccionada
        if item and col is not None:
            try:
                vals = tree.item(item, "values")
                val  = str(vals[col]) if col < len(vals) else ""
                if val and val not in ("—", ""):
                    tree.clipboard_clear()
                    tree.clipboard_append(val)
                    _toast(f"✔  {val[:60]}")
                    return
            except Exception:
                pass
        sel = tree.selection()
        if sel:
            vals = tree.item(sel[0], "values")
            tree.clipboard_clear()
            tree.clipboard_append("\t".join(str(v) for v in vals))
            _toast("Fila copiada ✔")

    tree.bind("<ButtonRelease-1>", _on_click, add=True)
    tree.bind("<Control-c>",       _on_ctrl_c, add=True)
    tree.bind("<Control-C>",       _on_ctrl_c, add=True)
    tree.bind("<MouseWheel>",
              lambda e: tree.yview_scroll(int(-e.delta / 120), "units"), add=True)

# ══════════════════════════════════════════════════════════════════════════════
#  WIDGETS REUTILIZABLES
# ══════════════════════════════════════════════════════════════════════════════
class SideBtn(ctk.CTkButton):
    """Botón de navegación del sidebar."""
    def __init__(self, parent, text, icon, command, **kw):
        super().__init__(parent,
            text=f"  {icon}  {text}",
            anchor="w",
            height=52,
            corner_radius=8,
            fg_color="transparent",
            hover_color=PAL["border"],
            text_color=PAL["txt_mid"],
            font=FONT(14),
            command=command, **kw)

    def set_active(self, active: bool):
        if active:
            self.configure(fg_color=PAL["accent2"], text_color="white",
                           font=FONT(14, "bold"))
        else:
            self.configure(fg_color="transparent", text_color=PAL["txt_mid"],
                           font=FONT(14))


class LogBox(ctk.CTkTextbox):
    """Consola de log con colores."""
    TAGS = {"ok": PAL["green"], "warn": PAL["orange"],
            "err": PAL["red"],  "dim": PAL["txt_mid"]}

    def __init__(self, parent, **kw):
        super().__init__(parent, font=MONO(10), state="disabled",
                         fg_color=PAL["log_bg"], text_color=PAL["txt"],
                         wrap="word", **kw)

    def write(self, msg: str, tag: str = ""):
        self.configure(state="normal")
        ts = time.strftime("%H:%M:%S")
        color = self.TAGS.get(tag, PAL["txt"])
        self.insert("end", f"{ts}  {msg}\n")
        self.see("end")
        self.configure(state="disabled")

    def clear(self):
        self.configure(state="normal")
        self.delete("1.0", "end")
        self.configure(state="disabled")


class StatCard(ctk.CTkFrame):
    def __init__(self, parent, label, icon, color, **kw):
        super().__init__(parent, fg_color=PAL["card"], corner_radius=14,
                         border_width=2, border_color=PAL["border"], **kw)
        ctk.CTkLabel(self, text=icon, font=FONT(26)).pack(pady=(16,0))
        self._val = ctk.CTkLabel(self, text="—", font=FONT(28, "bold"),
                                  text_color=color)
        self._val.pack()
        ctk.CTkLabel(self, text=label, font=FONT(11),
                     text_color=PAL["txt_mid"]).pack(pady=(4,14))

    def set(self, v): self._val.configure(text=f"{v:,}" if isinstance(v, int) else str(v))


def _anular_bd(valores, log_fn):
    """Anula vitro y malla en BD cuando el arte falla — evita números bloqueados."""
    if not _ASIGN_OK:
        return
    vitro = valores.get("VITRO", "").strip()
    malla = valores.get("MALLA", "").strip()
    anulados = []
    try:
        if vitro:
            _anular_asignacion("vitrojet", vitro)
            anulados.append(vitro)
        for m in malla.split("/"):
            m = m.strip()
            if not m:
                continue
            tab = "grandes" if m.upper().startswith("A-") else "pequenas"
            _anular_asignacion(tab, m)
            anulados.append(m)
        if anulados:
            log_fn(f"⚠ Asignación anulada en BD: {', '.join(anulados)}", "warn")
    except Exception as e:
        log_fn(f"WARN: no se pudo anular en BD: {e}", "warn")

# ══════════════════════════════════════════════════════════════════════════════
#  PESTAÑA — CREAR ARTE
# ══════════════════════════════════════════════════════════════════════════════
class TabArte(ctk.CTkFrame):
    def __init__(self, parent, **kw):
        super().__init__(parent, fg_color="transparent", **kw)
        self._ruta_base   = ctk.StringVar()
        self._ruta_dwg    = ctk.StringVar()
        self._compensar   = ctk.BooleanVar(value=False)
        self._resultados  = []
        self._on_art_done = None  # callback → AGPApp notifica a otras pestañas
        self._build()

    def _build(self):
        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)

        # Scroll interno para los cards
        self._inner = ctk.CTkScrollableFrame(self, fg_color="transparent", corner_radius=0)
        self._inner.grid(row=0, column=0, sticky="nsew")
        self._inner.columnconfigure(0, weight=1)

        # ── Sección inputs ────────────────────────────────────────────────────
        card_in = self._card("CONFIGURACIÓN", row=0)
        card_in.columnconfigure(1, weight=1)

        self._field(card_in, "Ruta vehículo / versión:", 0)
        self._e_base = ctk.CTkEntry(card_in, textvariable=self._ruta_base,
                                     height=42, font=FONT(12),
                                     fg_color=PAL["card2"], border_color=PAL["border"],
                                     text_color=PAL["accent"])
        self._e_base.grid(row=1, column=0, columnspan=2, sticky="ew", padx=(0,6), pady=(2,8))
        ctk.CTkButton(card_in, text="📂 Explorar", width=110, height=42,
                      fg_color=PAL["border"], hover_color=PAL["accent2"],
                      font=FONT(11), command=self._pick_base
                      ).grid(row=1, column=2, pady=(2,8))

        self._field(card_in, "Plano DWG original:", 2)
        self._e_dwg = ctk.CTkEntry(card_in, textvariable=self._ruta_dwg,
                                    height=42, font=FONT(12),
                                    fg_color=PAL["card2"], border_color=PAL["border"],
                                    text_color=PAL["accent"])
        self._e_dwg.grid(row=3, column=0, columnspan=2, sticky="ew", padx=(0,6), pady=(2,4))
        ctk.CTkButton(card_in, text="📂 Explorar", width=110, height=42,
                      fg_color=PAL["border"], hover_color=PAL["accent2"],
                      font=FONT(11), command=self._pick_dwg
                      ).grid(row=3, column=2, pady=(2,4))

        # ── Botones workflow ──────────────────────────────────────────────────
        card_wf = self._card("WORKFLOW", row=1)
        btn_row = ctk.CTkFrame(card_wf, fg_color="transparent")
        btn_row.pack(fill="x", pady=6)

        self._btns = []
        defs = [
            ("⬇  Extraer Plano",  PAL["green2"],   "#1a6640", self._extraer),
            ("🎨  Crear Arte",     PAL["purple"],   "#5b21b6", self._crear_arte),
            ("🔍  Buscar Arte",   "#E67E22",        "#ca6f1e", self._buscar),
            ("⚡  Todo en Uno",   "#E63946",        "#b71c2e", self._todo_en_uno),
        ]
        for txt, color, hover, cmd in defs:
            b = ctk.CTkButton(btn_row, text=txt, width=200, height=52,
                              fg_color=color, hover_color=hover,
                              font=FONT(12, "bold"), corner_radius=12,
                              command=cmd)
            b.pack(side="left", padx=6)
            self._btns.append(b)

        opt_row = ctk.CTkFrame(card_wf, fg_color="transparent")
        opt_row.pack(fill="x", pady=(6,0))
        ctk.CTkCheckBox(opt_row, text="Compensar perímetro  (offset 3 mm hacia adentro)",
                        variable=self._compensar,
                        font=FONT(11), text_color=PAL["txt_mid"],
                        fg_color=PAL["accent2"], hover_color=PAL["accent"],
                        checkmark_color="white", corner_radius=4,
                        ).pack(side="left", padx=6)


        self._prog = ctk.CTkProgressBar(card_wf, mode="indeterminate",
                                         height=6, progress_color=PAL["accent"])
        self._prog.pack(fill="x", pady=(4,0))

        # ── Tabla resultados ──────────────────────────────────────────────────
        import tkinter.ttk as ttk
        card_tbl = self._card("ARTES ENCONTRADOS", row=2)
        self._lbl_tbl = ctk.CTkLabel(card_tbl, text="0 resultados",
                                      font=FONT(11), text_color=PAL["txt_mid"])
        self._lbl_tbl.pack(anchor="w", padx=2, pady=(0,4))

        import tkinter as tk
        style = ttk.Style()
        style.theme_use("clam")
        style.configure("AGP.Treeview",
            background=PAL["card"],
            foreground=PAL["txt"],
            fieldbackground=PAL["card"],
            borderwidth=0,
            font=("Segoe UI", 11),
            rowheight=32)
        style.configure("AGP.Treeview.Heading",
            background=PAL["card2"],
            foreground=PAL["accent"],
            font=("Segoe UI", 10, "bold"),
            relief="flat",
            padding=(8, 6))
        style.map("AGP.Treeview", background=[("selected", PAL["accent2"])],
                  foreground=[("selected","white")])

        frm_tbl = ctk.CTkFrame(card_tbl, fg_color=PAL["card"], corner_radius=6)
        frm_tbl.pack(fill="both", expand=True)

        self._tree = ttk.Treeview(frm_tbl, style="AGP.Treeview",
                                   columns=("version","archivo","tipo"),
                                   show="headings", height=5)
        for col, w, lbl in [("version",320,"Versión"),("archivo",280,"Archivo"),("tipo",60,"Tipo")]:
            self._tree.heading(col, text=lbl)
            self._tree.column(col, width=w)
        self._tree.tag_configure("match", background="#0A2015", foreground=PAL["green"])
        self._tree.pack(side="left", fill="both", expand=True, padx=6, pady=6)
        sb = ttk.Scrollbar(frm_tbl, orient="vertical", command=self._tree.yview)
        sb.pack(side="right", fill="y", pady=6)
        self._tree.configure(yscrollcommand=sb.set)
        self._tree.bind("<Double-1>", self._on_doble_click)
        _setup_tree_copy(self._tree)

        ctk.CTkLabel(card_tbl, text="  Doble clic en fila verde → superponer en AutoCAD",
                     font=FONT(10), text_color=PAL["txt_dim"]).pack(anchor="w")

        # ── Log ───────────────────────────────────────────────────────────────
        card_log = self._card("CONSOLA", row=3)
        btn_row_log = ctk.CTkFrame(card_log, fg_color="transparent")
        btn_row_log.pack(fill="x")
        ctk.CTkButton(btn_row_log, text="Limpiar", width=80, height=26,
                      fg_color=PAL["border"], hover_color=PAL["red"],
                      font=FONT(10), command=lambda: self._log.clear()
                      ).pack(side="right")
        self._log = LogBox(card_log, height=180)
        self._log.pack(fill="both", expand=True, pady=(4,0))

    # ── helpers ──────────────────────────────────────────────────────────────
    def _card(self, title, row):
        outer = ctk.CTkFrame(self._inner, fg_color=PAL["card"],
                              corner_radius=10, border_width=1,
                              border_color=PAL["border"])
        outer.grid(row=row, column=0, sticky="ew", padx=4, pady=6)
        ctk.CTkLabel(outer, text=title, font=FONT(10, "bold"),
                     text_color=PAL["txt_mid"]).pack(anchor="w", padx=14, pady=(8,0))
        inner = ctk.CTkFrame(outer, fg_color="transparent")
        inner.pack(fill="both", expand=True, padx=14, pady=(4,10))
        return inner

    def _field(self, parent, text, row):
        ctk.CTkLabel(parent, text=text, font=FONT(11, "bold"),
                     text_color=PAL["txt_mid"]
                     ).grid(row=row, column=0, columnspan=3, sticky="w", pady=(6,0))

    def _pick_base(self):
        r = filedialog.askdirectory(title="Seleccionar carpeta")
        if r: self._ruta_base.set(r.replace("/","\\"))

    def _pick_dwg(self):
        r = filedialog.askopenfilename(
            title="Seleccionar plano DWG",
            initialdir=self._ruta_base.get() or "/",
            filetypes=[("AutoCAD DWG","*.dwg"),("Todos","*.*")])
        if r: self._ruta_dwg.set(r.replace("/","\\"))

    def _busy(self, on):
        state = "disabled" if on else "normal"
        for b in self._btns: b.configure(state=state)
        if on: self._prog.start()
        else:  self._prog.stop(); self._prog.set(0)

    def _log_fn(self, msg, tag=""):
        self.after(0, self._log.write, msg, tag)

    def _validar(self, dwg=True):
        if not os.path.isdir(self._ruta_base.get().strip()):
            _msgbox_topmost("warning", "Campo requerido", "Indica una ruta base válida.")
            return False
        if dwg and not os.path.isfile(self._ruta_dwg.get().strip().strip('"')):
            _msgbox_topmost("warning", "Campo requerido", "Selecciona el DWG del plano.")
            return False
        return True

    # ── acciones ─────────────────────────────────────────────────────────────
    def _extraer(self):
        if not self._validar(): return
        self._busy(True)
        threading.Thread(target=self._t_extraer, daemon=True).start()

    def _t_extraer(self):
        dwg = self._ruta_dwg.get().strip().strip('"')
        self._log_fn("="*50)
        self._log_fn("EXTRAER PLANO...", "ok")
        nombre = os.path.splitext(os.path.basename(dwg))[0]
        dest   = os.path.join(_ruta_planos(dwg), f"{nombre}_PLANO.dwg")
        try:
            if not _MOTOR_OK: raise RuntimeError("AutoCADMotor no disponible")
            if not _COM_OK:   raise RuntimeError("pywin32 no disponible")
            # AutoCADMotor.__init__ ya llama CoInitialize; lo balanceamos aquí
            motor = AutoCADMotor()
            try:
                motor.extraer_layers(dwg, dest, log_fn=lambda m: self._log_fn(m, "dim"))
            finally:
                motor.quit()  # quit() ya llama CoUninitialize internamente
            self._log_fn(f"Guardado → {dest}", "ok")
            subprocess.Popen(["explorer", "/select,", dest])
        except Exception as e:
            self._log_fn(str(e), "err")
        finally:
            self._busy(False)

    def _crear_arte(self):
        dwg = self._ruta_dwg.get().strip().strip('"')
        if not os.path.isfile(dwg):
            _msgbox_topmost("warning", "Campo requerido", "Selecciona el DWG del plano."); return
        nombre = os.path.splitext(os.path.basename(dwg))[0]
        if _PIPELINE_OK:
            valores = dialogo_cajetin(nombre)
        else:
            valores = {}
        if valores is None: return
        malla        = valores.get("MALLA","").strip()
        pieza        = valores.get("PIEZA","").strip()
        nombre_arte  = valores.get("NOMBRE ARTE","").strip()
        arte0  = _ruta_arte_salida(dwg, malla, pieza, nombre_arte)
        self._busy(True)
        threading.Thread(target=self._t_crear, args=(dwg, valores, arte0), daemon=True).start()

    def _t_crear(self, dwg, valores, arte0):
        self._log_fn("=" * 50)
        self._log_fn("CREAR ARTE...", "ok")
        try:
            _crear_arte_autocad(dwg, log_fn=lambda m: self._log_fn(m,"dim"),
                                valores_cajetin=valores, ruta_salida=arte0,
                                compensar=self._compensar.get())
            self._log_fn(f"Arte completado ✔  {os.path.basename(arte0)}", "ok")

            # Actualizar ruta real en BD
            vitro_bd = valores.get("VITRO","").strip()
            malla_bd = valores.get("MALLA","").strip()
            if vitro_bd or malla_bd:
                try:
                    if _ASIGN_OK:
                        _actualizar_ruta_arte(vitro_bd, malla_bd, arte0)
                        self._log_fn("Ruta guardada en BD ✔", "dim")
                except Exception as _re:
                    self._log_fn(f"WARN ruta BD: {_re}", "warn")
            if self._on_art_done:
                self.after(0, self._on_art_done)
        except ErrorGuardadoArte as e:
            self._log_fn(f"Arte creado, pero falló al guardar: {e}", "warn")
            self.after(0, lambda msg=str(e): _msgbox_topmost("warning", "Arte creado — Error al guardar", msg))
        except Exception as e:
            self._log_fn(f"ERROR: {e}", "err")
            self._log_fn(traceback.format_exc(), "err")
            _anular_bd(valores, self._log_fn)
        finally:
            self._busy(False)

    def _buscar(self):
        if not self._validar(dwg=False): return
        self._busy(True)
        threading.Thread(target=self._t_buscar, daemon=True).start()

    def _t_buscar(self):
        base  = self._ruta_base.get().strip()
        dwg   = self._ruta_dwg.get().strip().strip('"')
        cods  = _extraer_codigos(dwg) if dwg else []
        self._log_fn("BUSCAR ARTE...", "ok")
        res   = _buscar_artes(base, cods)
        self._resultados = res
        self.after(0, self._fill_table, res)
        self._log_fn(f"{len(res)} arte(s) encontrados.", "ok" if res else "warn")
        self._busy(False)

    def _fill_table(self, res):
        for i in self._tree.get_children(): self._tree.delete(i)
        for r in res:
            ext = os.path.splitext(r["archivo"])[1].upper().lstrip(".")
            self._tree.insert("","end", values=(r["version"],r["archivo"],ext), tags=("match",))
        self._lbl_tbl.configure(text=f"{len(res)} resultado{'s' if len(res)!=1 else ''}")

    def _todo_en_uno(self):
        if not self._validar(): return
        dwg    = self._ruta_dwg.get().strip().strip('"')
        nombre = os.path.splitext(os.path.basename(dwg))[0]
        if _PIPELINE_OK:
            valores = dialogo_cajetin(nombre)
        else:
            valores = {}
        if valores is None: return
        malla       = valores.get("MALLA","").strip()
        pieza       = valores.get("PIEZA","").strip()
        nombre_arte = valores.get("NOMBRE ARTE","").strip()
        arte0       = _ruta_arte_salida(dwg, malla, pieza, nombre_arte)
        self._busy(True)
        threading.Thread(target=self._t_todo, args=(dwg, valores, arte0), daemon=True).start()

    def _t_todo(self, dwg, valores, arte0):
        self._log_fn("="*50)
        self._log_fn("TODO EN UNO...", "ok")

        if not _MOTOR_OK or not _COM_OK or not _PIPELINE_OK:
            self._log_fn("Faltan dependencias: AutoCADMotor / pywin32 / pipeline", "err")
            self._busy(False)
            return

        nombre = os.path.splitext(os.path.basename(dwg))[0]
        plano  = os.path.join(_ruta_planos(dwg), f"{nombre}_PLANO.dwg")

        # Un único CoInitialize para todo el proceso en este hilo
        pythoncom.CoInitialize()
        try:
            # ── 1. Extraer plano ──────────────────────────────────────────────
            motor = AutoCADMotor()
            try:
                motor.extraer_layers(dwg, plano, log_fn=lambda m: self._log_fn(m, "dim"))
            finally:
                motor.quit()
            self._log_fn("Plano extraído ✔", "ok")

            # Esperar a que AutoCAD termine de cerrar el documento antes del siguiente open
            time.sleep(2.5)
            for _intento in range(8):
                try:
                    _test = win32com.client.GetActiveObject("AutoCAD.Application")
                    _test.Documents  # verificar que responde
                    break
                except Exception:
                    time.sleep(0.8)

            # ── 2. Crear arte piezas ──────────────────────────────────────────
            n = _crear_arte_autocad(plano, log_fn=lambda m: self._log_fn(m, "dim"),
                                    valores_cajetin=valores, ruta_salida=arte0,
                                    perim_index=0, _com_ya_init=True,
                                    compensar=self._compensar.get())
            self._log_fn(f"Arte guardado ✔  {os.path.basename(arte0)}", "ok")

            # Actualizar ruta real del arte en BD (el cajetin confirmo con ruta="" porque aun no existia)
            vitro_bd = valores.get("VITRO","").strip()
            malla_bd = valores.get("MALLA","").strip()
            if vitro_bd or malla_bd:
                try:
                    if _ASIGN_OK:
                        _actualizar_ruta_arte(vitro_bd, malla_bd, arte0)
                        self._log_fn("Ruta guardada en BD ✔", "dim")
                except Exception as _re:
                    self._log_fn(f"WARN ruta BD: {_re}", "warn")

            _malla_bd = valores.get("MALLA","").strip()
            _pieza_bd = valores.get("PIEZA","").strip()
            for i in range(1, n or 1):
                copia  = plano.replace("_PLANO.dwg", f"_PLANO_p{i+1}.dwg")
                arte_i = _ruta_arte_salida(dwg, _malla_bd, f"{_pieza_bd} {i+1}".strip())
                try:
                    shutil.copy2(plano, copia)
                    _crear_arte_autocad(copia, log_fn=lambda m: self._log_fn(m, "dim"),
                                        valores_cajetin=valores, ruta_salida=arte_i,
                                        perim_index=i, _com_ya_init=True,
                                        compensar=self._compensar.get())
                    self._log_fn(f"Arte {i+1} ✔  {os.path.basename(arte_i)}", "ok")
                except Exception as e:
                    self._log_fn(f"Pieza {i+1}: {e}", "warn")

        except ErrorGuardadoArte as e:
            self._log_fn(f"Arte creado, pero falló al guardar: {e}", "warn")
            self.after(0, lambda msg=str(e): _msgbox_topmost("warning", "Arte creado — Error al guardar", msg))
        except Exception as e:
            self._log_fn(f"ERROR: {e}", "err")
            self._log_fn(traceback.format_exc(), "err")
            _anular_bd(valores, self._log_fn)
            self.after(0, lambda msg=str(e): _msgbox_topmost("error", "Error — Todo en Uno", msg))
        finally:
            pythoncom.CoUninitialize()

        self._log_fn("Todo en Uno completado.", "ok")
        if self._on_art_done:
            self.after(0, self._on_art_done)
        self.after(0, lambda: subprocess.Popen(["explorer", os.path.dirname(arte0)]))
        self._busy(False)

    def _on_doble_click(self, _):
        sel = self._tree.selection()
        if not sel: return
        idx = self._tree.index(sel[0])
        if idx >= len(self._resultados): return
        r   = self._resultados[idx]
        dwg = self._ruta_dwg.get().strip().strip('"')
        if not dwg or not os.path.isfile(dwg):
            _msgbox_topmost("warning", "Plano requerido", "Indica el plano DWG para superponer."); return
        self._log_fn(f"Superponiendo: {r['archivo']}", "ok")
        self._busy(True)
        threading.Thread(target=self._t_overlay,
                         args=(r["ruta_completa"], dwg), daemon=True).start()

    def _t_overlay(self, arte, plano):
        if not _COM_OK:
            self._log_fn("pywin32 no disponible", "err")
            self._busy(False)
            return
        pythoncom.CoInitialize()
        try:
            from arte_maker import _overlay_autocad
            _overlay_autocad(arte, plano, log_fn=lambda m: self._log_fn(m, "dim"))
            self._log_fn("Superposición lista en AutoCAD.", "ok")
        except Exception as e:
            self._log_fn(str(e), "err")
        finally:
            pythoncom.CoUninitialize()
            self._busy(False)


# ══════════════════════════════════════════════════════════════════════════════
#  PESTAÑA — CONSULTAR BD
# ══════════════════════════════════════════════════════════════════════════════
class TabBD(ctk.CTkFrame):
    TABS = [
        ("Vitrojet",    "vitrojet",       "🔬"),
        ("Mallas G",    "grandes",        "🔷"),
        ("Mallas P",    "pequenas",       "🔹"),
        ("Vinilos",     "vinilos",        "🎨"),
        ("Pasta Plata", "pasta",          "🪙"),
        ("Glassjet V.", "glassjet_viejo", "📦"),
    ]
    STATS = [
        ("mallas_grandes",  "Mallas G",   "🔷", PAL["accent"]),
        ("mallas_pequenas", "Mallas P",   "🔹", PAL["purple"]),
        ("vitrojet",        "Vitrojet",   "🔬", "#0ea5e9"),
        ("pasta_plata",     "Pasta Plata","🪙", "#f59e0b"),
        ("glassjet_viejo",  "Glassjet V.","📦", PAL["purple"]),
        ("vinilos",         "Vinilos",    "🎨", "#ec4899"),
    ]
    QUERIES = {
        "vitrojet":  ("SELECT TOP(?) v.vitro,v.codigo_malla,v.tipo_malla,v.bnerig,v.vehiculo,v.version,v.ruta,v.responsable "
                      "FROM mallas.vitrojet v {where} ORDER BY TRY_CAST(SUBSTRING(v.vitro,3,50) AS INT) DESC",
                      ["Vitro","Malla","Tipo","B/N","Vehículo","Versión","Ruta","Responsable"],
                      ["vitro","codigo_malla","tipo_malla","bnerig","vehiculo","version","ruta","responsable"]),
        "grandes":   ("SELECT TOP(?) codigo,cod_veh,descripcion,pieza,tipo,version,ruta_dwg,responsable "
                      "FROM mallas.grandes {where} ORDER BY TRY_CAST(SUBSTRING(codigo,3,50) AS INT) DESC",
                      ["Código","Cód.Veh.","Descripción","Pieza","Tipo","Versión","Ruta","Responsable"],
                      ["codigo","cod_veh","descripcion","pieza","tipo","version","ruta_dwg","responsable"]),
        "pequenas":  ("SELECT TOP(?) codigo,cod_veh,descripcion,pieza,tipo,version,ruta_dwg,responsable "
                      "FROM mallas.pequenas {where} ORDER BY TRY_CAST(codigo AS INT) DESC",
                      ["Código","Cód.Veh.","Descripción","Pieza","Tipo","Versión","Ruta","Responsable"],
                      ["codigo","cod_veh","descripcion","pieza","tipo","version","ruta_dwg","responsable"]),
        "vinilos":   ("SELECT TOP(?) herramental,vehiculo,cod_vehiculo,version,pieza,tipo "
                      "FROM mallas.vinilos {where} ORDER BY herramental DESC",
                      ["Herramental","Vehículo","Cód.Veh.","Versión","Pieza","Tipo"],
                      ["herramental","vehiculo","cod_vehiculo","version","pieza","tipo"]),
        "pasta":     ("SELECT TOP(?) consecutivo,tipo,vehiculo,cod_vehiculo,version,pieza,ruta_archivo,caso "
                      "FROM mallas.pasta_plata {where} ORDER BY consecutivo DESC",
                      ["Consecutivo","RED/ANT","Vehículo","Cód.Veh.","Versión","Pieza","Ruta Archivo","Caso"],
                      ["consecutivo","tipo","vehiculo","cod_vehiculo","version","pieza","ruta_archivo","caso"]),
        "glassjet_viejo": ("SELECT TOP(?) id,malla,glassjet,part_number,tipo,vehiculo,homologacion_vitro "
                           "FROM mallas.glassjet_viejo {where} ORDER BY id DESC",
                           ["ID","Malla","Glassjet","Part Number","Tipo","Vehículo","Homol. Vitro"],
                           ["id","malla","glassjet","part_number","tipo","vehiculo","homologacion_vitro"]),
    }
    WHERE = {
        "vitrojet":       "WHERE v.vitro LIKE ? OR v.vehiculo LIKE ? OR v.codigo_malla LIKE ?",
        "grandes":        "WHERE descripcion LIKE ? OR codigo LIKE ? OR cod_veh LIKE ?",
        "pequenas":       "WHERE descripcion LIKE ? OR CAST(codigo AS NVARCHAR) LIKE ? OR cod_veh LIKE ?",
        "vinilos":        "WHERE vehiculo LIKE ? OR herramental LIKE ?",
        "pasta":          "WHERE vehiculo LIKE ? OR consecutivo LIKE ?",
        "glassjet_viejo": "WHERE malla LIKE ? OR glassjet LIKE ? OR vehiculo LIKE ?",
    }

    def __init__(self, parent, **kw):
        super().__init__(parent, fg_color="transparent", **kw)
        self._tab    = "vitrojet"
        self._timer  = None
        self._cards  = {}
        self._sync_running = False
        self._build()
        self.after(300, self._load_stats)

    def _build(self):
        self.columnconfigure(0, weight=1)
        self.rowconfigure(3, weight=1)

        # ── Stats ─────────────────────────────────────────────────────────────
        stats_frame = ctk.CTkFrame(self, fg_color="transparent")
        stats_frame.grid(row=0, column=0, sticky="ew", padx=4, pady=(4,0))
        for i, (key, lbl, icon, color) in enumerate(self.STATS):
            stats_frame.columnconfigure(i, weight=1)
            c = StatCard(stats_frame, lbl, icon, color)
            c.grid(row=0, column=i, padx=4, pady=6, sticky="ew")
            self._cards[key] = c

        # ── Barra búsqueda + tabs — 2 filas para que no se corte ─────────────
        bar = ctk.CTkFrame(self, fg_color=PAL["card"], corner_radius=10,
                           border_width=1, border_color=PAL["border"])
        bar.grid(row=1, column=0, sticky="ew", padx=4, pady=4)
        bar_in = ctk.CTkFrame(bar, fg_color="transparent")
        bar_in.pack(fill="x", padx=12, pady=(8,8))
        bar_in.columnconfigure(0, weight=1)

        # Fila 0: Tabs
        tab_bar = ctk.CTkFrame(bar_in, fg_color=PAL["bg"], corner_radius=8)
        tab_bar.grid(row=0, column=0, sticky="w", pady=(0,8))
        self._tab_btns = {}
        for i, (lbl, key, icon) in enumerate(self.TABS):
            b = ctk.CTkButton(tab_bar, text=f"{icon} {lbl}", width=120, height=36,
                              corner_radius=8, font=FONT(12),
                              fg_color=PAL["accent2"] if key=="vitrojet" else "transparent",
                              hover_color=PAL["border"],
                              command=lambda k=key: self._set_tab(k))
            b.grid(row=0, column=i, padx=2, pady=3)
            self._tab_btns[key] = b

        # Fila 1: Búsqueda
        srch = ctk.CTkFrame(bar_in, fg_color="transparent")
        srch.grid(row=1, column=0, sticky="ew")
        srch.columnconfigure(1, weight=1)
        ctk.CTkLabel(srch, text="🔍", font=FONT(14)).grid(row=0, column=0, padx=(0,6))
        self._search = ctk.CTkEntry(srch, placeholder_text="Buscar vehículo, código, malla...",
                                     height=40, font=FONT(13),
                                     fg_color=PAL["card2"], border_color=PAL["border"])
        self._search.grid(row=0, column=1, sticky="ew")
        self._search.bind("<KeyRelease>", self._on_key)

        # ── Panel Sincronizar Excel ───────────────────────────────────────────
        sync_card = ctk.CTkFrame(self, fg_color=PAL["card"], corner_radius=10,
                                  border_width=1, border_color=PAL["border"])
        sync_card.grid(row=2, column=0, sticky="ew", padx=4, pady=(0,4))
        sync_in = ctk.CTkFrame(sync_card, fg_color="transparent")
        sync_in.pack(fill="x", padx=12, pady=8)
        sync_in.columnconfigure(1, weight=1)

        self._btn_sync = ctk.CTkButton(
            sync_in, text="⟳  Sincronizar Excel → BD", width=200, height=38,
            font=FONT(12, "bold"), corner_radius=8,
            fg_color=PAL["green2"], hover_color="#1e5c36",
            command=self._sync_excel)
        self._btn_sync.grid(row=0, column=0, padx=(0,10))

        self._sync_prog = ctk.CTkProgressBar(sync_in, mode="indeterminate",
                                              height=4, progress_color=PAL["green"])
        self._sync_prog.grid(row=0, column=1, sticky="ew", padx=(0,10))

        self._lbl_sync = ctk.CTkLabel(
            sync_in, text="Actualiza la BD desde el Excel compartido en OneDrive",
            font=FONT(10), text_color=PAL["txt_mid"])
        self._lbl_sync.grid(row=0, column=2)

        self._sync_log = LogBox(sync_card, height=80)
        self._sync_log.pack(fill="x", padx=12, pady=(0,8))

        # ── Tabla ─────────────────────────────────────────────────────────────
        import tkinter.ttk as ttv
        card_tbl = ctk.CTkFrame(self, fg_color=PAL["card"], corner_radius=10,
                                 border_width=1, border_color=PAL["border"])
        card_tbl.grid(row=3, column=0, sticky="nsew", padx=4, pady=4)

        frm = ctk.CTkFrame(card_tbl, fg_color="transparent")
        frm.pack(fill="both", expand=True, padx=8, pady=(8,0))
        frm.grid_rowconfigure(0, weight=1)
        frm.grid_columnconfigure(0, weight=1)

        self._tree = ttv.Treeview(frm, style="AGP.Treeview", show="headings", height=18)
        self._tree.grid(row=0, column=0, sticky="nsew")
        sb = ttv.Scrollbar(frm, orient="vertical", command=self._tree.yview)
        sb.grid(row=0, column=1, sticky="ns")
        sb_x = ttv.Scrollbar(frm, orient="horizontal", command=self._tree.xview)
        sb_x.grid(row=1, column=0, sticky="ew")
        self._tree.configure(yscrollcommand=sb.set, xscrollcommand=sb_x.set)

        self._lbl_count = ctk.CTkLabel(card_tbl, text="",
                                        font=FONT(10), text_color=PAL["txt_dim"])
        self._lbl_count.pack(anchor="e", padx=12, pady=(0,6))

        self._build_tree_cols("vitrojet")
        _setup_tree_copy(self._tree, self._lbl_count)
        # No buscar en frío — se carga cuando el usuario entra a la pestaña

    _COL_W = {
        # Vitrojet
        "Vitro": 90, "Malla": 90, "Tipo": 55, "B/N": 55,
        "Vehículo": 160, "Versión": 90, "Ruta": 260,
        # Grandes / Pequeñas / Vinilos
        "Código": 85, "Cód.Veh.": 80, "Descripción": 160,
        "Pieza": 100, "Tipo malla": 60, "Herramental": 100,
        # Pasta
        "Consecutivo": 95, "Ruta Archivo": 260, "Caso": 90,
        # Glassjet Viejo
        "ID": 55, "Glassjet": 100, "Part Number": 130, "Homol. Vitro": 110,
        # Común
        "Responsable": 120,
    }
    def _build_tree_cols(self, tab):
        _, headers, _ = self.QUERIES[tab]
        self._tree.configure(columns=headers)
        for h in headers:
            self._tree.heading(h, text=h, anchor="w")
            w = self._COL_W.get(h, 100)
            self._tree.column(h, width=w, minwidth=55, stretch=False, anchor="w")

    def _set_tab(self, key):
        self._tab = key
        for k, b in self._tab_btns.items():
            b.configure(fg_color=PAL["accent2"] if k==key else "transparent")
        self._build_tree_cols(key)
        self._do_search()

    def _on_key(self, _):
        if self._timer: self.after_cancel(self._timer)
        self._timer = self.after(300, self._do_search)

    def _do_search(self):
        threading.Thread(target=self._t_search, daemon=True).start()

    def _t_search(self):
        q   = self._search.get().strip()
        tab = self._tab
        sql_tpl, headers, fields = self.QUERIES[tab]
        limit = 300
        # Limpiar tree siempre antes de la query — nunca mostrar datos de otro tab
        self.after(0, lambda: [self._tree.delete(i) for i in self._tree.get_children()])
        try:
            if q:
                where    = self.WHERE[tab]
                n_params = where.count("?")
                like     = f"%{q}%"
                params   = (limit,) + (like,) * n_params
                sql      = sql_tpl.format(where=where)
            else:
                params = (limit,)
                sql    = sql_tpl.format(where="")
            rows = db_query(sql, params)
        except Exception as e:
            err = str(e)
            self.after(0, lambda m=err: (
                [self._tree.delete(i) for i in self._tree.get_children()],
                self._tree.insert("", "end", values=(f"⚠  {m[:120]}",) + ("",) * (len(self._tree["columns"])-1)),
                self._lbl_count.configure(text="Sin conexión a BD", text_color=PAL["red"])
            ))
            return
        self.after(0, self._fill, rows, fields, headers)

    def _fill(self, rows, fields, headers):
        for i in self._tree.get_children(): self._tree.delete(i)
        for r in rows:
            vals = [str(r.get(f,"") or "—").replace("\r\n"," ").replace("\n"," ").replace("\r"," ") for f in fields]
            self._tree.insert("","end", values=vals)
        n = len(rows)
        self._lbl_count.configure(
            text=f"{n} resultado{'s' if n!=1 else ''}  (máx. 300 — busca para filtrar)",
            text_color=PAL["txt_dim"])

    def _load_stats(self):
        def _worker():
            try:
                conn = db_connect()
                cur  = conn.cursor()
                tabla_map = {
                    "mallas_grandes":  "mallas.grandes",
                    "mallas_pequenas": "mallas.pequenas",
                    "vitrojet":        "mallas.vitrojet",
                    "pasta_plata":     "mallas.pasta_plata",
                    "glassjet_viejo":  "mallas.glassjet_viejo",
                    "vinilos":         "mallas.vinilos",
                }
                for key, tabla in tabla_map.items():
                    cur.execute(f"SELECT COUNT(*) FROM {tabla}")
                    n = cur.fetchone()[0]
                    self.after(0, lambda k=key, v=n: self._cards[k].set(v))
                conn.close()
            except Exception:
                pass
        threading.Thread(target=_worker, daemon=True).start()

    # ── Sincronizar Excel ─────────────────────────────────────────────────────
    def _sync_excel(self):
        if self._sync_running:
            return
        self._sync_running = True
        self._btn_sync.configure(state="disabled", text="⟳  Sincronizando...")
        self._sync_prog.start()
        self._sync_log.clear()
        threading.Thread(target=self._t_sync, daemon=True).start()

    def _t_sync(self):
        import sys, os
        sys.path.insert(0, os.path.join(os.path.dirname(__file__), "db_app"))
        try:
            import importlib
            import db_app.importar_excel as imp
            importlib.reload(imp)
            ok = imp.main(log_fn=lambda m, tag="": self.after(0, self._sync_log.write, m, tag))
        except Exception as e:
            self.after(0, self._sync_log.write, f"Error al importar módulo: {e}", "err")
            ok = False
        finally:
            self.after(0, self._sync_done, ok)

    def _sync_done(self, ok):
        self._sync_running = False
        self._sync_prog.stop()
        self._sync_prog.set(0)
        self._btn_sync.configure(state="normal", text="⟳  Sincronizar Excel → BD")
        self._lbl_sync.configure(
            text="✔ Sincronización completada" if ok else "✘ Hubo errores — revisa el log",
            text_color=PAL["green"] if ok else PAL["orange"])
        self._load_stats()
        self._do_search()

    def refresh(self):
        """Refresca stats y tabla (llamado desde otras pestañas tras cambios)."""
        self._load_stats()
        self._do_search()


# ══════════════════════════════════════════════════════════════════════════════
#  PESTAÑA — GESTIÓN BD (editar / separar)
# ══════════════════════════════════════════════════════════════════════════════
class TabGestion(ctk.CTkFrame):
    TABS = [
        ("Vitrojet",   "vitrojet",    "🔬"),
        ("Mallas G",   "grandes",     "🔷"),
        ("Mallas P",   "pequenas",    "🔹"),
        ("Vinilos",    "vinilos",     "🎨"),
        ("Pasta Plata","pasta_plata", "🪙"),
    ]
    QUERIES = {
        "vitrojet": (
            "SELECT TOP(?) v.vitro,v.codigo_malla,v.tipo_malla,v.bnerig,v.vehiculo,v.version,"
            "COALESCE(g.pieza,p.pieza) AS pieza,v.ruta,v.responsable,v.estado,v.modificado_por,v.modificado_en "
            "FROM mallas.vitrojet v "
            "LEFT JOIN mallas.grandes g ON v.codigo_malla=g.codigo "
            "LEFT JOIN mallas.pequenas p ON v.codigo_malla=CAST(p.codigo AS NVARCHAR) "
            "{where} ORDER BY TRY_CAST(SUBSTRING(v.vitro,3,50) AS INT) DESC",
            ["Vitro","Malla","Tipo","B/N","Vehículo","Versión","Pieza","Ruta","Responsable","Estado","Modificado por","Modificado en"],
            ["vitro","codigo_malla","tipo_malla","bnerig","vehiculo","version","pieza","ruta","responsable","estado","modificado_por","modificado_en"],
            "vitro",
        ),
        "grandes": (
            "SELECT TOP(?) codigo,cod_veh,descripcion,pieza,tipo,version,ruta_dwg,responsable,estado,modificado_por,modificado_en "
            "FROM mallas.grandes {where} ORDER BY TRY_CAST(SUBSTRING(codigo,3,50) AS INT) DESC",
            ["Código","Cód.Veh.","Descripción","Pieza","Tipo","Versión","Ruta","Responsable","Estado","Modificado por","Modificado en"],
            ["codigo","cod_veh","descripcion","pieza","tipo","version","ruta_dwg","responsable","estado","modificado_por","modificado_en"],
            "codigo",
        ),
        "pequenas": (
            "SELECT TOP(?) codigo,cod_veh,descripcion,pieza,tipo,version,ruta_dwg,responsable,estado,modificado_por,modificado_en "
            "FROM mallas.pequenas {where} ORDER BY TRY_CAST(codigo AS INT) DESC",
            ["Código","Cód.Veh.","Descripción","Pieza","Tipo","Versión","Ruta","Responsable","Estado","Modificado por","Modificado en"],
            ["codigo","cod_veh","descripcion","pieza","tipo","version","ruta_dwg","responsable","estado","modificado_por","modificado_en"],
            "codigo",
        ),
        "vinilos": (
            "SELECT TOP(?) herramental,vehiculo,cod_vehiculo,version,pieza,tipo,ruta,modificado_por,modificado_en "
            "FROM mallas.vinilos {where} ORDER BY herramental DESC",
            ["Herramental","Vehículo","Cód.Veh.","Versión","Pieza","Tipo","Ruta","Modificado por","Modificado en"],
            ["herramental","vehiculo","cod_vehiculo","version","pieza","tipo","ruta","modificado_por","modificado_en"],
            "herramental",
        ),
        "pasta_plata": (
            "SELECT TOP(?) consecutivo,tipo,vehiculo,cod_vehiculo,version,pieza,ruta_archivo,caso,modificado_por,modificado_en "
            "FROM mallas.pasta_plata {where} ORDER BY consecutivo DESC",
            ["Consecutivo","RED/ANT","Vehículo","Cód.Veh.","Versión","Pieza","Ruta archivo","Caso","Modificado por","Modificado en"],
            ["consecutivo","tipo","vehiculo","cod_vehiculo","version","pieza","ruta_archivo","caso","modificado_por","modificado_en"],
            "consecutivo",
        ),
    }
    WHERE = {
        "vitrojet":   "WHERE v.vitro LIKE ? OR v.vehiculo LIKE ? OR v.codigo_malla LIKE ?",
        "grandes":    "WHERE descripcion LIKE ? OR codigo LIKE ? OR cod_veh LIKE ?",
        "pequenas":   "WHERE descripcion LIKE ? OR CAST(codigo AS NVARCHAR) LIKE ? OR cod_veh LIKE ?",
        "vinilos":    "WHERE vehiculo LIKE ? OR herramental LIKE ? OR pieza LIKE ?",
        "pasta_plata":"WHERE vehiculo LIKE ? OR consecutivo LIKE ? OR pieza LIKE ?",
    }
    EDITABLES = {
        "vitrojet": [
            ("Malla",       "codigo_malla"),
            ("B/N",         "bnerig"),
            ("Vehículo",    "vehiculo"),
            ("Versión",     "version"),
            ("Pieza",       "_pieza_malla"),   # columna virtual → cascade a malla linked
            ("Ruta",        "ruta"),
            ("Responsable", "responsable"),
        ],
        "grandes": [
            ("Cód.Veh.",    "cod_veh"),
            ("Descripción", "descripcion"),
            ("Pieza",       "pieza"),
            ("Tipo",        "tipo"),
            ("Versión",     "version"),
            ("Ruta",        "ruta_dwg"),
            ("Responsable", "responsable"),
        ],
        "pequenas": [
            ("Cód.Veh.",    "cod_veh"),
            ("Descripción", "descripcion"),
            ("Pieza",       "pieza"),
            ("Tipo",        "tipo"),
            ("Versión",     "version"),
            ("Ruta",        "ruta_dwg"),
            ("Responsable", "responsable"),
        ],
        "vinilos": [
            ("Vehículo",     "vehiculo"),
            ("Cód. Vehículo","cod_vehiculo"),
            ("Versión",      "version"),
            ("Pieza",        "pieza"),
            ("Tipo",         "tipo"),
            ("Ruta",         "ruta"),
        ],
        "pasta_plata": [
            ("RED/ANT",       "tipo"),
            ("Vehículo",      "vehiculo"),
            ("Cód. Vehículo", "cod_vehiculo"),
            ("Versión",       "version"),
            ("Pieza",         "pieza"),
            ("Ruta archivo",  "ruta_archivo"),
            ("Caso",          "caso"),
        ],
    }
    _TABLA = {
        "vitrojet":   "mallas.vitrojet",
        "grandes":    "mallas.grandes",
        "pequenas":   "mallas.pequenas",
        "vinilos":    "mallas.vinilos",
        "pasta_plata":"mallas.pasta_plata",
    }
    _PK = {
        "vitrojet":   "vitro",
        "grandes":    "codigo",
        "pequenas":   "codigo",
        "vinilos":    "herramental",
        "pasta_plata":"consecutivo",
    }
    _COL_W = {
        "Vitro":90, "Código":85, "Herramental":100, "Consecutivo":110,
        "Malla":90, "Cód.Veh.":75, "Tipo":60, "B/N":55,
        "Vehículo":170, "Versión":80, "Descripción":160,
        "Pieza":110, "Ruta":260, "Ruta archivo":260,
        "Responsable":120, "Estado":85, "Caso":80,
        "Modificado por":140, "Modificado en":130,
    }

    def __init__(self, parent, usuario_info=None, **kw):
        super().__init__(parent, fg_color="transparent", **kw)
        self._usuario         = usuario_info or {}
        self._tab             = "vitrojet"
        self._timer           = None
        self._rows            = []
        self._on_data_changed = None
        self._build()

    def _build(self):
        self.columnconfigure(0, weight=1)
        self.rowconfigure(2, weight=1)

        # ── Barra superior — 2 filas para que no se corte ─────────────────────
        top = ctk.CTkFrame(self, fg_color=PAL["card"], corner_radius=10,
                           border_width=1, border_color=PAL["border"])
        top.grid(row=0, column=0, sticky="ew", padx=4, pady=(4,2))
        top_in = ctk.CTkFrame(top, fg_color="transparent")
        top_in.pack(fill="x", padx=14, pady=(10,6))
        top_in.columnconfigure(0, weight=1)

        # Fila 0: Tabs
        tab_row = ctk.CTkFrame(top_in, fg_color=PAL["bg"], corner_radius=8)
        tab_row.grid(row=0, column=0, sticky="w", pady=(0,8))
        self._tab_btns = {}
        for i, (lbl, key, icon) in enumerate(self.TABS):
            b = ctk.CTkButton(tab_row, text=f"{icon} {lbl}", width=110, height=36,
                              corner_radius=8, font=FONT(11),
                              fg_color=PAL["accent2"] if key=="vitrojet" else "transparent",
                              hover_color=PAL["border"],
                              command=lambda k=key: self._set_tab(k))
            b.grid(row=0, column=i, padx=2, pady=3)
            self._tab_btns[key] = b

        # Fila 1: botones de acción + búsqueda
        row1 = ctk.CTkFrame(top_in, fg_color="transparent")
        row1.grid(row=1, column=0, sticky="ew")
        row1.columnconfigure(1, weight=1)

        btn_grp = ctk.CTkFrame(row1, fg_color="transparent")
        btn_grp.grid(row=0, column=0, padx=(0,12))

        self._btn_sep = ctk.CTkButton(
            btn_grp, text="＋  Separar vitro / malla", width=185, height=38,
            font=FONT(11, "bold"), corner_radius=8,
            fg_color=PAL["orange"], hover_color="#b35c00",
            command=self._separar)
        self._btn_sep.pack(side="left", padx=(0,8))

        ctk.CTkButton(
            btn_grp, text="📝  Insertar manual", width=150, height=38,
            font=FONT(11, "bold"), corner_radius=8,
            fg_color=PAL["green2"], hover_color="#1e5c36",
            command=self._insertar_manual).pack(side="left")

        srch_grp = ctk.CTkFrame(row1, fg_color="transparent")
        srch_grp.grid(row=0, column=1, sticky="ew")
        srch_grp.columnconfigure(1, weight=1)
        ctk.CTkLabel(srch_grp, text="🔍", font=FONT(14)).grid(row=0, column=0, padx=(0,6))
        self._search = ctk.CTkEntry(
            srch_grp, placeholder_text="Buscar vehículo, código, malla...",
            height=38, font=FONT(13),
            fg_color=PAL["card2"], border_color=PAL["border"])
        self._search.grid(row=0, column=1, sticky="ew")
        self._search.bind("<KeyRelease>", self._on_key)

        # ── Hint edición ──────────────────────────────────────────────────────
        hint = ctk.CTkFrame(self, fg_color="transparent")
        hint.grid(row=1, column=0, sticky="ew", padx=14, pady=(0,2))
        ctk.CTkLabel(hint, text="✏  Doble clic sobre una fila para editar · Pestañas Vitrojet/Mallas: ver y editar asignaciones · Botón Insertar: agregar pasta de plata, vinilos o glassjet",
                     font=FONT(10), text_color=PAL["txt_dim"]).pack(anchor="w")

        # ── Tabla ─────────────────────────────────────────────────────────────
        import tkinter.ttk as ttv
        card_tbl = ctk.CTkFrame(self, fg_color=PAL["card"], corner_radius=10,
                                border_width=1, border_color=PAL["border"])
        card_tbl.grid(row=2, column=0, sticky="nsew", padx=4, pady=4)

        frm = ctk.CTkFrame(card_tbl, fg_color="transparent")
        frm.pack(fill="both", expand=True, padx=8, pady=(8,0))
        frm.grid_rowconfigure(0, weight=1)
        frm.grid_columnconfigure(0, weight=1)

        self._tree = ttv.Treeview(frm, style="AGP.Treeview", show="headings", height=22)
        self._tree.grid(row=0, column=0, sticky="nsew")
        sb = ttv.Scrollbar(frm, orient="vertical", command=self._tree.yview)
        sb.grid(row=0, column=1, sticky="ns")
        sb_x = ttv.Scrollbar(frm, orient="horizontal", command=self._tree.xview)
        sb_x.grid(row=1, column=0, sticky="ew")
        self._tree.configure(yscrollcommand=sb.set, xscrollcommand=sb_x.set)
        self._tree.bind("<Double-1>", self._on_doble_click)
        self._tree.bind("<Return>",   self._on_doble_click)

        self._lbl_count = ctk.CTkLabel(card_tbl, text="",
                                       font=FONT(10), text_color=PAL["txt_dim"])
        self._lbl_count.pack(anchor="e", padx=12, pady=(0,6))

        self._build_cols("vitrojet")
        _setup_tree_copy(self._tree, self._lbl_count)
        # No llamar _do_search() aquí — se llama cuando el usuario entra a la pestaña

    def _build_cols(self, tab):
        headers = self.QUERIES[tab][1]
        self._tree.configure(columns=headers)
        for h in headers:
            self._tree.heading(h, text=h, anchor="w")
            w = self._COL_W.get(h, 100)
            self._tree.column(h, width=w, minwidth=50, stretch=False, anchor="w")

    def _set_tab(self, key):
        self._tab = key
        for k, b in self._tab_btns.items():
            b.configure(fg_color=PAL["accent2"] if k==key else "transparent")
        self._build_cols(key)
        self._do_search()

    def _on_key(self, _):
        if self._timer: self.after_cancel(self._timer)
        self._timer = self.after(300, self._do_search)

    def _do_search(self):
        threading.Thread(target=self._t_search, daemon=True).start()

    def _t_search(self):
        q   = self._search.get().strip()
        tab = self._tab
        sql_tpl, headers, fields, _ = self.QUERIES[tab]
        limit = 300
        self.after(0, lambda: [self._tree.delete(i) for i in self._tree.get_children()])
        try:
            if q:
                where    = self.WHERE[tab]
                n_params = where.count("?")
                params   = (limit,) + (f"%{q}%",) * n_params
                sql      = sql_tpl.format(where=where)
            else:
                params = (limit,)
                sql    = sql_tpl.format(where="")
            rows = db_query(sql, params)
        except Exception as e:
            err = str(e)
            self.after(0, lambda m=err: (
                [self._tree.delete(i) for i in self._tree.get_children()],
                self._tree.insert("", "end", values=(f"⚠  {m[:120]}",) + ("",) * (len(self._tree["columns"])-1)),
                self._lbl_count.configure(text="Sin conexión a BD", text_color=PAL["red"])
            ))
            return
        self.after(0, self._fill, rows, fields, headers)

    def _fill(self, rows, fields, headers):
        self._rows = rows
        for i in self._tree.get_children(): self._tree.delete(i)
        for r in rows:
            vals = []
            for f in fields:
                v = r.get(f, "") or ""
                v = str(v).replace("\r\n", " ").replace("\n", " ").replace("\r", " ") if v else ""
                vals.append(v)
            self._tree.insert("", "end", values=vals)
        n = len(rows)
        self._lbl_count.configure(
            text=f"{n} resultado{'s' if n!=1 else ''}  (máx. 300 — busca para filtrar)",
            text_color=PAL["txt_dim"])

    def _on_doble_click(self, event):
        item = self._tree.identify_row(event.y) if hasattr(event, "y") else None
        sel  = self._tree.selection()
        iid  = item or (sel[0] if sel else None)
        if not iid: return
        try:
            idx = self._tree.index(iid)
        except Exception:
            return
        if idx >= len(self._rows): return
        self._abrir_editor(self._rows[idx])

    def _abrir_editor(self, fila):
        import tkinter as _tk
        tab       = self._tab
        editables = self.EDITABLES[tab]
        pk_col    = self._PK[tab]
        pk_val    = fila.get(pk_col, "")
        tabla     = self._TABLA[tab]

        try:
            root_tk = _tk._default_root
        except Exception:
            root_tk = None

        win = _tk.Toplevel(root_tk)
        win.title(f"Editar — {pk_val}")
        win.configure(bg="#1e1e2e")
        win.resizable(False, True)
        win.attributes("-topmost", True)
        if root_tk:
            win.grab_set()

        # Header
        hdr = _tk.Frame(win, bg=PAL["accent"], height=4)
        hdr.pack(fill="x")
        _tk.Frame(win, bg="#1e1e2e", height=10).pack(fill="x")
        _tk.Label(win, text="✏  Editar registro",
                  font=("Segoe UI", 15, "bold"), fg="#e2e8f0", bg="#1e1e2e"
                  ).pack(anchor="w", padx=24, pady=(0,2))
        _tk.Label(win, text=f"{pk_col.upper()}:  {pk_val}",
                  font=("Segoe UI", 10), fg=PAL["accent"], bg="#1e1e2e"
                  ).pack(anchor="w", padx=24, pady=(0,12))

        # Línea separadora
        _tk.Frame(win, bg="#2d2d44", height=1).pack(fill="x", padx=18, pady=(0,14))

        entries_edit = {}
        for lbl_txt, campo in editables:
            row_f = _tk.Frame(win, bg="#1e1e2e")
            row_f.pack(fill="x", padx=24, pady=4)
            _tk.Label(row_f, text=lbl_txt, width=14, anchor="e",
                      font=("Segoe UI", 9), fg="#94a3b8", bg="#1e1e2e"
                      ).pack(side="left", padx=(0,10))
            # _pieza_malla es campo virtual → leer de columna "pieza" del JOIN
            val = fila.get("pieza" if campo == "_pieza_malla" else campo, "") or ""
            _DROPDOWNS = {
                "caso":   (["Caso 1", "Caso 2"],  "Caso 1"),
                "bnerig": (["BN", "BNI"],          "BN"),
            }
            _tipo_pasta = (campo == "tipo" and tab == "pasta_plata")
            if campo in _DROPDOWNS or _tipo_pasta:
                opciones, default = (["RED", "ANT"], "RED") if _tipo_pasta else _DROPDOWNS[campo]
                cur_val = str(val) if val in opciones else default
                var = _tk.StringVar(value=cur_val)
                opt = _tk.OptionMenu(row_f, var, *opciones)
                opt.config(bg="#2d2d44", fg="#e2e8f0", activebackground="#3d4f7c",
                           activeforeground="#e2e8f0", relief="flat", bd=0,
                           font=("Segoe UI", 10), width=10, highlightthickness=0)
                opt["menu"].config(bg="#2d2d44", fg="#e2e8f0",
                                   activebackground="#3d4f7c", font=("Segoe UI", 10))
                opt.pack(side="left", ipady=3)
                entries_edit[campo] = var
            else:
                ent = _tk.Entry(row_f, width=42,
                                font=("Segoe UI", 10),
                                bg="#2d2d44", fg="#e2e8f0",
                                insertbackground="#e2e8f0",
                                relief="flat", bd=6)
                ent.pack(side="left", ipady=4)
                ent.insert(0, str(val))
                entries_edit[campo] = ent

        _tk.Frame(win, bg="#1e1e2e", height=6).pack()
        _tk.Frame(win, bg="#2d2d44", height=1).pack(fill="x", padx=18, pady=(0,12))

        # Estado (solo lectura, info)
        estado_val = fila.get("estado", "") or "—"
        color_est  = {"ASIGNADO": "#22c55e", "PENDIENTE": "#f59e0b",
                      "CANCELADO": "#ef4444"}.get(estado_val, "#94a3b8")
        est_f = _tk.Frame(win, bg="#1e1e2e")
        est_f.pack(fill="x", padx=24, pady=(0,16))
        _tk.Label(est_f, text="Estado", width=14, anchor="e",
                  font=("Segoe UI", 9), fg="#94a3b8", bg="#1e1e2e"
                  ).pack(side="left", padx=(0,10))
        _tk.Label(est_f, text=f"  {estado_val}  ",
                  font=("Segoe UI", 9, "bold"), fg=color_est, bg="#252538",
                  relief="flat", bd=4
                  ).pack(side="left")

        # Botones
        btn_f = _tk.Frame(win, bg="#1e1e2e")
        btn_f.pack(fill="x", padx=24, pady=(0,18))

        msg_lbl = _tk.Label(btn_f, text="", font=("Segoe UI", 9),
                            fg="#22c55e", bg="#1e1e2e")
        msg_lbl.pack(anchor="w", pady=(0,8))

        def _guardar():
            nuevos_raw = {campo: (ent.get().strip() or None) for campo, ent in entries_edit.items()}
            if not nuevos_raw: return
            # Separar campos virtuales de los reales
            pieza_virtual = nuevos_raw.pop("_pieza_malla", _SENTINEL)
            nuevos = nuevos_raw
            usuario_nombre = self._usuario.get("nombre") or self._usuario.get("usuario") or "desconocido"
            try:
                cn  = db_connect()
                cur = cn.cursor()

                # Leer valores actuales para trazabilidad (solo columnas reales de la tabla)
                if nuevos:
                    campos_sel = ", ".join(nuevos.keys())
                    cur.execute(f"SELECT {campos_sel} FROM {tabla} WHERE {pk_col}=?", (pk_val,))
                    row_ant = cur.fetchone()
                    valores_ant = {}
                    if row_ant:
                        for i, campo in enumerate(nuevos.keys()):
                            valores_ant[campo] = str(row_ant[i]) if row_ant[i] is not None else None

                    # UPDATE con modificado_por y modificado_en
                    sets   = [f"{c}=?" for c in nuevos] + ["modificado_por=?", "modificado_en=SYSDATETIME()"]
                    params = list(nuevos.values()) + [usuario_nombre, pk_val]
                    cur.execute(f"UPDATE {tabla} SET {', '.join(sets)} WHERE {pk_col}=?", params)
                else:
                    valores_ant = {}

                # INSERT en trazabilidad por cada campo que cambió
                for campo, val_nuevo in nuevos.items():
                    val_ant = valores_ant.get(campo)
                    val_nuevo_s = str(val_nuevo) if val_nuevo is not None else None
                    if val_ant != val_nuevo_s:
                        cur.execute(
                            "INSERT INTO MALLAS.TRAZABILIDAD "
                            "(tabla, pk_campo, pk_valor, campo, valor_anterior, valor_nuevo, usuario) "
                            "VALUES (?,?,?,?,?,?,?)",
                            (tabla, pk_col, str(pk_val), campo, val_ant, val_nuevo_s, usuario_nombre)
                        )

                # ── Cascade pieza_virtual: actualiza pieza en malla linked ────
                if pieza_virtual is not _SENTINEL and "vitrojet" in tabla:
                    try:
                        cur.execute(
                            "SELECT codigo_malla FROM mallas.vitrojet WHERE vitro=?",
                            (pk_val,))
                        r_cm2 = cur.fetchone()
                        if r_cm2 and r_cm2[0]:
                            cod_m2 = str(r_cm2[0]).strip()
                            cur.execute(
                                "UPDATE mallas.grandes SET pieza=? WHERE CAST(codigo AS NVARCHAR)=?",
                                (pieza_virtual, cod_m2))
                            if cur.rowcount == 0:
                                cur.execute(
                                    "UPDATE mallas.pequenas SET pieza=? WHERE CAST(codigo AS NVARCHAR)=?",
                                    (pieza_virtual, cod_m2))
                    except Exception:
                        pass

                # ── Cascade ruta: sincroniza vitro ↔ malla vinculada ─────
                try:
                    if "vitrojet" in tabla and "ruta" in nuevos and nuevos["ruta"]:
                        cur.execute(
                            "SELECT codigo_malla FROM mallas.vitrojet WHERE vitro=?",
                            (pk_val,))
                        r_cm = cur.fetchone()
                        if r_cm and r_cm[0]:
                            cod_m = str(r_cm[0]).strip()
                            cur.execute(
                                "UPDATE mallas.grandes SET ruta_dwg=? "
                                "WHERE CAST(codigo AS NVARCHAR)=?",
                                (nuevos["ruta"], cod_m))
                            if cur.rowcount == 0:
                                cur.execute(
                                    "UPDATE mallas.pequenas SET ruta_dwg=? "
                                    "WHERE CAST(codigo AS NVARCHAR)=?",
                                    (nuevos["ruta"], cod_m))
                    elif ("grandes" in tabla or "pequenas" in tabla) \
                            and "ruta_dwg" in nuevos and nuevos["ruta_dwg"]:
                        cur.execute(
                            "SELECT vitro FROM mallas.vitrojet WHERE codigo_malla=?",
                            (str(pk_val),))
                        r_vt = cur.fetchone()
                        if r_vt and r_vt[0]:
                            cur.execute(
                                "UPDATE mallas.vitrojet SET ruta=? WHERE vitro=?",
                                (nuevos["ruta_dwg"], r_vt[0]))
                except Exception:
                    pass  # cascade es best-effort, no falla el guardado principal

                cn.commit()
                cn.close()
                msg_lbl.configure(text="✔ Guardado correctamente", fg="#22c55e")
                win.after(1200, win.destroy)
                self.after(200, self._do_search)
                if self._on_data_changed:
                    self.after(400, self._on_data_changed)
            except Exception as ex:
                msg_lbl.configure(text=f"✘ Error: {ex}", fg="#ef4444")

        btn_ok = _tk.Button(btn_f, text="  Guardar cambios  ",
                            font=("Segoe UI", 10, "bold"),
                            bg=PAL["accent"], fg="white", relief="flat",
                            activebackground="#2563eb", cursor="hand2",
                            command=_guardar)
        btn_ok.pack(side="left", ipadx=6, ipady=6, padx=(0,10))

        btn_cancel = _tk.Button(btn_f, text="  Cancelar  ",
                                font=("Segoe UI", 10),
                                bg="#2d2d44", fg="#94a3b8", relief="flat",
                                activebackground="#3d3d5c", cursor="hand2",
                                command=win.destroy)
        btn_cancel.pack(side="left", ipadx=6, ipady=6, padx=(0,20))

        def _anular():
            import ctypes as _ct
            pk_label = pk_col.upper()
            confirmar_txt = (
                f"¿Seguro que quieres ANULAR este registro?\n\n"
                f"  {pk_label}: {pk_val}\n\n"
                f"Esto borrará todos los datos asociados (vehículo, ruta, responsable, etc.)\n"
                f"y pondrá el número como CANCELADO para que otro lo pueda tomar.\n\n"
                f"Si tiene vitro+malla vinculados, ambos quedarán cancelados."
            )
            # MB_YESNO | MB_ICONWARNING | MB_TOPMOST — siempre al frente
            resp = _ct.windll.user32.MessageBoxW(
                0, confirmar_txt, "Confirmar anulación", 0x04 | 0x30 | 0x40000)
            if resp != 6:  # 6 = IDYES
                return
            try:
                res = _anular_asignacion(tab, pk_val)
                partes = [f"{k}: {v}" for k, v in res.items() if v > 0]
                msg_lbl.configure(
                    text=f"✔ Anulado — {', '.join(partes) if partes else 'sin cambios'}",
                    fg="#f59e0b")
                win.after(1500, win.destroy)
                self.after(200, self._do_search)
                if self._on_data_changed:
                    self.after(400, self._on_data_changed)
            except Exception as ex:
                msg_lbl.configure(text=f"✘ Error: {ex}", fg="#ef4444")

        if tab in ("vinilos", "pasta_plata"):
            return  # sin botón anular para estas tablas

        btn_anular = _tk.Button(btn_f, text="⚠  Quedó mal — Anular",
                                font=("Segoe UI", 10, "bold"),
                                bg="#7f1d1d", fg="#fca5a5", relief="flat",
                                activebackground="#991b1b", cursor="hand2",
                                command=_anular)
        btn_anular.pack(side="right", ipadx=6, ipady=6)

        win.update_idletasks()
        # Centrar y limitar altura a pantalla
        try:
            sh = win.winfo_screenheight()
            w  = win.winfo_reqwidth()
            h  = min(win.winfo_reqheight(), sh - 80)
            rx = root_tk.winfo_rootx() + root_tk.winfo_width()  // 2 - w // 2
            ry = root_tk.winfo_rooty() + root_tk.winfo_height() // 2 - h // 2
            win.geometry(f"{w}x{h}+{rx}+{ry}")
        except Exception:
            pass

    def _insertar_manual(self):
        import tkinter as _tk
        from tkinter import messagebox as _mb

        try:
            root_tk = _tk._default_root
        except Exception:
            root_tk = None

        # ── Ventana selector de tabla ─────────────────────────────────────────
        sel_win = _tk.Toplevel(root_tk)
        sel_win.title("Insertar manual")
        sel_win.configure(bg="#1e1e2e")
        sel_win.resizable(False, False)
        sel_win.attributes("-topmost", True)
        if root_tk:
            sel_win.grab_set()

        _tk.Frame(sel_win, bg=PAL["green2"], height=4).pack(fill="x")
        _tk.Frame(sel_win, bg="#1e1e2e", height=10).pack()
        _tk.Label(sel_win, text="📝  Insertar registro manual",
                  font=("Segoe UI", 14, "bold"), fg="#e2e8f0", bg="#1e1e2e"
                  ).pack(padx=24, anchor="w")
        _tk.Label(sel_win, text="Elige la tabla donde quieres insertar:",
                  font=("Segoe UI", 9), fg="#94a3b8", bg="#1e1e2e"
                  ).pack(padx=24, pady=(4,14), anchor="w")

        TABLAS_INS = [
            ("🪙  Pasta de Plata",  "pasta_plata"),
            ("🎨  Vinilos",         "vinilos"),
        ]

        def _abrir(tabla):
            sel_win.destroy()
            self._form_insertar(tabla, root_tk)

        for lbl, key in TABLAS_INS:
            _tk.Button(sel_win, text=lbl, font=("Segoe UI", 11, "bold"),
                       bg="#2d2d44", fg="#e2e8f0", relief="flat",
                       activebackground="#3d3d5c", cursor="hand2",
                       width=28, pady=10,
                       command=lambda k=key: _abrir(k)
                       ).pack(fill="x", padx=20, pady=4)

        _tk.Frame(sel_win, bg="#1e1e2e", height=14).pack()
        sel_win.update_idletasks()
        try:
            rx = root_tk.winfo_rootx() + root_tk.winfo_width()  // 2 - sel_win.winfo_width()  // 2
            ry = root_tk.winfo_rooty() + root_tk.winfo_height() // 2 - sel_win.winfo_height() // 2
            sel_win.geometry(f"+{rx}+{ry}")
        except Exception:
            pass

    # Campos de cada tabla: (label, campo_bd, sugerido_o_None, solo_lectura)
    _FORM_CAMPOS = {
        "pasta_plata": [
            ("Consecutivo",  "consecutivo",  "__AUTO__",  True),
            ("RED/ANT",          "tipo",         None,        False),
            ("Nombre vehículo",  "vehiculo",     None,        False),
            ("Cód. Vehículo","cod_vehiculo", None,        False),
            ("Versión",      "version",      None,        False),
            ("Pieza",        "pieza",        None,        False),
            ("Ruta archivo", "ruta_archivo", None,        False),
            ("Caso",         "caso",         None,        False),
        ],
        "vinilos": [
            ("Herramental",  "herramental",  "__AUTO__",  True),
            ("Vehículo",     "vehiculo",     None,        False),
            ("Cód. Vehículo","cod_vehiculo", None,        False),
            ("Versión",      "version",      None,        False),
            ("Pieza",        "pieza",        None,        False),
            ("Tipo",         "tipo",         None,        False),
        ],
    }
    _FORM_TITULO = {
        "pasta_plata": ("🪙", "Pasta de Plata", "mallas.pasta_plata"),
        "vinilos":     ("🎨", "Vinilos",         "mallas.vinilos"),
    }

    def _form_insertar(self, tabla, root_tk):
        import tkinter as _tk
        from tkinter import messagebox as _mb

        icon, titulo, tabla_sql = self._FORM_TITULO[tabla]
        campos = self._FORM_CAMPOS[tabla]

        def _next_consecutivo():
            """Calcula el siguiente código según el formato de la tabla."""
            try:
                cn2 = db_connect()
                cur2 = cn2.cursor()
                if tabla == "pasta_plata":
                    cur2.execute("SELECT ISNULL(MAX(TRY_CAST(SUBSTRING(consecutivo,3,50) AS INT)),0)+1 FROM mallas.pasta_plata")
                    n = cur2.fetchone()[0]
                    resultado = f"S-{n:05d}"
                elif tabla == "vinilos":
                    cur2.execute("SELECT ISNULL(MAX(TRY_CAST(SUBSTRING(herramental,4,50) AS INT)),0)+1 FROM mallas.vinilos")
                    n = cur2.fetchone()[0]
                    resultado = f"VC-{n:04d}"
                else:
                    resultado = ""
                cn2.close()
                return resultado
            except Exception:
                return "?"

        auto_consecutivo = _next_consecutivo() if tabla in ("pasta_plata", "vinilos") else None

        win = _tk.Toplevel(root_tk)
        win.title(f"Insertar — {titulo}")
        win.configure(bg="#1e1e2e")
        win.resizable(False, False)
        win.attributes("-topmost", True)
        if root_tk:
            win.grab_set()

        _tk.Frame(win, bg=PAL["green2"], height=4).pack(fill="x")
        _tk.Frame(win, bg="#1e1e2e", height=10).pack()
        _tk.Label(win, text=f"{icon}  Insertar — {titulo}",
                  font=("Segoe UI", 14, "bold"), fg="#e2e8f0", bg="#1e1e2e"
                  ).pack(anchor="w", padx=24, pady=(0,2))
        _tk.Label(win, text="Completa los datos del nuevo registro:",
                  font=("Segoe UI", 9), fg="#94a3b8", bg="#1e1e2e"
                  ).pack(anchor="w", padx=24, pady=(0,12))
        _tk.Frame(win, bg="#2d2d44", height=1).pack(fill="x", padx=18, pady=(0,12))

        entries_ins = {}
        for lbl_txt, campo, sugerido, readonly in campos:
            row_f = _tk.Frame(win, bg="#1e1e2e")
            row_f.pack(fill="x", padx=24, pady=4)
            _tk.Label(row_f, text=lbl_txt, width=17, anchor="e",
                      font=("Segoe UI", 9), fg="#94a3b8", bg="#1e1e2e"
                      ).pack(side="left", padx=(0,10))

            val_inicial = auto_consecutivo if sugerido == "__AUTO__" else (sugerido or "")
            _DROP_INS = {
                "caso":  (["Caso 1", "Caso 2"], "Caso 1"),
                "tipo":  (["RED", "ANT"],        "RED") if tabla == "pasta_plata" else None,
            }
            _drop_cfg = _DROP_INS.get(campo)
            if _drop_cfg:
                opciones, default = _drop_cfg
                var = _tk.StringVar(value=default)
                opt = _tk.OptionMenu(row_f, var, *opciones)
                opt.config(bg="#2d2d44", fg="#e2e8f0", activebackground="#3d4f7c",
                           activeforeground="#e2e8f0", relief="flat", bd=0,
                           font=("Segoe UI", 10), width=10, highlightthickness=0)
                opt["menu"].config(bg="#2d2d44", fg="#e2e8f0",
                                   activebackground="#3d4f7c", font=("Segoe UI", 10))
                opt.pack(side="left", ipady=3)
                entries_ins[campo] = var
            else:
                bg_ent = "#1a2a1a" if readonly else "#2d2d44"
                ent = _tk.Entry(row_f, width=40,
                                font=("Segoe UI", 10),
                                bg=bg_ent, fg="#e2e8f0" if not readonly else "#6ee7b7",
                                insertbackground="#e2e8f0",
                                relief="flat", bd=6,
                                state="normal")
                ent.pack(side="left", ipady=4)
                ent.insert(0, str(val_inicial))
                if readonly:
                    ent.configure(state="readonly")
                    _tk.Label(row_f, text="auto", font=("Segoe UI", 7, "bold"),
                              fg=PAL["green2"], bg="#1e1e2e", padx=4).pack(side="left", padx=4)
                entries_ins[campo] = ent

        _tk.Frame(win, bg="#1e1e2e", height=6).pack()
        _tk.Frame(win, bg="#2d2d44", height=1).pack(fill="x", padx=18, pady=(0,12))

        btn_f   = _tk.Frame(win, bg="#1e1e2e")
        btn_f.pack(fill="x", padx=24, pady=(0,18))
        msg_lbl = _tk.Label(btn_f, text="", font=("Segoe UI", 9),
                            fg="#22c55e", bg="#1e1e2e")
        msg_lbl.pack(anchor="w", pady=(0,8))

        def _insertar():
            # Recoger valores del formulario (saltando el PK auto)
            pk_auto = {"pasta_plata": "consecutivo", "vinilos": "herramental"}.get(tabla)
            cols_datos = [c for _, c, _, _ in campos if c != pk_auto]
            vals_datos = [entries_ins[c].get().strip() or None for c in cols_datos]

            # Validar que haya al menos un campo con dato
            if not any(vals_datos):
                msg_lbl.configure(text="✘ Completa al menos un campo", fg="#ef4444")
                return

            try:
                cn = db_connect()
                if pk_auto:
                    # INSERT...SELECT atómico: calcula MAX e inserta en una sola
                    # operación — Azure SQL no puede entregar el mismo número a
                    # dos conexiones concurrentes porque el SELECT con UPDLOCK
                    # bloquea la lectura hasta que el INSERT confirma.
                    if tabla == "pasta_plata":
                        expr_pk = ("'S-' + RIGHT('00000' + CAST("
                                   "ISNULL(MAX(TRY_CAST(SUBSTRING(consecutivo,3,50) AS INT)),0)+1"
                                   " AS VARCHAR(10)), 5)")
                    else:  # vinilos
                        expr_pk = ("'VC-' + RIGHT('0000' + CAST("
                                   "ISNULL(MAX(TRY_CAST(SUBSTRING(herramental,4,50) AS INT)),0)+1"
                                   " AS VARCHAR(10)), 4)")

                    all_cols = [pk_auto] + cols_datos
                    placeholders_datos = ",".join(["?"] * len(cols_datos))
                    sql = (
                        f"INSERT INTO {tabla_sql} ({','.join(all_cols)}) "
                        f"SELECT {expr_pk}, {placeholders_datos} "
                        f"FROM {tabla_sql} WITH (UPDLOCK, HOLDLOCK)"
                    )
                    cur = cn.cursor()
                    cur.execute(sql, vals_datos)
                    # Obtener el código que quedó insertado para mostrarlo
                    pk_col_name = pk_auto
                    cur.execute(f"SELECT TOP 1 {pk_col_name} FROM {tabla_sql} "
                                f"ORDER BY {pk_col_name} DESC")
                    pk_insertado = cur.fetchone()[0]
                else:
                    all_cols = [c for _, c, _, _ in campos]
                    vals_all  = [entries_ins[c].get().strip() or None for c in all_cols]
                    placeholders = ",".join(["?"] * len(all_cols))
                    sql = f"INSERT INTO {tabla_sql} ({','.join(all_cols)}) VALUES ({placeholders})"
                    cn.execute(sql, vals_all)
                    pk_insertado = None

                cn.commit()
                cn.close()

                # Actualizar el campo readonly con el código real insertado
                if pk_auto and pk_insertado:
                    entries_ins[pk_auto].configure(state="normal")
                    entries_ins[pk_auto].delete(0, _tk.END)
                    entries_ins[pk_auto].insert(0, str(pk_insertado))
                    entries_ins[pk_auto].configure(state="readonly")

                msg_lbl.configure(
                    text=f"✔ Insertado: {pk_insertado or 'OK'}", fg="#22c55e")
                win.after(1500, win.destroy)
                self.after(200, self._do_search)
                if self._on_data_changed:
                    self.after(400, self._on_data_changed)
            except Exception as ex:
                msg_lbl.configure(text=f"✘ Error: {str(ex)[:90]}", fg="#ef4444")

        _tk.Button(btn_f, text="  Insertar registro  ",
                   font=("Segoe UI", 10, "bold"),
                   bg=PAL["green2"], fg="white", relief="flat",
                   activebackground="#1e5c36", cursor="hand2",
                   command=_insertar
                   ).pack(side="left", ipadx=6, ipady=6, padx=(0,10))

        _tk.Button(btn_f, text="  Cancelar  ",
                   font=("Segoe UI", 10),
                   bg="#2d2d44", fg="#94a3b8", relief="flat",
                   activebackground="#3d3d5c", cursor="hand2",
                   command=win.destroy
                   ).pack(side="left", ipadx=6, ipady=6)

        win.update_idletasks()
        try:
            rx = root_tk.winfo_rootx() + root_tk.winfo_width()  // 2 - win.winfo_width()  // 2
            ry = root_tk.winfo_rooty() + root_tk.winfo_height() // 2 - win.winfo_height() // 2
            win.geometry(f"+{rx}+{ry}")
        except Exception:
            pass

    def refresh(self):
        """Refresca la tabla (llamado desde otras pestañas tras cambios)."""
        self._do_search()

    def _separar(self):
        try:
            import tkinter as _tk
            prop = _dialogo_separar(parent_win=_tk._default_root)
            if prop:
                self._do_search()
                if self._on_data_changed:
                    self.after(200, self._on_data_changed)
        except Exception as e:
            _msgbox_topmost("error", "Error", str(e))


# ══════════════════════════════════════════════════════════════════════════════
#  PESTAÑA — SCANNER DE ÓRDENES
# ══════════════════════════════════════════════════════════════════════════════
def _conectar_comercial():
    if _pymssql is None:
        raise RuntimeError("pymssql no disponible")
    try:
        conn = _pymssql.connect(
            server="192.168.2.23",
            user="Consulta",
            password="@GPgl4$$2021",
            database="Comercial",
            timeout=10,
            login_timeout=10,
            tds_version="7.3",
        )
        return _ConnWrap(conn)
    except Exception as e:
        raise RuntimeError(f"No se pudo conectar a Comercial (192.168.2.23)\n{e}")

def _conectar_sap():
    if _pymssql is None:
        raise RuntimeError("pymssql no disponible")
    try:
        conn = _pymssql.connect(
            server="agpcolsap.database.windows.net",
            port=1433,
            user="Viewer",
            password="AgpconsCol2023",
            database="DB_COL_SAP",
            timeout=15,
            login_timeout=15,
            charset="UTF-8",
            tds_version="7.3",
        )
        return _ConnWrap(conn)
    except Exception as e:
        raise RuntimeError(f"No se pudo conectar a SAP Azure\n{e}")


def _conectar_calendario():
    if _pymssql is None:
        raise RuntimeError("pymssql no disponible")
    try:
        conn = _pymssql.connect(
            server="agpcolcalendario.database.windows.net",
            port=1433,
            user="Consulta",
            password="@GPgl4$$2021",
            database="CalendarioAGP",
            timeout=15,
            login_timeout=15,
            charset="UTF-8",
            tds_version="7.3",
        )
        return _ConnWrap(conn)
    except Exception as e:
        raise RuntimeError(f"No se pudo conectar a Calendario\n{e}")


def _conectar_sap_prod():
    if _pymssql is None:
        raise RuntimeError("pymssql no disponible")
    try:
        conn = _pymssql.connect(
            server="agpcol.database.windows.net",
            port=1433,
            user="Consulta",
            password="@GPgl4$$2021",
            database="agpc-productivity",
            timeout=15,
            login_timeout=15,
            charset="UTF-8",
            tds_version="7.3",
        )
        return _ConnWrap(conn)
    except Exception as e:
        raise RuntimeError(f"No se pudo conectar a Producción\n{e}")


_PIEZAS_MAP = {
    "000": "Parabrisas",
    "001": "Lat. Del. Izq.", "002": "Lat. Del. Der.",
    "003": "Lat. Tra. Izq.", "004": "Lat. Tra. Der.",
    "005": "Ventilete Tra. Izq.", "006": "Ventilete Tra. Der.",
    "007": "Cabina Tra. Izq.", "008": "Cabina Tra. Der.",
    "009": "Posterior", "010": "Techo Solar Del.",
    "011": "Lat. Ext. Izq.", "012": "Lat. Ext. Der.",
    "013": "Post. Izq.", "014": "Post. Der.",
    "015": "Claraboya Izq.", "016": "Claraboya Der.",
    "017": "Mirilla", "018": "Probeta",
    "019": "Ventilete Del. Izq.", "020": "Ventilete Del. Der.",
    "021": "Cabina Del. Izq.", "022": "Cabina Del. Der.",
    "023": "Cabina Sup. Izq.", "024": "Cabina Sup. Der.",
    "025": "Techo Solar B", "026": "Parabrisas Der.",
    "027": "Parabrisas Izq.", "030": "Partición",
    "031": "Arquitectura", "040": "Pummel",
    "085": "Post. Secundario", "087": "Techo Solar Céntrico",
    "088": "Techo Solar D", "090": "Techo Solar Panorámico",
}


class TabScanner(ctk.CTkFrame):

    def __init__(self, parent, **kw):
        super().__init__(parent, fg_color=PAL["bg"], **kw)
        self._build()

    # ── helpers de copiar ─────────────────────────────────────────────────────
    def _copiar(self, texto):
        if texto:
            self.clipboard_clear(); self.clipboard_append(texto); self.update()

    def _copiar_ruta_vitro(self):  self._copiar(self._vitro_ruta_val)
    def _copiar_alerta(self):      self._copiar(self._alerta_txt)

    def _build(self):
        import tkinter.ttk as _ttk
        self.columnconfigure(0, weight=1)
        self.rowconfigure(1, weight=1)
        self.configure(fg_color="#0d1117")

        # ══ Barra de búsqueda ════════════════════════════════════════════════
        top = ctk.CTkFrame(self, fg_color="#161b22", corner_radius=0)
        top.grid(row=0, column=0, sticky="ew")
        top.columnconfigure(0, weight=1)
        sr = ctk.CTkFrame(top, fg_color="transparent")
        sr.pack(fill="x", padx=28, pady=20)
        sr.columnconfigure(0, weight=1)
        self._entry = ctk.CTkEntry(
            sr, placeholder_text="  🔍  Escanea o escribe el número de orden...",
            height=62, font=FONT(18),
            fg_color="#0d1117", border_color="#21262d",
            border_width=2, text_color="#e6edf3", corner_radius=14,
        )
        self._entry.grid(row=0, column=0, sticky="ew", padx=(0, 12))
        self._entry.bind("<Return>",   lambda _: self._buscar())
        self._entry.bind("<KP_Enter>", lambda _: self._buscar())
        self._entry.focus_set()
        ctk.CTkButton(
            sr, text="  BUSCAR  ", width=130, height=62,
            font=FONT(14, "bold"), corner_radius=14,
            fg_color="#1f6feb", hover_color="#388bfd",
            text_color="white",
            command=self._buscar,
        ).grid(row=0, column=1)
        self._prog = ctk.CTkProgressBar(top, mode="indeterminate",
                                         height=2, progress_color="#58a6ff",
                                         fg_color="#0d1117")
        self._prog.pack(fill="x")
        self._prog.set(0)

        # ══ Zona resultado ═══════════════════════════════════════════════════
        self._zona = ctk.CTkFrame(self, fg_color="transparent")
        self._zona.grid(row=1, column=0, sticky="nsew", padx=12, pady=8)
        self._zona.columnconfigure(0, weight=3)
        self._zona.columnconfigure(1, weight=2)
        self._zona.rowconfigure(0, weight=1)

        # ── Columna izquierda ─────────────────────────────────────────────
        col_izq = ctk.CTkFrame(self._zona, fg_color="transparent")
        col_izq.grid(row=0, column=0, sticky="nsew", padx=(0, 10))
        col_izq.columnconfigure(0, weight=1)

        # · Fila ORDEN | ZFER  (header bar unificada)
        hbar = ctk.CTkFrame(col_izq, fg_color="#161b22", corner_radius=16,
                             border_width=1, border_color="#21262d")
        hbar.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        hbar.columnconfigure((0, 1), weight=1)

        def _id_block(parent, etiqueta, color_bg, color_txt, col_idx):
            blk = ctk.CTkFrame(parent, fg_color=color_bg, corner_radius=10)
            blk.grid(row=0, column=col_idx, sticky="ew",
                     padx=(10, 5) if col_idx == 0 else (5, 10), pady=8)
            blk.columnconfigure(0, weight=1)
            ctk.CTkLabel(blk, text=etiqueta, font=FONT(9, "bold"),
                         text_color=color_txt).pack(anchor="w", padx=12, pady=(8, 0))
            lbl = ctk.CTkLabel(blk, text="—", font=FONT(32, "bold"),
                               text_color="#e6edf3")
            lbl.pack(anchor="w", padx=14, pady=(0, 12))
            return lbl

        self._c_orden = _id_block(hbar, "ORDEN", "#0d1f33", "#58a6ff", 0)
        self._c_zfer  = _id_block(hbar, "ZFER",  "#0d2218", "#3fb950", 1)

        # · Tarjeta Vehículo (borde izquierdo azul)
        wrap_veh = ctk.CTkFrame(col_izq, fg_color="#58a6ff", corner_radius=14)
        wrap_veh.grid(row=1, column=0, sticky="ew", pady=(0, 8))
        wrap_veh.columnconfigure(0, weight=1)
        veh_card = ctk.CTkFrame(wrap_veh, fg_color="#161b22", corner_radius=12)
        veh_card.pack(fill="both", expand=True, padx=(4, 0), pady=0)
        veh_card.columnconfigure(1, weight=1)

        # Fila 0: logo + nombre vehículo
        veh_right = ctk.CTkFrame(veh_card, fg_color="transparent")
        veh_right.grid(row=0, column=0, sticky="ew", padx=(14, 12), pady=(12, 0))
        veh_right.columnconfigure(0, weight=1)
        self._lbl_vehiculo = ctk.CTkLabel(
            veh_right, text="—", font=FONT(22, "bold"),
            text_color="#e6edf3", anchor="w",
        )
        self._lbl_vehiculo.grid(row=0, column=0, sticky="w")

        self._lbl_logo = ctk.CTkLabel(
            veh_card, text="", font=FONT(16, "bold"),
            text_color="#0d1117", fg_color="#1f6feb",
            corner_radius=10, width=90, height=38,
        )
        self._lbl_logo.grid(row=0, column=1, padx=(0, 14), pady=(12, 0), sticky="e")
        # _lbl_version existe solo para el reset (texto vacío, no visible)
        self._lbl_version = ctk.CTkLabel(veh_right, text="")

        # Fila de celdas — 5 columnas uniformes: LOTE | VERSIÓN | CÓD.VEH. | TIPO PIEZA | TRAZABILIDAD
        info4 = ctk.CTkFrame(veh_card, fg_color="transparent")
        info4.grid(row=1, column=0, columnspan=2, sticky="ew", padx=14, pady=(10, 14))
        info4.columnconfigure((0, 1, 2, 3, 4), weight=1)

        def _info_cell(parent, titulo, col_n, color_val):
            px = (0 if col_n == 0 else 10, 0)
            ctk.CTkLabel(parent, text=titulo, font=FONT(9, "bold"),
                         text_color="#484f58").grid(row=0, column=col_n, sticky="w", padx=px)
            lbl = ctk.CTkLabel(parent, text="—", font=FONT(18, "bold"),
                               text_color=color_val, anchor="w")
            lbl.grid(row=1, column=col_n, sticky="w", padx=px)
            return lbl

        self._lbl_lote         = _info_cell(info4, "LOTE",          0, "#e3b341")
        self._lbl_version2     = _info_cell(info4, "VERSIÓN",        1, "#58a6ff")
        self._lbl_sap_codveh   = _info_cell(info4, "CÓD. VEH.",      2, "#79c0ff")
        self._lbl_sap_tipo     = _info_cell(info4, "TIPO PIEZA",     3, "#d2a8ff")
        self._lbl_trazabilidad = _info_cell(info4, "TRAZABILIDAD",   4, "#ffa657")

        # · Box VITRO (card con borde izquierdo verde)
        wrap_v = ctk.CTkFrame(col_izq, fg_color="#3fb950", corner_radius=16)
        wrap_v.grid(row=2, column=0, sticky="ew")
        wrap_v.columnconfigure(0, weight=1)
        wrap_v.rowconfigure(0, weight=1)
        box_v = ctk.CTkFrame(wrap_v, fg_color="#161b22", corner_radius=14)
        box_v.pack(fill="both", expand=True, padx=(4, 0), pady=0)
        box_v.columnconfigure(0, weight=1)

        vit_hdr = ctk.CTkFrame(box_v, fg_color="transparent")
        vit_hdr.pack(fill="x", padx=16, pady=(14, 4))
        ctk.CTkFrame(vit_hdr, fg_color="#3fb950", width=6, height=6,
                     corner_radius=3).pack(side="left", padx=(0, 7))
        ctk.CTkLabel(vit_hdr, text="VITRO", font=FONT(13, "bold"),
                     text_color="#3fb950").pack(side="left")

        self._lbl_vitro2 = ctk.CTkLabel(
            box_v, text="—", font=FONT(36, "bold"),
            text_color="#e6edf3", wraplength=620, justify="left",
        )
        self._lbl_vitro2.pack(anchor="w", padx=16, pady=(4, 16))

        self._vitro_ruta_val = ""
        self._lbl_vitro_ruta = ctk.CTkLabel(box_v, text="")  # oculto, solo datos

        # ── Columna derecha: MALLAS ───────────────────────────────────────
        wrap_m = ctk.CTkFrame(self._zona, fg_color="#8957e5", corner_radius=16)
        wrap_m.grid(row=0, column=1, sticky="nsew", padx=(10, 0))
        wrap_m.columnconfigure(0, weight=1)
        wrap_m.rowconfigure(0, weight=1)
        col_der = ctk.CTkFrame(wrap_m, fg_color="#161b22", corner_radius=14)
        col_der.pack(fill="both", expand=True, padx=(4, 0), pady=0)
        col_der.columnconfigure(0, weight=1)
        col_der.rowconfigure(1, weight=1)

        mal_hdr = ctk.CTkFrame(col_der, fg_color="transparent")
        mal_hdr.pack(fill="x", padx=18, pady=(16, 4))
        ctk.CTkFrame(mal_hdr, fg_color="#8957e5", width=6, height=6,
                     corner_radius=3).pack(side="left", padx=(0, 8))
        ctk.CTkLabel(mal_hdr, text="MALLAS", font=FONT(13, "bold"),
                     text_color="#d2a8ff").pack(side="left")
        self._lbl_mallas_count = ctk.CTkLabel(
            mal_hdr, text="", font=FONT(9),
            text_color="#484f58", fg_color="#21262d",
            corner_radius=8, width=24, height=20)
        self._lbl_mallas_count.pack(side="left", padx=(8, 0))

        ctk.CTkFrame(col_der, fg_color="#21262d", height=1
                     ).pack(fill="x", padx=18, pady=(0, 8))

        self._mallas_box = ctk.CTkScrollableFrame(
            col_der, fg_color="transparent", corner_radius=0)
        self._mallas_box.pack(fill="both", expand=True, padx=10, pady=(0, 10))
        self._mallas_box.columnconfigure(0, weight=1)
        self._mallas_labels = []

        # ── Alerta rutas (oculta) ─────────────────────────────────────────
        self._alerta_frame = ctk.CTkFrame(
            self, fg_color="#2d1b00", corner_radius=0, border_width=0)
        self._alerta_frame.grid(row=2, column=0, sticky="ew")
        self._alerta_frame.columnconfigure(0, weight=1)
        # Barra de acento naranja arriba
        ctk.CTkFrame(self._alerta_frame, fg_color="#f78166", height=2,
                     corner_radius=0).grid(row=0, column=0, columnspan=2,
                                           sticky="ew")
        self._alerta_lbl = ctk.CTkLabel(
            self._alerta_frame, text="", font=FONT(12, "bold"),
            text_color="#ffa657", justify="left", wraplength=900, anchor="w",
        )
        self._alerta_lbl.grid(row=1, column=0, sticky="w", padx=20, pady=(8, 10))
        ctk.CTkButton(
            self._alerta_frame, text="📋 COPIAR ALERTA", width=150, height=32,
            font=FONT(11, "bold"), corner_radius=8,
            fg_color="#5a2e0e", hover_color="#7a3e10",
            text_color="#ffa657",
            command=self._copiar_alerta,
        ).grid(row=1, column=1, padx=(8, 20), pady=(8, 10))
        self._alerta_frame.grid_remove()
        self._alerta_txt = ""

        self._zona.grid_remove()

    # ── Lógica ───────────────────────────────────────────────────────────────
    def _buscar(self):
        orden = self._entry.get().strip()
        if not orden:
            return
        self._prog.start()
        self._entry.configure(state="disabled")
        threading.Thread(target=self._t_buscar, args=(orden,), daemon=True).start()

    def _t_buscar(self, orden):
        try:
            cn = _conectar_comercial()
            cur = cn.cursor()
            cur.execute(
                "SELECT TOP 1 ZFER FROM [dbo].[BI_TAB_FIN_TURNO_SAP] "
                "WHERE ORDEN = ? AND ZFER IS NOT NULL", (orden,))
            row = cur.fetchone()
            cn.close()
            if not row:
                self.after(0, self._mostrar_error, f"Orden  {orden}  no encontrada")
                return
            zfer = str(row[0]).strip()

            cn2 = _conectar_sap()
            cur2 = cn2.cursor()
            cur2.execute(
                "SELECT DENOMINACION_COMPONENTE, TEXTO1, TEXTO2 "
                "FROM ODATA_ZFER_BOM WHERE MATERIAL = ?", (zfer,))
            filas = cur2.fetchall()
            cn2.close()

            _EXCLUIR = ("M", "G")
            _EXCLUIR_DESC = ("MOLDE LLENO", "PLANTILLA", "ANTENA")
            mallas, t1s, t2s = [], [], []
            for comp, t1, t2 in filas:
                if comp and str(comp).strip():
                    c = str(comp).strip()
                    cu = c.upper()
                    if (not cu.startswith(_EXCLUIR) and
                            not any(x in cu for x in _EXCLUIR_DESC)):
                        mallas.append(c)
                if t1 and str(t1).strip() and str(t1).strip() not in t1s:
                    t1s.append(str(t1).strip())
                if t2 and str(t2).strip() and str(t2).strip() not in t2s:
                    t2s.append(str(t2).strip())

            # ── Consultas paralelas: Calendario + SAP pieza + Rutas ──────
            vitro_cod = t2s[0].lstrip(";").strip().split()[0] if t2s else None
            cal        = {}
            sap_info   = {}
            vitro_ruta = None
            malla_rutas = {m: None for m in mallas}
            _ruta_err  = None

            def _q_calendario():
                try:
                    cn_c = _conectar_calendario()
                    cur_c = cn_c.cursor()
                    cur_c.execute(
                        "SELECT TOP 1 PedidoGenesis, "
                        "RIGHT('0000000000' + CAST(Lote AS VARCHAR), 10), "
                        "Logo, ZFER, CodVehiculo, Vehiculo, Version, Cliente "
                        "FROM [dbo].[TCAL_CALENDARIO_COLOMBIA_DIRECT] "
                        "WHERE Orden = ?", (orden,))
                    r = cur_c.fetchone()
                    cn_c.close()
                    if r:
                        cal.update({
                            "lote": str(r[1] or "").strip(),
                            "logo": str(r[2] or "").strip(),
                            "cod_veh": str(r[4] or ""),
                            "vehiculo": str(r[5] or "").strip(),
                            "version": str(r[6] or "").strip(),
                            "cliente": str(r[7] or "").strip(),
                        })
                except Exception:
                    pass

            def _q_pieza():
                # ODATA_ZFER_CLASS_001 está en la BD SAP (misma que ODATA_ZFER_BOM)
                try:
                    cn_s = _conectar_sap()
                    cur_s = cn_s.cursor()
                    zfer_pad = zfer.zfill(18)
                    cur_s.execute(
                        "SELECT ATNAM, ATWRT FROM ODATA_ZFER_CLASS_001 "
                        "WHERE MATERIAL IN (?, ?) AND ATNAM IN ('Z_PIECE_TYPE','Z_TRAZABILITY')",
                        (zfer, zfer_pad))
                    for atnam, atwrt in cur_s.fetchall():
                        sap_info[atnam] = str(atwrt or "").strip()
                    cn_s.close()
                except Exception:
                    pass

            def _q_rutas():
                nonlocal vitro_ruta, _ruta_err
                try:
                    cn3 = db_connect()
                    cur3 = cn3.cursor()
                    if vitro_cod:
                        try:
                            cur3.execute(
                                "SELECT TOP 1 ruta FROM mallas.vitrojet WHERE vitro = ?",
                                (vitro_cod,))
                            r = cur3.fetchone()
                            vitro_ruta = str(r[0]).strip() if r and r[0] else None
                        except Exception:
                            pass
                    for m in mallas:
                        m_cod = m.split()[0]
                        try:
                            cur3.execute(
                                "SELECT TOP 1 ruta_dwg FROM mallas.grandes WHERE codigo = ?",
                                (m_cod,))
                            r = cur3.fetchone()
                            if not r or not r[0]:
                                cur3.execute(
                                    "SELECT TOP 1 ruta_dwg FROM mallas.pequenas WHERE codigo = ?",
                                    (m_cod,))
                                r = cur3.fetchone()
                            malla_rutas[m] = str(r[0]).strip() if r and r[0] else None
                        except Exception:
                            malla_rutas[m] = None
                    cn3.close()
                except Exception as _e:
                    _ruta_err = str(_e)

            # Lanzar las 3 consultas en paralelo
            _threads = [
                threading.Thread(target=_q_calendario, daemon=True),
                threading.Thread(target=_q_pieza,      daemon=True),
                threading.Thread(target=_q_rutas,      daemon=True),
            ]
            for t in _threads: t.start()
            for t in _threads: t.join(timeout=12)

            self.after(0, self._mostrar_resultado, orden, zfer,
                       " / ".join(t1s) or "—",
                       " / ".join(v.lstrip(";") for v in t2s), mallas,
                       vitro_ruta, malla_rutas, _ruta_err, cal, sap_info)
        except Exception as e:
            self.after(0, self._mostrar_error, str(e)[:120])

    def _mostrar_resultado(self, orden, zfer, vitro1, vitro2, mallas,
                           vitro_ruta=None, malla_rutas=None,
                           ruta_err=None, cal=None, sap_info=None):
        self._prog.stop(); self._prog.set(0)
        self._entry.configure(state="normal")
        malla_rutas = malla_rutas or {}
        cal = cal or {}
        sap_info = sap_info or {}

        # Chips
        self._c_orden.configure(text=str(orden))
        self._c_zfer.configure(text=str(zfer))

        # Tarjeta calendario
        logo     = cal.get("logo", "")
        vehiculo = cal.get("vehiculo", "—")
        version  = cal.get("version", "")
        lote     = cal.get("lote", "—")
        cliente  = cal.get("cliente", "—")
        self._lbl_logo.configure(
            text=logo if logo else " ",
            fg_color=PAL["accent"] if logo else PAL["border"],
        )
        self._lbl_vehiculo.configure(text=vehiculo)
        self._lbl_version.configure(
            text=f"Versión: {version}" if version else "")
        self._lbl_lote.configure(text=lote)
        self._lbl_version2.configure(text=version if version else "—")

        # Info SAP
        tipo_raw  = sap_info.get("Z_PIECE_TYPE", "")
        tipo_txt  = tipo_raw if tipo_raw else "—"
        traz_raw  = sap_info.get("Z_TRAZABILITY", "")
        traz_txt  = traz_raw if traz_raw else "—"
        cod_veh   = cal.get("cod_veh", "—") or "—"
        self._lbl_sap_codveh.configure(text=cod_veh)
        self._lbl_sap_tipo.configure(text=tipo_txt)
        self._lbl_trazabilidad.configure(text=traz_txt)
        self._lbl_mallas_count.configure(text=str(len(mallas)) if mallas else "")

        # Vitro
        self._lbl_vitro2.configure(text=vitro2 or "—")
        self._vitro_ruta_val = vitro_ruta or ""
        if ruta_err:
            self._lbl_vitro_ruta.configure(
                text=f"Error BD: {ruta_err[:80]}", text_color="#f85149")
        elif vitro_ruta:
            self._lbl_vitro_ruta.configure(
                text=vitro_ruta, text_color="#3fb950")
        else:
            self._lbl_vitro_ruta.configure(
                text="— sin ruta registrada —", text_color="#484f58")

        # Calcular si hay rutas distintas (por directorio)
        import os.path as _op
        _vd = _op.dirname(vitro_ruta).lower().rstrip("/\\") if vitro_ruta else None
        def _misma_dir(r):
            return _vd and r and _op.dirname(r).lower().rstrip("/\\") == _vd

        # Actualizar label de vitro ruta según si es "común" o solo de vitro
        rutas_distintas_nombres = [m for m, r in malla_rutas.items()
                                   if r and not _misma_dir(r)]
        if vitro_ruta and not rutas_distintas_nombres:
            todas_rutas = [vitro_ruta] + [r for r in malla_rutas.values() if r]
            if len([r for r in malla_rutas.values() if r]) > 0:
                self._lbl_vitro_ruta.configure(
                    text=vitro_ruta, text_color=PAL["txt"])
            else:
                self._lbl_vitro_ruta.configure(
                    text=vitro_ruta, text_color=PAL["txt"])

        # Mallas
        for w in self._mallas_labels:
            w.destroy()
        self._mallas_labels.clear()
        _nums = ["①","②","③","④","⑤","⑥","⑦","⑧","⑨","⑩",
                 "⑪","⑫","⑬","⑭","⑮","⑯","⑰","⑱","⑲","⑳"]
        for i, m in enumerate(mallas):
            mf = ctk.CTkFrame(self._mallas_box, fg_color="#0d1117",
                              corner_radius=12, border_width=1,
                              border_color="#21262d")
            mf.grid(row=i, column=0, sticky="ew", pady=(0, 6))
            mf.columnconfigure(1, weight=1)
            # Número badge
            num = _nums[i] if i < len(_nums) else str(i+1)
            ctk.CTkLabel(mf, text=num, font=FONT(22, "bold"),
                         text_color="#8957e5", width=40,
                         ).grid(row=0, column=0, padx=(12, 8), pady=(12, 8), sticky="n")
            # Nombre malla
            ctk.CTkLabel(mf, text=m, font=FONT(17, "bold"),
                         text_color="#e6edf3", anchor="w",
                         ).grid(row=0, column=1, sticky="w", pady=(12, 8), padx=(0, 12))
            ruta_m = malla_rutas.get(m)
            misma = _misma_dir(ruta_m)
            if misma:
                ctk.CTkLabel(
                    mf, text="✓  Misma carpeta que Vitro",
                    font=FONT(10), text_color="#3fb950", anchor="w",
                ).grid(row=1, column=0, columnspan=2, sticky="w",
                       padx=(10, 10), pady=(0, 10))
            else:
                rc = ctk.CTkFrame(mf, fg_color="#161b22", corner_radius=8,
                                   border_width=1, border_color="#30363d")
                rc.grid(row=1, column=0, columnspan=2, sticky="ew",
                        padx=8, pady=(0, 8))
                rc.columnconfigure(0, weight=1)
                ctk.CTkLabel(
                    rc,
                    text=ruta_m if ruta_m else "— sin ruta registrada —",
                    font=MONO(10),
                    text_color="#e6edf3" if ruta_m else "#484f58",
                    anchor="w", wraplength=270,
                ).grid(row=0, column=0, sticky="w", padx=10, pady=6)
                if ruta_m:
                    def _mk(r=ruta_m): return lambda: self._copiar(r)
                    ctk.CTkButton(
                        rc, text="📋", width=36, height=28,
                        font=FONT(12), corner_radius=6,
                        fg_color="#0d2218", hover_color="#238636",
                        text_color="#3fb950", command=_mk(),
                    ).grid(row=0, column=1, padx=(2, 8), pady=4)
            self._mallas_labels.append(mf)

        # Alerta rutas
        self._alerta_frame.grid_remove()
        self._alerta_txt = ""
        if vitro_ruta and rutas_distintas_nombres:
            self._alerta_txt = (
                f"RUTAS DISTINTAS — Vitro: {_op.dirname(vitro_ruta)} | "
                f"Malla(s): {', '.join(rutas_distintas_nombres)}")
            self._alerta_lbl.configure(text=f"⚠  {self._alerta_txt}")
            self._alerta_frame.grid()

        self._zona.grid()
        self._entry.delete(0, "end")
        self._entry.focus_set()

    def _mostrar_error(self, msg):
        self._prog.stop(); self._prog.set(0)
        self._entry.configure(state="normal", border_color=PAL["red"])
        self.after(2000, lambda: self._entry.configure(border_color=PAL["accent"]))
        self._entry.select_range(0, "end")

        for w in self._mallas_labels:
            w.destroy()
        self._mallas_labels.clear()
        err = ctk.CTkLabel(
            self._mallas_box, text=f"  ✘  {msg}",
            font=FONT(14, "bold"), text_color=PAL["red"],
            fg_color=PAL["card"], corner_radius=6, anchor="w",
        )
        err.grid(row=0, column=0, sticky="ew", pady=2)
        self._mallas_labels.append(err)
        self._c_orden.configure(text="—")
        self._c_zfer.configure(text="—")
        self._lbl_vitro2.configure(text="No encontrado")
        self._vitro_ruta_val = ""
        self._lbl_logo.configure(text=" ", fg_color=PAL["border"])
        self._lbl_vehiculo.configure(text="—")
        self._lbl_version.configure(text="")
        self._lbl_lote.configure(text="—")
        self._lbl_version2.configure(text="—")
        self._lbl_sap_codveh.configure(text="—")
        self._lbl_sap_tipo.configure(text="—")
        self._lbl_trazabilidad.configure(text="—")
        self._lbl_mallas_count.configure(text="")
        self._alerta_frame.grid_remove()
        self._alerta_txt = ""
        self._zona.grid()


# ══════════════════════════════════════════════════════════════════════════════
#  PESTAÑA — ROLES Y PERMISOS  (solo admin)
# ══════════════════════════════════════════════════════════════════════════════
class TabRoles(ctk.CTkFrame):
    ROLES = ["admin", "dibujante", "planta", "— sin rol —"]
    COLS  = ("Nombre", "Usuario", "Rol", "Estatus")
    ANCHOS = (220, 240, 110, 80)

    def __init__(self, parent, usuario_info=None, **kw):
        super().__init__(parent, fg_color="transparent", **kw)
        self._usuario = usuario_info or {}
        self._rows = []
        self._sel_id = None
        self._build()

    def _build(self):
        self.columnconfigure(0, weight=1)
        self.rowconfigure(1, weight=1)

        # ── barra superior ──────────────────────────────────────────────────
        top = ctk.CTkFrame(self, fg_color=PAL["card"], corner_radius=10)
        top.grid(row=0, column=0, sticky="ew", padx=20, pady=(16,8))
        ctk.CTkLabel(top, text="Gestión de Roles", font=FONT(15,"bold"),
                     text_color=PAL["txt"]).pack(side="left", padx=16, pady=10)
        ctk.CTkButton(top, text="↺  Recargar", width=110, height=32,
                      fg_color=PAL["card2"], hover_color=PAL["border"],
                      font=FONT(11), command=self.refresh
                      ).pack(side="right", padx=12, pady=10)

        # ── cuerpo ──────────────────────────────────────────────────────────
        body = ctk.CTkFrame(self, fg_color="transparent")
        body.grid(row=1, column=0, sticky="nsew", padx=20, pady=(0,16))
        body.columnconfigure(0, weight=1)
        body.columnconfigure(1, weight=0)
        body.rowconfigure(0, weight=1)

        # Treeview
        tv_frame = ctk.CTkFrame(body, fg_color=PAL["card"], corner_radius=10)
        tv_frame.grid(row=0, column=0, sticky="nsew", padx=(0,10))
        tv_frame.rowconfigure(0, weight=1)
        tv_frame.columnconfigure(0, weight=1)

        style = __import__("tkinter.ttk", fromlist=["Style"]).Style()
        style.configure("Roles.Treeview",
                        background=PAL["card"], foreground=PAL["txt"],
                        fieldbackground=PAL["card"], rowheight=34,
                        borderwidth=0, font=("Segoe UI", 10))
        style.configure("Roles.Treeview.Heading",
                        background=PAL["card2"], foreground=PAL["txt_mid"],
                        font=("Segoe UI", 10, "bold"), relief="flat")
        style.map("Roles.Treeview", background=[("selected", PAL["accent"])])

        import tkinter.ttk as ttk
        self._tree = ttk.Treeview(tv_frame, style="Roles.Treeview",
                                  columns=self.COLS, show="headings",
                                  selectmode="browse")
        for col, w in zip(self.COLS, self.ANCHOS):
            self._tree.heading(col, text=col)
            self._tree.column(col, width=w, minwidth=60,
                              anchor="center" if col in ("Rol","Estatus") else "w")
        sb = ctk.CTkScrollbar(tv_frame, command=self._tree.yview)
        self._tree.configure(yscrollcommand=sb.set)
        self._tree.grid(row=0, column=0, sticky="nsew")
        sb.grid(row=0, column=1, sticky="ns")
        self._tree.bind("<<TreeviewSelect>>", self._on_select)
        self._tree.bind("<MouseWheel>",
                        lambda e: self._tree.yview_scroll(int(-e.delta / 120), "units"))

        # Panel editor lateral
        panel = ctk.CTkFrame(body, fg_color=PAL["card"], corner_radius=10, width=240)
        panel.grid(row=0, column=1, sticky="nsew")
        panel.grid_propagate(False)
        panel.columnconfigure(0, weight=1)

        ctk.CTkLabel(panel, text="Editar usuario", font=FONT(13,"bold"),
                     text_color=PAL["txt"]).pack(pady=(20,4), padx=16, anchor="w")
        ctk.CTkFrame(panel, fg_color=PAL["border"], height=1).pack(fill="x", padx=12)

        ctk.CTkLabel(panel, text="Nombre", font=FONT(10),
                     text_color=PAL["txt_mid"]).pack(padx=16, pady=(14,2), anchor="w")
        self._lbl_nombre = ctk.CTkLabel(panel, text="—", font=FONT(11,"bold"),
                                        text_color=PAL["txt"], wraplength=200,
                                        justify="left")
        self._lbl_nombre.pack(padx=16, anchor="w")

        ctk.CTkLabel(panel, text="Usuario", font=FONT(10),
                     text_color=PAL["txt_mid"]).pack(padx=16, pady=(10,2), anchor="w")
        self._lbl_usuario = ctk.CTkLabel(panel, text="—", font=FONT(10),
                                          text_color=PAL["txt_mid"])
        self._lbl_usuario.pack(padx=16, anchor="w")

        ctk.CTkLabel(panel, text="Rol", font=FONT(10),
                     text_color=PAL["txt_mid"]).pack(padx=16, pady=(16,4), anchor="w")
        self._combo_rol = ctk.CTkComboBox(panel, values=self.ROLES,
                                           width=200, height=36,
                                           fg_color=PAL["card2"],
                                           border_color=PAL["border"],
                                           button_color=PAL["accent"],
                                           font=FONT(12))
        self._combo_rol.pack(padx=16)
        self._combo_rol.set("— sin rol —")

        ctk.CTkLabel(panel, text="Estatus", font=FONT(10),
                     text_color=PAL["txt_mid"]).pack(padx=16, pady=(14,4), anchor="w")
        self._switch_est = ctk.CTkSwitch(panel, text="Activo",
                                          font=FONT(11),
                                          fg_color=PAL["border"],
                                          progress_color=PAL["ok"] if hasattr(PAL,"ok") else "#22C55E")
        self._switch_est.pack(padx=16, anchor="w")
        self._switch_est.select()

        self._lbl_msg = ctk.CTkLabel(panel, text="", font=FONT(11),
                                      text_color=PAL["accent"], wraplength=200)
        self._lbl_msg.pack(padx=16, pady=(10,0))

        ctk.CTkButton(panel, text="💾  Guardar cambios", height=40,
                      corner_radius=10,
                      fg_color=PAL["accent"], hover_color=PAL["accent2"] if "accent2" in PAL else "#2563EB",
                      font=FONT(12,"bold"),
                      command=self._guardar
                      ).pack(padx=16, pady=(16,8), fill="x")

        self.after(200, self.refresh)

    # ── datos ────────────────────────────────────────────────────────────────
    def refresh(self):
        self._tree.delete(*self._tree.get_children())
        self._rows.clear()
        try:
            cn = db_connect(); cur = cn.cursor()
            cur.execute(
                "SELECT id, nombre, usuario, rol, estatus "
                "FROM MALLAS.APP_USUARIOS ORDER BY nombre"
            )
            for row in cur.fetchall():
                rid, nombre, usuario, rol, estatus = row
                rol_disp = rol or "— sin rol —"
                est_disp = "Activo" if estatus else "Inactivo"
                iid = self._tree.insert("", "end",
                                        values=(nombre, usuario, rol_disp, est_disp))
                self._rows.append({"iid": iid, "id": rid, "nombre": nombre,
                                   "usuario": usuario, "rol": rol, "estatus": estatus})
            cn.close()
        except Exception as e:
            self._lbl_msg.configure(text=f"Error BD: {e}", text_color="#EF4444")

    def _on_select(self, _=None):
        sel = self._tree.selection()
        if not sel: return
        iid = sel[0]
        u = next((r for r in self._rows if r["iid"] == iid), None)
        if not u: return
        self._sel_id = u["id"]
        self._lbl_nombre.configure(text=u["nombre"] or "—")
        self._lbl_usuario.configure(text=u["usuario"] or "—")
        self._combo_rol.set(u["rol"] or "— sin rol —")
        if u["estatus"]: self._switch_est.select()
        else: self._switch_est.deselect()
        self._lbl_msg.configure(text="")

    def _guardar(self):
        if not self._sel_id:
            self._lbl_msg.configure(text="Selecciona un usuario", text_color="#EF4444")
            return
        rol_val  = self._combo_rol.get()
        rol_sql  = None if rol_val == "— sin rol —" else rol_val
        estatus  = 1 if self._switch_est.get() else 0
        try:
            cn = db_connect(); cur = cn.cursor()
            cur.execute(
                "UPDATE MALLAS.APP_USUARIOS SET rol=%s, estatus=%s, "
                "actualizado_en=SYSDATETIME() WHERE id=%s",
                (rol_sql, estatus, self._sel_id)
            )
            cn.commit(); cn.close()
            self._lbl_msg.configure(text="✔ Guardado", text_color="#22C55E")
            self.refresh()
        except Exception as e:
            self._lbl_msg.configure(text=f"Error: {e}", text_color="#EF4444")


# ══════════════════════════════════════════════════════════════════════════════
#  APP PRINCIPAL
# ══════════════════════════════════════════════════════════════════════════════
_PERMISOS = {
    # nombre_pestaña → roles que pueden verla
    "Crear Arte":       {"admin", "dibujante"},
    "Consultar BD":     {"admin", "dibujante"},
    "Gestión BD":       {"admin", "dibujante"},
    "Scanner":          {"admin", "dibujante", "planta"},
    "Roles y Permisos": {"admin"},
}

class AGPApp(ctk.CTk):
    PAGES = [
        ("Crear Arte",       "🎨", TabArte),
        ("Consultar BD",     "🔍", TabBD),
        ("Gestión BD",       "✏", TabGestion),
        ("Scanner",          "📷", TabScanner),
        ("Roles y Permisos", "👥", TabRoles),
    ]

    def __init__(self, usuario_info=None):
        super().__init__()
        self._usuario = usuario_info or {}
        self._rol     = (self._usuario.get("rol") or "").lower()
        self.title("AGP Glass — Suite")
        self.geometry("1280x820")
        self.minsize(1000, 680)
        self._active = None
        self._frames = {}
        self._build()
        threading.Thread(target=self._limpiar_pendientes, daemon=True).start()
        threading.Thread(target=self._test_conexion, daemon=True).start()

    def _test_conexion(self):
        """Prueba la conexión al arrancar y muestra popup claro si falla."""
        import time
        time.sleep(1.5)  # esperar que la UI cargue
        try:
            db_connect().close()
        except Exception as e:
            msg = str(e)
            self.after(0, lambda: self._mostrar_error_bd(msg))

    def _mostrar_error_bd(self, msg):
        import tkinter as _tk
        win = _tk.Toplevel(self)
        win.title("Sin conexión a la base de datos")
        win.configure(bg="#1e1e2e")
        win.resizable(False, False)
        win.attributes("-topmost", True)
        win.grab_set()

        _tk.Frame(win, bg="#ef4444", height=5).pack(fill="x")
        _tk.Label(win, text="⚠  Sin conexión a la base de datos",
                  font=("Segoe UI", 14, "bold"), fg="#fca5a5", bg="#1e1e2e"
                  ).pack(padx=28, pady=(18, 6), anchor="w")

        _tk.Label(win, text=msg, font=("Segoe UI", 10), fg="#94a3b8",
                  bg="#1e1e2e", wraplength=440, justify="left"
                  ).pack(padx=28, pady=(0, 14), anchor="w")

        _tk.Frame(win, bg="#2d2d44", height=1).pack(fill="x", padx=20, pady=(0,12))

        soluciones = (
            "Posibles causas:\n"
            "  1. Falta el driver ODBC  →  ejecuta INSTALAR_DRIVER.bat\n"
            "  2. Sin internet o red interna\n"
            "  3. Firewall bloqueando el puerto 1433"
        )
        _tk.Label(win, text=soluciones, font=("Segoe UI", 10),
                  fg="#e2e8f0", bg="#1e1e2e", justify="left"
                  ).pack(padx=28, pady=(0, 20), anchor="w")

        _tk.Button(win, text="  Cerrar  ", font=("Segoe UI", 10, "bold"),
                   bg="#3b82f6", fg="white", relief="flat",
                   activebackground="#2563eb", cursor="hand2",
                   command=win.destroy
                   ).pack(pady=(0, 20))

        win.update_idletasks()
        x = self.winfo_rootx() + self.winfo_width()  // 2 - win.winfo_reqwidth()  // 2
        y = self.winfo_rooty() + self.winfo_height() // 2 - win.winfo_reqheight() // 2
        win.geometry(f"+{x}+{y}")

    def _limpiar_pendientes(self):
        try:
            if _ASIGN_OK:
                nv, ng, np_ = _limpiar_pendientes_huerfanos()
                if nv + ng + np_ > 0:
                    print(f"[startup] Pendientes huerfanos cancelados: {nv} vitros, {ng} grandes, {np_} pequenas")
        except Exception as e:
            print(f"[startup] WARN limpiar pendientes: {e}")

    def _build(self):
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        # ── SIDEBAR ───────────────────────────────────────────────────────────
        sidebar = ctk.CTkFrame(self, width=240, fg_color=PAL["sidebar"],
                               corner_radius=0)
        sidebar.grid(row=0, column=0, sticky="nsew")
        sidebar.grid_propagate(False)

        # Logo
        logo_frame = ctk.CTkFrame(sidebar, fg_color="transparent")
        logo_frame.pack(fill="x", padx=16, pady=(24,8))
        ctk.CTkLabel(logo_frame, text="AGP", font=FONT(13, "bold"),
                     text_color=PAL["txt_dim"]).pack(anchor="w")
        ctk.CTkLabel(logo_frame, text="Glass Suite", font=FONT(24, "bold"),
                     text_color=PAL["accent"]).pack(anchor="w")
        # 2px accent line below logo
        ctk.CTkFrame(sidebar, fg_color=PAL["accent"], height=2
                     ).pack(fill="x", padx=12, pady=(4,10))

        # Nav buttons — filtrados por rol del usuario actual
        self._nav_btns = {}
        self._pages_activas = [
            (name, icon, cls) for name, icon, cls in self.PAGES
            if self._rol in _PERMISOS.get(name, set())
        ]
        for name, icon, cls in self._pages_activas:
            b = SideBtn(sidebar, text=name, icon=icon,
                        command=lambda n=name: self._show(n))
            b.pack(fill="x", padx=8, pady=2)
            self._nav_btns[name] = b

        # Info de usuario en sidebar
        ctk.CTkFrame(sidebar, fg_color=PAL["border"], height=1
                     ).pack(fill="x", padx=12, pady=(12,6))
        nombre_corto = (self._usuario.get("nombre") or "").split()[0] if self._usuario.get("nombre") else ""
        ctk.CTkLabel(sidebar, text=f"👤  {nombre_corto}",
                     font=FONT(11,"bold"), text_color=PAL["txt"]
                     ).pack(anchor="w", padx=16, pady=(0,2))
        ctk.CTkLabel(sidebar, text=self._rol.capitalize() if self._rol else "Sin rol",
                     font=FONT(10), text_color=PAL["accent"]
                     ).pack(anchor="w", padx=16)

        # Section label below nav buttons
        ctk.CTkFrame(sidebar, fg_color=PAL["border"], height=1
                     ).pack(fill="x", padx=12, pady=(12,4))
        ctk.CTkLabel(sidebar, text="AGP GLASS", font=FONT(9, "bold"),
                     text_color=PAL["txt_dim"]).pack(anchor="w", padx=16)
        ctk.CTkLabel(sidebar, text="Suite Ingeniería", font=FONT(9),
                     text_color=PAL["txt_dim"]).pack(anchor="w", padx=16, pady=(0,8))

        # Footer
        ctk.CTkLabel(sidebar, text="v2.0  ·  Ingeniería",
                     font=FONT(9), text_color=PAL["txt_dim"]
                     ).pack(side="bottom", pady=8)
        ctk.CTkFrame(sidebar, fg_color=PAL["border"], height=1
                     ).pack(fill="x", padx=12, pady=2, side="bottom")

        # Acelerador de teclas
        ctk.CTkLabel(sidebar, text="Alt+1 Arte  ·  Alt+2 BD\nAlt+3 Gestión  ·  Alt+4 Scanner",
                     font=FONT(9), text_color=PAL["txt_dim"], justify="center"
                     ).pack(side="bottom", pady=2)

        # ── CONTENIDO ─────────────────────────────────────────────────────────
        content_wrapper = ctk.CTkFrame(self, fg_color=PAL["bg"], corner_radius=0)
        content_wrapper.grid(row=0, column=1, sticky="nsew")
        content_wrapper.grid_columnconfigure(0, weight=1)
        content_wrapper.grid_rowconfigure(1, weight=1)

        # Header title bar
        self._header_frame = ctk.CTkFrame(content_wrapper, fg_color=PAL["card"],
                                           corner_radius=0)
        self._header_frame.grid(row=0, column=0, sticky="ew")
        header_inner = ctk.CTkFrame(self._header_frame, fg_color="transparent")
        header_inner.pack(fill="x", padx=24, pady=(16, 0))
        self._header_title = ctk.CTkLabel(header_inner, text="",
                                           font=FONT(22, "bold"),
                                           text_color=PAL["txt"])
        self._header_title.pack(anchor="w")
        self._header_sub = ctk.CTkLabel(header_inner, text="",
                                         font=FONT(12),
                                         text_color=PAL["txt_mid"])
        self._header_sub.pack(anchor="w", pady=(2, 0))
        ctk.CTkFrame(self._header_frame, fg_color=PAL["border"], height=1
                     ).pack(fill="x", padx=0, pady=(12, 0))

        # Contenedor plano para stacking de pestañas con place()
        self._content = ctk.CTkFrame(content_wrapper, fg_color=PAL["bg"], corner_radius=0)
        self._content.grid(row=1, column=0, sticky="nsew", padx=0, pady=0)
        self._content.bind("<Configure>", self._on_content_resize)

        # Header top strip
        ctk.CTkFrame(self, fg_color=PAL["accent"], height=3
                     ).grid(row=0, column=0, columnspan=2, sticky="new")

        # Instanciar solo las páginas activas para este rol
        for name, icon, cls in self._pages_activas:
            kwargs = {}
            if cls in (TabGestion, TabRoles):
                kwargs["usuario_info"] = self._usuario
            f = cls(self._content, **kwargs)
            f.place(x=0, y=0, relwidth=1, relheight=1)
            f.place_forget()
            self._frames[name] = f

        # Conectar callbacks entre pestañas (solo si existen)
        if "Consultar BD" in self._frames and "Gestión BD" in self._frames:
            def _refresh_bd():
                self._frames["Consultar BD"].refresh()
                self._frames["Gestión BD"].refresh()
            self._frames["Crear Arte"]._on_art_done     = _refresh_bd if "Crear Arte" in self._frames else None
            self._frames["Gestión BD"]._on_data_changed = self._frames["Consultar BD"].refresh

        # Atajos de teclado dinámicos
        nombres = [n for n, _, _ in self._pages_activas]
        for i, n in enumerate(nombres, 1):
            self.bind(f"<Alt-Key-{i}>", lambda _, name=n: self._show(name))

        primera = nombres[0] if nombres else None
        if primera:
            self._show(primera)

    _PAGE_SUBTITLES = {
        "Crear Arte":       "Pipeline AutoCAD — extraer plano, crear y buscar artes",
        "Consultar BD":     "Base de datos Azure SQL — vitros, mallas, vinilos, pasta plata",
        "Gestión BD":       "Editar, separar e insertar registros en la BD",
        "Scanner":          "Escáner de órdenes de producción — barcode → SmartFactory → SAP",
        "Roles y Permisos": "Gestión de roles y accesos de usuarios",
    }

    def _on_content_resize(self, e):
        for f in self._frames.values():
            try: f.place_configure(width=e.width, height=e.height)
            except Exception: pass

    def _show(self, name):
        if self._active:
            self._frames[self._active].place_forget()
            self._nav_btns[self._active].set_active(False)
        self._frames[name].place(x=0, y=0, relwidth=1, relheight=1)
        self._frames[name].lift()
        self._nav_btns[name].set_active(True)
        self._active = name
        self._header_title.configure(text=name)
        self._header_sub.configure(text=self._PAGE_SUBTITLES.get(name, ""))
        # Forzar redibujado del layout antes de cargar datos (evita vista compacta)
        self.update_idletasks()
        # Refrescar BD al entrar a esas pestañas (datos siempre actualizados)
        if name in ("Consultar BD", "Gestión BD"):
            self._frames[name].refresh()

# ══════════════════════════════════════════════════════════════════════════════
if __name__ == "__main__":
    # ── Login ──────────────────────────────────────────────────────────────
    _USUARIO_ACTUAL = None
    try:
        from login_window import LoginWindow
        login = LoginWindow()
        login.mainloop()
        if login.usuario_info is None:
            sys.exit(0)
        _USUARIO_ACTUAL = login.usuario_info
    except Exception as _le:
        print(f"[login] error: {_le}")

    # ── Sin rol asignado ────────────────────────────────────────────────────
    if _USUARIO_ACTUAL and not _USUARIO_ACTUAL.get("rol"):
        _pal = {"bg": "#0F1117", "card": "#1C2333", "accent": "#3B82F6",
                "txt": "#F1F5F9", "txt_mid": "#94A3B8"}
        _win = ctk.CTk()
        _win.title("AGP Glass")
        _win.geometry("460x300")
        _win.resizable(False, False)
        _win.configure(fg_color=_pal["bg"])
        _win.update_idletasks()
        _wx = (_win.winfo_screenwidth()  - 460) // 2
        _wy = (_win.winfo_screenheight() - 300) // 2
        _win.geometry(f"460x300+{_wx}+{_wy}")
        _card = ctk.CTkFrame(_win, fg_color=_pal["card"], corner_radius=16)
        _card.place(relx=0.5, rely=0.5, anchor="center", relwidth=0.88, relheight=0.82)
        ctk.CTkLabel(_card, text="⚠", font=CTkFont(size=40)).pack(pady=(28,4))
        ctk.CTkLabel(_card, text="Sin rol asignado",
                     font=CTkFont(size=18, weight="bold"),
                     text_color=_pal["txt"]).pack()
        _nom = (_USUARIO_ACTUAL.get("nombre") or "").split()[0]
        ctk.CTkLabel(_card,
                     text=f"Hola {_nom}, aún no tienes ningún rol.\nHabla con un administrador para que te asigne acceso.",
                     font=CTkFont(size=12), text_color=_pal["txt_mid"],
                     justify="center", wraplength=340).pack(pady=(10,0))
        ctk.CTkButton(_card, text="Cerrar", width=120, height=36,
                      fg_color=_pal["accent"], corner_radius=10,
                      command=_win.destroy).pack(pady=(18,0))
        _win.mainloop()
        sys.exit(0)

    # ── App principal ───────────────────────────────────────────────────────
    AGPApp(usuario_info=_USUARIO_ACTUAL).mainloop()
