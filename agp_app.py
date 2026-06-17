#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
AGP Glass — App unificada
  · Crear Arte    (pipeline AutoCAD)
  · Consultar BD  (Azure SQL — vitros, mallas, vinilos, pasta plata)
Requiere: customtkinter, pyodbc, pywin32
"""

import os, sys, time, threading, subprocess

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
    from crear_arte_acad import dialogo_cajetin, pipeline as _pipeline_acad
    _PIPELINE_OK = True
except Exception:
    _PIPELINE_OK = False

import re, math, shutil, json, datetime
from tkinter import filedialog, messagebox

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
    "bg":        "#0B0D17",
    "sidebar":   "#0F1220",
    "card":      "#131626",
    "card2":     "#171B2E",
    "border":    "#1E2540",
    "accent":    "#4FACFF",
    "accent2":   "#2980CC",
    "green":     "#50FA7B",
    "green2":    "#27AE60",
    "orange":    "#FFB86C", 
    "red":       "#FF5555",
    "purple":    "#8B5CF6",
    "txt":       "#ECF0FF",
    "txt_mid":   "#7A90B0",
    "txt_dim":   "#3A4A6A",
    "log_bg":    "#080A14",
}

FONT = lambda s, w="normal": CTkFont(family="Segoe UI", size=s, weight=w)
MONO = lambda s=11: CTkFont(family="Consolas", size=s)

# ══════════════════════════════════════════════════════════════════════════════
#  BD — helpers
# ══════════════════════════════════════════════════════════════════════════════
def db_connect():
    if pyodbc is None:
        raise RuntimeError("pyodbc no instalado")
    for drv in ["ODBC Driver 18 for SQL Server", "ODBC Driver 17 for SQL Server"]:
        try:
            cs = CONN_AZURE.replace("ODBC Driver 17 for SQL Server", drv)
            return pyodbc.connect(cs, timeout=20)
        except Exception:
            continue
    raise RuntimeError("No se pudo conectar a Azure SQL")

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

def _ruta_arte_salida(ruta_dwg, malla="", pieza=""):
    artes = os.path.join(os.path.dirname(os.path.abspath(ruta_dwg)), "ARTES")
    os.makedirs(artes, exist_ok=True)
    dest = artes
    try:
        for e in os.listdir(artes):
            if os.path.isdir(os.path.join(artes, e)) and e.upper() == "BN":
                dest = os.path.join(artes, e); break
    except Exception: pass
    partes = [p.strip() for p in [malla, pieza] if p.strip()]
    nombre = ("P " + " ".join(partes) if partes else
              "P " + os.path.splitext(os.path.basename(ruta_dwg))[0]) + ".dwg"
    while "  " in nombre: nombre = nombre.replace("  ", " ")
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
        time.sleep(2)
        try: doc.Activate(); time.sleep(0.5)
        except Exception: pass
        n = _pipeline_acad(doc, log_fn=log_fn, valores_cajetin=valores_cajetin,
                           ruta_salida=ruta_salida, perim_index=perim_index,
                           compensar=compensar)
        return n or 1
    finally:
        if not _com_ya_init:
            pythoncom.CoUninitialize()

# ══════════════════════════════════════════════════════════════════════════════
#  WIDGETS REUTILIZABLES
# ══════════════════════════════════════════════════════════════════════════════
class SideBtn(ctk.CTkButton):
    """Botón de navegación del sidebar."""
    def __init__(self, parent, text, icon, command, **kw):
        super().__init__(parent,
            text=f"  {icon}  {text}",
            anchor="w",
            height=44,
            corner_radius=8,
            fg_color="transparent",
            hover_color=PAL["border"],
            text_color=PAL["txt_mid"],
            font=FONT(13),
            command=command, **kw)

    def set_active(self, active: bool):
        if active:
            self.configure(fg_color=PAL["accent2"], text_color="white",
                           font=FONT(13, "bold"))
        else:
            self.configure(fg_color="transparent", text_color=PAL["txt_mid"],
                           font=FONT(13))


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
        super().__init__(parent, fg_color=PAL["card"], corner_radius=10,
                         border_width=1, border_color=PAL["border"], **kw)
        ctk.CTkLabel(self, text=icon, font=FONT(20)).pack(pady=(12,0))
        self._val = ctk.CTkLabel(self, text="—", font=FONT(22, "bold"),
                                  text_color=color)
        self._val.pack()
        ctk.CTkLabel(self, text=label, font=FONT(10),
                     text_color=PAL["txt_mid"]).pack(pady=(0,10))

    def set(self, v): self._val.configure(text=f"{v:,}" if isinstance(v, int) else str(v))


# ══════════════════════════════════════════════════════════════════════════════
#  PESTAÑA — CREAR ARTE
# ══════════════════════════════════════════════════════════════════════════════
class TabArte(ctk.CTkFrame):
    def __init__(self, parent, **kw):
        super().__init__(parent, fg_color="transparent", **kw)
        self._ruta_base  = ctk.StringVar()
        self._ruta_dwg   = ctk.StringVar()
        self._compensar  = ctk.BooleanVar(value=False)
        self._resultados = []
        self._build()

    def _build(self):
        self.columnconfigure(0, weight=1)

        # ── Sección inputs ────────────────────────────────────────────────────
        card_in = self._card("CONFIGURACIÓN", row=0)
        card_in.columnconfigure(1, weight=1)

        self._field(card_in, "Ruta vehículo / versión:", 0)
        self._e_base = ctk.CTkEntry(card_in, textvariable=self._ruta_base,
                                     height=36, font=FONT(12),
                                     fg_color=PAL["card2"], border_color=PAL["border"],
                                     text_color=PAL["accent"])
        self._e_base.grid(row=1, column=0, columnspan=2, sticky="ew", padx=(0,6), pady=(2,8))
        ctk.CTkButton(card_in, text="📂 Explorar", width=110, height=36,
                      fg_color=PAL["border"], hover_color=PAL["accent2"],
                      font=FONT(11), command=self._pick_base
                      ).grid(row=1, column=2, pady=(2,8))

        self._field(card_in, "Plano DWG original:", 2)
        self._e_dwg = ctk.CTkEntry(card_in, textvariable=self._ruta_dwg,
                                    height=36, font=FONT(12),
                                    fg_color=PAL["card2"], border_color=PAL["border"],
                                    text_color=PAL["accent"])
        self._e_dwg.grid(row=3, column=0, columnspan=2, sticky="ew", padx=(0,6), pady=(2,4))
        ctk.CTkButton(card_in, text="📂 Explorar", width=110, height=36,
                      fg_color=PAL["border"], hover_color=PAL["accent2"],
                      font=FONT(11), command=self._pick_dwg
                      ).grid(row=3, column=2, pady=(2,4))

        # ── Botones workflow ──────────────────────────────────────────────────
        card_wf = self._card("WORKFLOW", row=1)
        btn_row = ctk.CTkFrame(card_wf, fg_color="transparent")
        btn_row.pack(fill="x", pady=6)

        self._btns = []
        defs = [
            ("▶  Extraer Plano",  PAL["green2"],   "#1a6640", self._extraer),
            ("✦  Crear Arte",     PAL["purple"],   "#5b21b6", self._crear_arte),
            ("⊕  Buscar Arte",   "#E67E22",        "#ca6f1e", self._buscar),
            ("⚡  Todo en Uno",   "#E63946",        "#b71c2e", self._todo_en_uno),
        ]
        for txt, color, hover, cmd in defs:
            b = ctk.CTkButton(btn_row, text=txt, width=168, height=46,
                              fg_color=color, hover_color=hover,
                              font=FONT(12, "bold"), corner_radius=8,
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
                                         height=4, progress_color=PAL["accent"])
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
        style.configure("AGP.Treeview", background=PAL["card"],
                        foreground=PAL["txt_mid"], fieldbackground=PAL["card"],
                        borderwidth=0, font=("Segoe UI", 10), rowheight=28)
        style.configure("AGP.Treeview.Heading", background=PAL["card2"],
                        foreground=PAL["accent"], font=("Segoe UI", 9, "bold"),
                        relief="flat")
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
        self._log = LogBox(card_log, height=160)
        self._log.pack(fill="both", expand=True, pady=(4,0))

        self.rowconfigure(3, weight=1)

    # ── helpers ──────────────────────────────────────────────────────────────
    def _card(self, title, row):
        outer = ctk.CTkFrame(self, fg_color=PAL["card"],
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
            messagebox.showwarning("Campo requerido", "Indica una ruta base válida.")
            return False
        if dwg and not os.path.isfile(self._ruta_dwg.get().strip().strip('"')):
            messagebox.showwarning("Campo requerido", "Selecciona el DWG del plano.")
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
                motor.quit()
                pythoncom.CoUninitialize()
            self._log_fn(f"Guardado → {dest}", "ok")
            subprocess.Popen(["explorer", "/select,", dest])
        except Exception as e:
            self._log_fn(str(e), "err")
        finally:
            self._busy(False)

    def _crear_arte(self):
        dwg = self._ruta_dwg.get().strip().strip('"')
        if not os.path.isfile(dwg):
            messagebox.showwarning("Campo requerido","Selecciona el DWG del plano."); return
        self._busy(True)
        threading.Thread(target=self._t_crear, args=(dwg,), daemon=True).start()

    def _t_crear(self, dwg):
        self._log_fn("="*50)
        self._log_fn("CREAR ARTE...", "ok")
        try:
            _crear_arte_autocad(dwg, log_fn=lambda m: self._log_fn(m,"dim"),
                                compensar=self._compensar.get())
            self._log_fn("Arte completado.", "ok")
        except Exception as e:
            self._log_fn(str(e), "err")
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
        self._busy(True)
        threading.Thread(target=self._t_todo, args=(dwg, valores), daemon=True).start()

    def _t_todo(self, dwg, valores):
        malla = valores.get("MALLA","").strip()
        pieza = valores.get("PIEZA","").strip()
        self._log_fn("="*50)
        self._log_fn("TODO EN UNO...", "ok")

        if not _MOTOR_OK or not _COM_OK or not _PIPELINE_OK:
            self._log_fn("Faltan dependencias: AutoCADMotor / pywin32 / pipeline", "err")
            self._busy(False)
            return

        nombre = os.path.splitext(os.path.basename(dwg))[0]
        plano  = os.path.join(_ruta_planos(dwg), f"{nombre}_PLANO.dwg")
        arte0  = _ruta_arte_salida(dwg, malla, pieza)

        # Un único CoInitialize para todo el proceso en este hilo
        pythoncom.CoInitialize()
        try:
            # ── 1. Extraer plano ──────────────────────────────────────────────
            motor = AutoCADMotor()  # AutoCADMotor.init también llama CoInitialize (ref count +1)
            try:
                motor.extraer_layers(dwg, plano, log_fn=lambda m: self._log_fn(m, "dim"))
            finally:
                motor.quit()        # libera referencia COM pero NO CoUninitialize
            self._log_fn("Plano extraído ✔", "ok")
            time.sleep(2.0)

            # ── 2. Crear arte piezas ──────────────────────────────────────────
            n = _crear_arte_autocad(plano, log_fn=lambda m: self._log_fn(m, "dim"),
                                    valores_cajetin=valores, ruta_salida=arte0,
                                    perim_index=0, _com_ya_init=True,
                                    compensar=self._compensar.get())
            self._log_fn(f"Arte guardado ✔  {os.path.basename(arte0)}", "ok")

            for i in range(1, n or 1):
                copia  = plano.replace("_PLANO.dwg", f"_PLANO_p{i+1}.dwg")
                arte_i = _ruta_arte_salida(dwg, malla, f"{pieza} {i+1}".strip())
                try:
                    shutil.copy2(plano, copia)
                    _crear_arte_autocad(copia, log_fn=lambda m: self._log_fn(m, "dim"),
                                        valores_cajetin=valores, ruta_salida=arte_i,
                                        perim_index=i, _com_ya_init=True,
                                        compensar=self._compensar.get())
                    self._log_fn(f"Arte {i+1} ✔  {os.path.basename(arte_i)}", "ok")
                except Exception as e:
                    self._log_fn(f"Pieza {i+1}: {e}", "warn")

        except Exception as e:
            self._log_fn(str(e), "err")
        finally:
            pythoncom.CoUninitialize()

        self._log_fn("Todo en Uno completado.", "ok")
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
            messagebox.showwarning("Plano requerido","Indica el plano DWG para superponer."); return
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
        ("Vitrojet",   "vitrojet",  "🔬"),
        ("Mallas G",   "grandes",   "🔷"),
        ("Mallas P",   "pequenas",  "🔹"),
        ("Vinilos",    "vinilos",   "🎨"),
        ("Pasta Plata","pasta",     "🪙"),
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
        "vitrojet":  ("SELECT TOP(?) v.vitro,v.codigo_malla,v.tipo_malla,v.bnerig,v.vehiculo,v.version "
                      "FROM mallas.vitrojet v {where} ORDER BY v.vitro DESC",
                      ["Vitro","Malla","Tipo","B/N","Vehículo","Versión"],
                      ["vitro","codigo_malla","tipo_malla","bnerig","vehiculo","version"]),
        "grandes":   ("SELECT TOP(?) codigo,cod_veh,descripcion,pieza,tipo,version "
                      "FROM mallas.grandes {where} ORDER BY codigo",
                      ["Código","Cód.Veh.","Descripción","Pieza","Tipo","Versión"],
                      ["codigo","cod_veh","descripcion","pieza","tipo","version"]),
        "pequenas":  ("SELECT TOP(?) codigo,cod_veh,descripcion,pieza,tipo,version "
                      "FROM mallas.pequenas {where} ORDER BY codigo",
                      ["Código","Cód.Veh.","Descripción","Pieza","Tipo","Versión"],
                      ["codigo","cod_veh","descripcion","pieza","tipo","version"]),
        "vinilos":   ("SELECT TOP(?) herramental,vehiculo,cod_vehiculo,version,pieza,tipo "
                      "FROM mallas.vinilos {where} ORDER BY herramental",
                      ["Herramental","Vehículo","Cód.Veh.","Versión","Pieza","Tipo"],
                      ["herramental","vehiculo","cod_vehiculo","version","pieza","tipo"]),
        "pasta":     ("SELECT TOP(?) consecutivo,tipo,vehiculo,cod_vehiculo,version,pieza,ruta_archivo,caso "
                      "FROM mallas.pasta_plata {where} ORDER BY consecutivo",
                      ["Consecutivo","Tipo","Vehículo","Cód.Veh.","Versión","Pieza","Ruta Archivo","Caso"],
                      ["consecutivo","tipo","vehiculo","cod_vehiculo","version","pieza","ruta_archivo","caso"]),
    }
    WHERE = {
        "vitrojet": "WHERE v.vitro LIKE ? OR v.vehiculo LIKE ? OR v.codigo_malla LIKE ?",
        "grandes":  "WHERE descripcion LIKE ? OR codigo LIKE ? OR cod_veh LIKE ?",
        "pequenas": "WHERE descripcion LIKE ? OR CAST(codigo AS NVARCHAR) LIKE ? OR cod_veh LIKE ?",
        "vinilos":  "WHERE vehiculo LIKE ? OR herramental LIKE ?",
        "pasta":    "WHERE vehiculo LIKE ? OR consecutivo LIKE ?",
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

        # ── Barra búsqueda + tabs ─────────────────────────────────────────────
        bar = ctk.CTkFrame(self, fg_color=PAL["card"], corner_radius=10,
                           border_width=1, border_color=PAL["border"])
        bar.grid(row=1, column=0, sticky="ew", padx=4, pady=4)
        bar_in = ctk.CTkFrame(bar, fg_color="transparent")
        bar_in.pack(fill="x", padx=12, pady=10)
        bar_in.columnconfigure(1, weight=1)

        ctk.CTkLabel(bar_in, text="🔍", font=FONT(14)
                     ).grid(row=0, column=0, padx=(0,6))
        self._search = ctk.CTkEntry(bar_in, placeholder_text="Buscar vehículo, código, malla...",
                                     height=38, font=FONT(12),
                                     fg_color=PAL["card2"], border_color=PAL["border"])
        self._search.grid(row=0, column=1, sticky="ew", padx=(0,10))
        self._search.bind("<KeyRelease>", self._on_key)

        tab_bar = ctk.CTkFrame(bar_in, fg_color=PAL["bg"], corner_radius=8)
        tab_bar.grid(row=0, column=2)
        self._tab_btns = {}
        for i, (lbl, key, icon) in enumerate(self.TABS):
            b = ctk.CTkButton(tab_bar, text=f"{icon} {lbl}", width=100, height=34,
                              corner_radius=6, font=FONT(11),
                              fg_color=PAL["accent2"] if key=="vitrojet" else "transparent",
                              hover_color=PAL["border"],
                              command=lambda k=key: self._set_tab(k))
            b.grid(row=0, column=i, padx=2, pady=3)
            self._tab_btns[key] = b

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
        frm.pack(fill="both", expand=True, padx=8, pady=8)

        self._tree = ttv.Treeview(frm, style="AGP.Treeview", show="headings", height=18)
        self._tree.pack(side="left", fill="both", expand=True)
        sb = ttv.Scrollbar(frm, orient="vertical", command=self._tree.yview)
        sb.pack(side="right", fill="y")
        self._tree.configure(yscrollcommand=sb.set)

        self._lbl_count = ctk.CTkLabel(card_tbl, text="",
                                        font=FONT(10), text_color=PAL["txt_dim"])
        self._lbl_count.pack(anchor="e", padx=12, pady=(0,6))

        self._build_tree_cols("vitrojet")
        self._do_search()

    def _build_tree_cols(self, tab):
        _, headers, _ = self.QUERIES[tab]
        self._tree.configure(columns=headers)
        for h in headers:
            self._tree.heading(h, text=h)
            w = 300 if h == "Ruta Archivo" else (180 if h in ("Descripción","Vehículo","Info Malla") else 100)
            self._tree.column(h, width=w, minwidth=60)

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
        limit = 100
        try:
            if q:
                where  = self.WHERE[tab]
                n_params = where.count("?")
                like   = f"%{q}%"
                params = (limit,) + (like,) * n_params
                sql    = sql_tpl.format(where=where)
            else:
                params = (limit,)
                sql    = sql_tpl.format(where="")
            rows = db_query(sql, params)
        except Exception as e:
            self.after(0, lambda: self._lbl_count.configure(
                text=f"Error BD: {str(e)[:60]}", text_color=PAL["red"]))
            return
        self.after(0, self._fill, rows, fields, headers)

    def _fill(self, rows, fields, headers):
        for i in self._tree.get_children(): self._tree.delete(i)
        for r in rows:
            vals = [str(r.get(f,"") or "—") for f in fields]
            self._tree.insert("","end", values=vals)
        n = len(rows)
        self._lbl_count.configure(
            text=f"{n} resultado{'s' if n!=1 else ''}  (máx. 100)",
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


# ══════════════════════════════════════════════════════════════════════════════
#  PESTAÑA — SCANNER DE ÓRDENES
# ══════════════════════════════════════════════════════════════════════════════
_CS_COMERCIAL = (
    "DRIVER={ODBC Driver 17 for SQL Server};"
    "SERVER=192.168.2.23;DATABASE=Comercial;"
    "UID=Consulta;PWD=@GPgl4$$2021;"
    "TrustServerCertificate=yes;Connection Timeout=10;"
)
_CS_SAP = (
    "DRIVER={ODBC Driver 17 for SQL Server};"
    "SERVER=agpcolsap.database.windows.net,1433;DATABASE=DB_COL_SAP;"
    "UID=Viewer;PWD=AgpconsCol2023;"
    "Encrypt=yes;TrustServerCertificate=no;Connection Timeout=15;"
)

def _conectar_comercial():
    for drv in ["ODBC Driver 18 for SQL Server", "ODBC Driver 17 for SQL Server"]:
        try:
            return pyodbc.connect(_CS_COMERCIAL.replace("ODBC Driver 17 for SQL Server", drv))
        except Exception:
            continue
    raise RuntimeError("No se pudo conectar a Comercial (192.168.2.23)")

def _conectar_sap():
    for drv in ["ODBC Driver 18 for SQL Server", "ODBC Driver 17 for SQL Server"]:
        try:
            return pyodbc.connect(_CS_SAP.replace("ODBC Driver 17 for SQL Server", drv))
        except Exception:
            continue
    raise RuntimeError("No se pudo conectar a SAP Azure")


class TabScanner(ctk.CTkFrame):

    def __init__(self, parent, **kw):
        super().__init__(parent, fg_color=PAL["bg"], **kw)
        self._build()

    def _build(self):
        self.columnconfigure(0, weight=1)

        # ── Barra de escaneo ─────────────────────────────────────────────────
        top = ctk.CTkFrame(self, fg_color=PAL["card"], corner_radius=0)
        top.grid(row=0, column=0, sticky="ew")
        top.columnconfigure(0, weight=1)

        search_row = ctk.CTkFrame(top, fg_color="transparent")
        search_row.pack(fill="x", padx=30, pady=18)
        search_row.columnconfigure(0, weight=1)

        self._entry = ctk.CTkEntry(
            search_row,
            placeholder_text="Escanea o escribe el número de orden...",
            height=60, font=FONT(20),
            fg_color=PAL["card2"], border_color=PAL["accent"],
            border_width=2, text_color=PAL["txt"], corner_radius=10,
        )
        self._entry.grid(row=0, column=0, sticky="ew", padx=(0, 12))
        self._entry.bind("<Return>",   lambda _: self._buscar())
        self._entry.bind("<KP_Enter>", lambda _: self._buscar())
        self._entry.focus_set()

        ctk.CTkButton(
            search_row, text="BUSCAR", width=130, height=60,
            font=FONT(15, "bold"), corner_radius=10,
            fg_color=PAL["accent2"], hover_color=PAL["accent"],
            command=self._buscar,
        ).grid(row=0, column=1)

        self._prog = ctk.CTkProgressBar(top, mode="indeterminate",
                                         height=4, progress_color=PAL["accent"])
        self._prog.pack(fill="x", padx=0, pady=0)
        self._prog.set(0)

        # ── Zona resultado ───────────────────────────────────────────────────
        self._zona = ctk.CTkFrame(self, fg_color="transparent")
        self._zona.grid(row=1, column=0, sticky="nsew", padx=20, pady=20)
        self._zona.columnconfigure((0, 1), weight=1)
        self._zona.rowconfigure(1, weight=1)
        self.rowconfigure(1, weight=1)

        # — Fila superior: ORDEN  |  ZFER ——————————————————————————————————
        self._c_orden = self._chip(self._zona, "ORDEN", "—",
                                   PAL["accent"], "#0a1a35", col=0)
        self._c_zfer  = self._chip(self._zona, "ZFER",  "—",
                                   PAL["green"],  "#0a2010", col=1)

        # — Fila inferior: VITRO  |  MALLAS ————————————————————————————————
        # Vitro
        box_v = ctk.CTkFrame(self._zona, fg_color=PAL["card"],
                              corner_radius=16, border_width=2,
                              border_color=PAL["accent2"])
        box_v.grid(row=1, column=0, sticky="nsew", padx=(0, 10), pady=0)
        box_v.columnconfigure(0, weight=1)

        ctk.CTkLabel(box_v, text="VITRO",
                     font=FONT(11, "bold"), text_color=PAL["accent"]
                     ).pack(anchor="w", padx=24, pady=(20, 4))
        ctk.CTkFrame(box_v, fg_color=PAL["accent2"], height=2
                     ).pack(fill="x", padx=24, pady=(0, 14))
        self._lbl_vitro = ctk.CTkLabel(
            box_v, text="—", font=FONT(22, "bold"),
            text_color=PAL["txt"], wraplength=480, justify="left",
        )
        self._lbl_vitro.pack(anchor="w", padx=24)
        self._lbl_vitro2 = ctk.CTkLabel(
            box_v, text="", font=FONT(16),
            text_color=PAL["txt_mid"], wraplength=480, justify="left",
        )
        self._lbl_vitro2.pack(anchor="w", padx=24, pady=(6, 20))

        # Mallas
        box_m = ctk.CTkFrame(self._zona, fg_color=PAL["card"],
                              corner_radius=16, border_width=2,
                              border_color="#5b21b6")
        box_m.grid(row=1, column=1, sticky="nsew", padx=(10, 0), pady=0)
        box_m.columnconfigure(0, weight=1)
        box_m.rowconfigure(1, weight=1)

        ctk.CTkLabel(box_m, text="MALLAS",
                     font=FONT(11, "bold"), text_color=PAL["purple"]
                     ).pack(anchor="w", padx=24, pady=(20, 4))
        ctk.CTkFrame(box_m, fg_color=PAL["purple"], height=2
                     ).pack(fill="x", padx=24, pady=(0, 10))

        self._mallas_box = ctk.CTkScrollableFrame(
            box_m, fg_color="transparent", corner_radius=0)
        self._mallas_box.pack(fill="both", expand=True, padx=16, pady=(0, 16))
        self._mallas_box.columnconfigure(0, weight=1)
        self._mallas_labels = []

        # ── Historial (compacto al fondo) ────────────────────────────────────
        import tkinter.ttk as _ttk
        card_hist = ctk.CTkFrame(self, fg_color=PAL["card2"],
                                  corner_radius=0)
        card_hist.grid(row=2, column=0, sticky="ew")

        import tkinter as _tk
        _style = _ttk.Style()
        _style.configure("Hist.Treeview",
                         background=PAL["card2"], foreground=PAL["txt_mid"],
                         fieldbackground=PAL["card2"], rowheight=24,
                         font=("Segoe UI", 9))
        _style.configure("Hist.Treeview.Heading",
                         background=PAL["card2"], foreground=PAL["accent"],
                         font=("Segoe UI", 8, "bold"), relief="flat")
        _style.map("Hist.Treeview",
                   background=[("selected", PAL["accent2"])],
                   foreground=[("selected", "white")])

        frm_h = ctk.CTkFrame(card_hist, fg_color="transparent")
        frm_h.pack(fill="x", padx=10, pady=6)
        self._tree_hist = _ttk.Treeview(
            frm_h, style="Hist.Treeview",
            columns=("hora","orden","zfer","vitro","mallas"),
            show="headings", height=3,
        )
        for col, w, lbl in [("hora",65,"Hora"),("orden",95,"Orden"),
                             ("zfer",95,"ZFER"),("vitro",260,"Vitro"),
                             ("mallas",340,"Mallas")]:
            self._tree_hist.heading(col, text=lbl)
            self._tree_hist.column(col, width=w, minwidth=40)
        self._tree_hist.pack(side="left", fill="x", expand=True)
        sb_h = _ttk.Scrollbar(frm_h, orient="vertical",
                               command=self._tree_hist.yview)
        sb_h.pack(side="right", fill="y")
        self._tree_hist.configure(yscrollcommand=sb_h.set)
        self._tree_hist.bind("<Double-1>", self._hist_click)

        self._zona.grid_remove()

    def _chip(self, parent, label, value, color, bg, col):
        frm = ctk.CTkFrame(parent, fg_color=bg, corner_radius=14,
                            border_width=2, border_color=color)
        frm.grid(row=0, column=col, sticky="ew",
                 padx=(0,10) if col==0 else (10,0), pady=(0,16))
        ctk.CTkLabel(frm, text=label, font=FONT(10, "bold"),
                     text_color=color).pack(anchor="w", padx=20, pady=(14,0))
        lbl = ctk.CTkLabel(frm, text=value, font=FONT(34, "bold"),
                           text_color=color)
        lbl.pack(anchor="w", padx=20, pady=(0,14))
        return lbl

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

            mallas, t1s, t2s = [], [], []
            for comp, t1, t2 in filas:
                if comp and str(comp).strip():
                    mallas.append(str(comp).strip())
                if t1 and str(t1).strip() and str(t1).strip() not in t1s:
                    t1s.append(str(t1).strip())
                if t2 and str(t2).strip() and str(t2).strip() not in t2s:
                    t2s.append(str(t2).strip())

            self.after(0, self._mostrar_resultado, orden, zfer,
                       " / ".join(t1s) or "—", " / ".join(t2s), mallas)
        except Exception as e:
            self.after(0, self._mostrar_error, str(e)[:80])

    def _mostrar_resultado(self, orden, zfer, vitro1, vitro2, mallas):
        self._prog.stop(); self._prog.set(0)
        self._entry.configure(state="normal")

        self._c_orden.configure(text=str(orden))
        self._c_zfer.configure(text=str(zfer))
        self._lbl_vitro.configure(text=vitro1)
        self._lbl_vitro2.configure(text=vitro2)

        for lbl in self._mallas_labels:
            lbl.destroy()
        self._mallas_labels.clear()
        for i, m in enumerate(mallas):
            bg = PAL["card"] if i % 2 == 0 else PAL["card2"]
            lbl = ctk.CTkLabel(
                self._mallas_box, text=f"  {m}",
                font=FONT(15, "bold"), text_color=PAL["txt"],
                fg_color=bg, corner_radius=6, anchor="w",
            )
            lbl.grid(row=i, column=0, sticky="ew", pady=2)
            self._mallas_labels.append(lbl)

        hora = time.strftime("%H:%M:%S")
        mallas_str = " | ".join(mallas[:4]) + (" ..." if len(mallas) > 4 else "")
        self._tree_hist.insert("", 0, values=(hora, orden, zfer, vitro1[:45], mallas_str[:60]))
        for h in self._tree_hist.get_children()[50:]:
            self._tree_hist.delete(h)

        self._zona.grid()
        self._entry.delete(0, "end")
        self._entry.focus_set()

    def _mostrar_error(self, msg):
        self._prog.stop(); self._prog.set(0)
        self._entry.configure(state="normal",
                               border_color=PAL["red"])
        self.after(2000, lambda: self._entry.configure(border_color=PAL["accent"]))
        self._entry.select_range(0, "end")

        for lbl in self._mallas_labels:
            lbl.destroy()
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
        self._lbl_vitro.configure(text="No encontrado")
        self._lbl_vitro2.configure(text="")
        self._zona.grid()

    def _hist_click(self, _):
        sel = self._tree_hist.selection()
        if not sel: return
        vals = self._tree_hist.item(sel[0], "values")
        if vals:
            self._entry.delete(0, "end")
            self._entry.insert(0, vals[1])
            self._buscar()


# ══════════════════════════════════════════════════════════════════════════════
#  APP PRINCIPAL
# ══════════════════════════════════════════════════════════════════════════════
class AGPApp(ctk.CTk):
    PAGES = [
        ("Crear Arte",    "🎨", TabArte),
        ("Consultar BD",  "🔍", TabBD),
        ("Scanner",       "📷", TabScanner),
    ]

    def __init__(self):
        super().__init__()
        self.title("AGP Glass — Suite")
        self.geometry("1280x820")
        self.minsize(1000, 680)
        self._active = None
        self._frames = {}
        self._build()

    def _build(self):
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        # ── SIDEBAR ───────────────────────────────────────────────────────────
        sidebar = ctk.CTkFrame(self, width=220, fg_color=PAL["sidebar"],
                               corner_radius=0)
        sidebar.grid(row=0, column=0, sticky="nsew")
        sidebar.grid_propagate(False)

        # Logo
        logo_frame = ctk.CTkFrame(sidebar, fg_color="transparent")
        logo_frame.pack(fill="x", padx=16, pady=(24,8))
        ctk.CTkLabel(logo_frame, text="AGP", font=FONT(11, "bold"),
                     text_color=PAL["txt_dim"]).pack(anchor="w")
        ctk.CTkLabel(logo_frame, text="Glass Suite", font=FONT(20, "bold"),
                     text_color=PAL["accent"]).pack(anchor="w")
        ctk.CTkFrame(sidebar, fg_color=PAL["border"], height=1
                     ).pack(fill="x", padx=12, pady=10)

        # Nav buttons
        self._nav_btns = {}
        for name, icon, cls in self.PAGES:
            b = SideBtn(sidebar, text=name, icon=icon,
                        command=lambda n=name: self._show(n))
            b.pack(fill="x", padx=8, pady=2)
            self._nav_btns[name] = b

        # Footer
        ctk.CTkFrame(sidebar, fg_color=PAL["border"], height=1
                     ).pack(fill="x", padx=12, pady=10, side="bottom")
        ctk.CTkLabel(sidebar, text="AGP Group © 2025",
                     font=FONT(9), text_color=PAL["txt_dim"]
                     ).pack(side="bottom", pady=8)

        # Acelerador de teclas
        ctk.CTkLabel(sidebar, text="Alt+1 Arte  ·  Alt+2 BD  ·  Alt+3 Scanner",
                     font=FONT(9), text_color=PAL["txt_dim"]
                     ).pack(side="bottom", pady=2)

        # ── CONTENIDO ─────────────────────────────────────────────────────────
        self._content = ctk.CTkScrollableFrame(self, fg_color=PAL["bg"],
                                                corner_radius=0)
        self._content.grid(row=0, column=1, sticky="nsew", padx=0, pady=0)
        self._content.columnconfigure(0, weight=1)

        # Header top strip
        ctk.CTkFrame(self, fg_color=PAL["accent"], height=3
                     ).grid(row=0, column=0, columnspan=2, sticky="new")

        # Instanciar páginas
        for name, icon, cls in self.PAGES:
            f = cls(self._content)
            f.grid(row=0, column=0, sticky="nsew")
            self._frames[name] = f

        self.bind("<Alt-Key-1>", lambda _: self._show("Crear Arte"))
        self.bind("<Alt-Key-2>", lambda _: self._show("Consultar BD"))
        self.bind("<Alt-Key-3>", lambda _: self._show("Scanner"))

        self._show("Crear Arte")

    def _show(self, name):
        if self._active:
            self._frames[self._active].grid_remove()
            self._nav_btns[self._active].set_active(False)
        self._frames[name].grid()
        self._nav_btns[name].set_active(True)
        self._active = name


# ══════════════════════════════════════════════════════════════════════════════
if __name__ == "__main__":
    AGPApp().mainloop()
