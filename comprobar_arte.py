#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
AGP GROUP — Comprobar Arte
Ejecutar:  py comprobar_arte.py

Flujo:
  1. Indica la ruta de búsqueda (vehiculo / modelo / version o cualquier nivel)
     y el plano DWG original.
  2. COMPROBAR busca recursivamente en carpetas ARTES archivos cuyo nombre
     contenga el mismo código base que el plano.
  3. Doble clic en fila verde → AutoCAD abre el arte y superpone el plano
     exactamente en 0,0 para verificar visualmente si coinciden.
"""
import os
import sys
import time
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

try:
    import win32com.client
    import pythoncom
except ImportError:
    tk.Tk().withdraw()
    messagebox.showerror("Error", "Falta pywin32.\nEjecuta:  pip install pywin32")
    sys.exit(1)

# ─── paleta ───────────────────────────────────────────────────────────────────
C = {
    "bg":       "#1E2A38",
    "accent":   "#2E86AB",
    "panel":    "#263545",
    "btn_warn": "#E67E22",
    "txt":      "#ECF0F1",
    "txt_dim":  "#95A5A6",
    "entry_bg": "#1A2533",
    "entry_fg": "#ECF0F1",
    "log_bg":   "#0D1B2A",
    "log_ok":   "#2ECC71",
    "log_warn": "#F39C12",
    "log_err":  "#E74C3C",
}
FT  = ("Segoe UI", 14, "bold")
FL  = ("Segoe UI", 10)
FB  = ("Segoe UI", 10, "bold")
FLG = ("Consolas",  9)
FS  = ("Segoe UI",  8)


# ─── búsqueda ─────────────────────────────────────────────────────────────────

def _codigo_base(ruta_archivo: str) -> str:
    """
    '1708 008 030 A_PLANO.dwg' → '1708 008 030 A'
    '1708 008 030 A_ARTE.dwg'  → '1708 008 030 A'
    """
    base = os.path.splitext(os.path.basename(ruta_archivo))[0].strip()
    return base.rsplit("_", 1)[0].strip() if "_" in base else base


def buscar_artes(ruta: str, codigo: str) -> list:
    """Recorre ruta buscando carpetas ARTES; marca coincidencias con el código."""
    resultados = []
    codigo_cmp = codigo.upper().strip()
    for raiz, dirs, archivos in os.walk(ruta):
        dirs[:] = [d for d in dirs if not d.startswith(".")]
        if os.path.basename(raiz).upper() != "ARTES":
            continue
        for archivo in sorted(archivos):
            if os.path.splitext(archivo)[1].lower() not in (".dwg", ".3dm"):
                continue
            nombre_cmp = os.path.splitext(archivo)[0].upper()
            coincide   = bool(codigo_cmp) and (codigo_cmp in nombre_cmp)
            rel        = os.path.relpath(raiz, ruta)
            resultados.append({
                "version":       rel,
                "archivo":       archivo,
                "ruta_completa": os.path.join(raiz, archivo),
                "coincide":      coincide,
            })
    resultados.sort(key=lambda x: (not x["coincide"], x["version"], x["archivo"]))
    return resultados


# ─── overlay AutoCAD ──────────────────────────────────────────────────────────

def overlay_en_autocad(ruta_arte: str, ruta_plano: str, log_fn=None):
    """
    Abre el arte DWG en AutoCAD y adjunta el plano como XREF en 0,0,0.
    Usa la API COM directa (AttachExternalReference) que es confiable con
    rutas que tienen espacios. Si falla, intenta con XATTACH por SendCommand.
    """
    if log_fn is None:
        log_fn = print

    pythoncom.CoInitialize()
    try:
        try:
            acad = win32com.client.GetActiveObject("AutoCAD.Application")
        except Exception:
            raise RuntimeError(
                "AutoCAD no está abierto.\n"
                "Abre AutoCAD primero y vuelve a intentarlo."
            )

        ruta_arte  = os.path.abspath(ruta_arte)
        ruta_plano = os.path.abspath(ruta_plano)

        log_fn(f"  Abriendo arte:  {os.path.basename(ruta_arte)}")
        doc = acad.Documents.Open(ruta_arte, False, False)
        time.sleep(2.5)

        log_fn(f"  Superponiendo:  {os.path.basename(ruta_plano)}")

        # Intentar con API COM (más confiable con rutas con espacios)
        try:
            pt = win32com.client.VARIANT(
                pythoncom.VT_ARRAY | pythoncom.VT_R8, (0.0, 0.0, 0.0)
            )
            doc.ModelSpace.AttachExternalReference(
                ruta_plano,           # ruta del plano
                "PLANO_VERIFICACION", # nombre de la referencia
                pt,                   # inserción en 0,0,0
                1.0, 1.0, 1.0,        # escala X, Y, Z
                0.0,                  # ángulo de rotación
                False,                # False = Attach (no Overlay)
                None,                 # sin contraseña
            )
            log_fn("  XREF adjuntado por API COM.")
        except Exception as e_api:
            # Fallback: XATTACH por línea de comandos
            log_fn(f"  (API COM: {e_api}) — usando XATTACH por comando...")
            doc.SendCommand(
                f'-XATTACH "{ruta_plano}"\n'
                f'PLANO_VERIFICACION\n'
                f'\n'
                f'0,0,0\n'
                f'1\n'
                f'0\n'
            )
            time.sleep(2.0)

        # Zoom para ver la superposición completa
        doc.SendCommand("ZOOM E \n")
        time.sleep(0.8)
        doc.SendCommand("REGEN \n")
        time.sleep(0.4)

        log_fn("  ────────────────────────────────────────────")
        log_fn("  Superposición lista en AutoCAD.")
        log_fn("  ✔ Si el perímetro del plano coincide con el arte → ARTE CORRECTO")
        log_fn("  ✘ Si no coincide → el arte necesita corrección.")
        log_fn("  ────────────────────────────────────────────")

    finally:
        pythoncom.CoUninitialize()


# ─── App ──────────────────────────────────────────────────────────────────────

class ComprobarArteApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("AGP GROUP — Comprobar Arte")
        self.configure(bg=C["bg"])
        self.resizable(True, True)
        self.minsize(760, 600)
        self._ruta_busq  = tk.StringVar()
        self._dwg_plano  = tk.StringVar()
        self._resultados: list = []
        self._build_ui()
        self._centrar(840, 680)

    # ── UI ────────────────────────────────────────────────────────────────────

    def _build_ui(self):
        hdr = tk.Frame(self, bg=C["accent"], pady=12)
        hdr.pack(fill="x")
        tk.Label(hdr, text="AGP GROUP  —  Comprobar Arte",
                 font=FT, bg=C["accent"], fg="white").pack()

        form = tk.Frame(self, bg=C["bg"], padx=24, pady=16)
        form.pack(fill="x")
        form.columnconfigure(1, weight=1)

        def _fila(row, label, var, hint, explorar_fn, top_pad=4):
            tk.Label(form, text=label, font=FL, bg=C["bg"], fg=C["txt"],
                     anchor="w", width=15
                     ).grid(row=row, column=0, sticky="w",
                            pady=(top_pad, 4))
            tk.Entry(form, textvariable=var, bg=C["entry_bg"],
                     fg=C["entry_fg"], insertbackground=C["entry_fg"],
                     relief="flat", font=FL
                     ).grid(row=row, column=1, sticky="ew",
                            padx=(0, 6), pady=(top_pad, 4))
            tk.Button(form, text="Explorar…", bg=C["accent"], fg="white",
                      relief="flat", font=FL, cursor="hand2",
                      command=explorar_fn
                      ).grid(row=row, column=2, pady=(top_pad, 4))
            tk.Label(form, text=hint, font=FS, bg=C["bg"],
                     fg=C["txt_dim"]
                     ).grid(row=row + 1, column=1, columnspan=2, sticky="w")

        _fila(0, "Ruta búsqueda:", self._ruta_busq,
              "Vehiculo / Modelo / Version — cualquier nivel",
              lambda: self._pick_dir(self._ruta_busq))
        _fila(2, "Plano DWG:", self._dwg_plano,
              "DWG original del plano (para superponer al arte)",
              self._pick_dwg, top_pad=12)

        # Botón
        btns = tk.Frame(self, bg=C["bg"], pady=8)
        btns.pack(fill="x", padx=24)
        self._btn = tk.Button(
            btns, text="  COMPROBAR ARTE  ", command=self._comprobar,
            bg=C["btn_warn"], fg="white", relief="flat", font=FB,
            activebackground=C["btn_warn"], activeforeground="white",
            padx=16, pady=9, cursor="hand2")
        self._btn.pack(side="left")

        # Tabla
        lf = tk.LabelFrame(self, text="Artes encontrados",
                           bg=C["bg"], fg=C["txt_dim"], font=FL,
                           padx=8, pady=6)
        lf.pack(fill="x", padx=24, pady=(6, 0))
        self._tree = ttk.Treeview(lf,
                                  columns=("estado", "ruta", "archivo"),
                                  show="headings", height=5)
        self._tree.heading("estado",  text="Estado")
        self._tree.heading("ruta",    text="Ruta relativa")
        self._tree.heading("archivo", text="Archivo")
        self._tree.column("estado",  width=120, anchor="center", stretch=False)
        self._tree.column("ruta",    width=340)
        self._tree.column("archivo", width=270)
        self._tree.tag_configure("match",
                                 background="#1A3A1A", foreground=C["log_ok"])
        self._tree.tag_configure("other",
                                 background=C["panel"],  foreground=C["txt_dim"])
        self._tree.pack(side="left", fill="x", expand=True)
        sb = ttk.Scrollbar(lf, command=self._tree.yview)
        sb.pack(side="right", fill="y")
        self._tree.configure(yscrollcommand=sb.set)
        self._tree.bind("<Double-1>", self._abrir_overlay)

        tk.Label(self,
                 text="Doble clic en fila verde → abre AutoCAD y superpone plano sobre arte",
                 font=FS, bg=C["bg"], fg=C["txt_dim"]
                 ).pack(anchor="w", padx=28, pady=(2, 0))

        # Log
        lf_log = tk.LabelFrame(self, text="Log",
                               bg=C["bg"], fg=C["txt_dim"], font=FL,
                               padx=8, pady=6)
        lf_log.pack(fill="both", expand=True, padx=24, pady=(6, 16))
        self._log_w = tk.Text(lf_log, bg=C["log_bg"], fg=C["txt"],
                              font=FLG, relief="flat", state="disabled",
                              wrap="word")
        self._log_w.pack(side="left", fill="both", expand=True)
        for tag, color in [("ok",   C["log_ok"]), ("warn", C["log_warn"]),
                           ("err",  C["log_err"]), ("dim",  C["txt_dim"])]:
            self._log_w.tag_config(tag, foreground=color)
        sb2 = ttk.Scrollbar(lf_log, command=self._log_w.yview)
        sb2.pack(side="right", fill="y")
        self._log_w.configure(yscrollcommand=sb2.set)

    # ── helpers ───────────────────────────────────────────────────────────────

    def _centrar(self, w, h):
        x = (self.winfo_screenwidth()  - w) // 2
        y = (self.winfo_screenheight() - h) // 2
        self.geometry(f"{w}x{h}+{x}+{y}")

    def _log(self, msg: str, tag: str = ""):
        self._log_w.configure(state="normal")
        self._log_w.insert("end", f"{time.strftime('%H:%M:%S')}  {msg}\n", tag or "")
        self._log_w.see("end")
        self._log_w.configure(state="disabled")

    def _busy(self, activo: bool):
        self._btn.configure(state="disabled" if activo else "normal")
        self.update_idletasks()

    def _pick_dir(self, var: tk.StringVar):
        ruta = filedialog.askdirectory()
        if ruta:
            var.set(ruta.replace("/", "\\"))

    def _pick_dwg(self):
        inicial = self._ruta_busq.get().strip() or "/"
        ruta = filedialog.askopenfilename(
            title="Seleccionar plano DWG", initialdir=inicial,
            filetypes=[("AutoCAD DWG", "*.dwg"), ("Todos", "*.*")])
        if ruta:
            self._dwg_plano.set(ruta.replace("/", "\\"))

    # ── comprobar ─────────────────────────────────────────────────────────────

    def _comprobar(self):
        ruta = self._ruta_busq.get().strip()
        if not ruta:
            messagebox.showwarning("Campo requerido",
                                   "Indica la carpeta de búsqueda.")
            return
        if not os.path.isdir(ruta):
            messagebox.showerror("Ruta no encontrada",
                                 f"No existe o no es accesible:\n{ruta}")
            return
        self._busy(True)
        threading.Thread(target=self._t_comprobar,
                         args=(ruta, self._dwg_plano.get().strip()),
                         daemon=True).start()

    def _t_comprobar(self, ruta: str, dwg_plano: str):
        self._log("─" * 58)
        self._log(f"Buscando en: {ruta}", "dim")

        codigo = _codigo_base(dwg_plano) if dwg_plano else ""
        if codigo:
            self._log(f'Código del plano: "{codigo}"', "dim")
        else:
            self._log("Sin plano — se listan todos los archivos en ARTES.", "warn")

        resultados = buscar_artes(ruta, codigo)
        self._resultados = resultados
        self.after(0, self._poblar_tabla, resultados)

        coincidencias = [r for r in resultados if r["coincide"]]
        otros         = [r for r in resultados if not r["coincide"]]

        if not resultados:
            self._log("No se encontraron carpetas ARTES con DWG/3DM.", "warn")
        else:
            if coincidencias:
                self._log(
                    f"✔  {len(coincidencias)} coincidencia(s) para '{codigo}':", "ok")
                for r in coincidencias:
                    self._log(f"   {r['version']}  →  {r['archivo']}", "ok")
                self._log(
                    "Doble clic en fila verde para abrir en AutoCAD y verificar.", "ok")
            else:
                self._log(f'Sin coincidencias para "{codigo}".', "warn")
            if otros:
                self._log(
                    f"   ({len(otros)} archivo(s) más en ARTES sin coincidencia)", "dim")

        self._busy(False)

    def _poblar_tabla(self, resultados: list):
        for item in self._tree.get_children():
            self._tree.delete(item)
        for r in resultados:
            estado = "✔  COINCIDE" if r["coincide"] else "—"
            self._tree.insert("", "end",
                              values=(estado, r["version"], r["archivo"]),
                              tags=("match" if r["coincide"] else "other",))

    # ── overlay ───────────────────────────────────────────────────────────────

    def _abrir_overlay(self, _event):
        sel = self._tree.selection()
        if not sel:
            return
        idx = self._tree.index(sel[0])
        if idx >= len(self._resultados):
            return
        r = self._resultados[idx]

        if not r["ruta_completa"].lower().endswith(".dwg"):
            messagebox.showinfo("Solo DWG",
                f"La superposición solo funciona con DWG.\n{r['archivo']}")
            return

        dwg_plano = self._dwg_plano.get().strip()

        # Si no es una ruta de archivo válida, buscar el DWG del plano
        # automáticamente en la ruta de búsqueda usando el texto como código
        if not os.path.isfile(dwg_plano):
            codigo_buscar = _codigo_base(dwg_plano) if dwg_plano else _codigo_base(r["archivo"])
            dwg_plano = self._buscar_plano_dwg(codigo_buscar)
            if not dwg_plano:
                messagebox.showwarning(
                    "Plano no encontrado",
                    f'No se encontró un DWG de plano con código "{codigo_buscar}".\n'
                    f'Usa Explorar para seleccionar el DWG original del plano.'
                )
                return

        self._log(f"Iniciando overlay: {r['archivo']}", "ok")
        self._busy(True)
        threading.Thread(
            target=self._t_overlay,
            args=(r["ruta_completa"], dwg_plano),
            daemon=True).start()

    def _buscar_plano_dwg(self, codigo: str) -> str | None:
        """
        Busca recursivamente en la ruta de búsqueda un DWG cuyo nombre
        contenga el código dado pero que NO esté en una carpeta ARTES
        (es el plano original, no el arte).
        Devuelve la ruta completa o None si no encuentra.
        """
        ruta_busq  = self._ruta_busq.get().strip()
        codigo_cmp = codigo.upper().strip()
        for raiz, dirs, archivos in os.walk(ruta_busq):
            dirs[:] = [d for d in dirs if not d.startswith(".")]
            # Saltar carpetas ARTES — ahí están los artes, no los planos
            if os.path.basename(raiz).upper() == "ARTES":
                dirs[:] = []
                continue
            for archivo in archivos:
                if not archivo.lower().endswith(".dwg"):
                    continue
                nombre_cmp = os.path.splitext(archivo)[0].upper()
                if codigo_cmp in nombre_cmp:
                    return os.path.join(raiz, archivo)
        return None

    def _t_overlay(self, ruta_arte: str, ruta_plano: str):
        try:
            overlay_en_autocad(ruta_arte, ruta_plano,
                               log_fn=lambda m: self._log(m, "dim"))
        except RuntimeError as e:
            self._log(str(e), "err")
        except Exception as e:
            self._log(f"ERROR: {e}", "err")
        finally:
            self._busy(False)


# ─────────────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    ComprobarArteApp().mainloop()
