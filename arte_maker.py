#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
AGP GROUP — Creador de Artes
Ejecutar:  py arte_maker.py

El dibujante entrega:
  1. Ruta completa de la version, ej:
     \\\\192.168.2.37\\ingenieria\\PRODUCCION\\AGP PLANOS TECNICOS\\ACURA\\ACURA MDX 4D U 2014--398\\V-06 SVM CL6
  2. El archivo DWG del plano
  3. Si tiene o no degrade

La app entra sola a  [ruta_version]/ARTES/  para verificar y guardar.
"""
import os
import sys
import time
import tempfile
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

try:
    from verificacion import extraer_sufijo, buscar_en_version, ruta_artes_de_version
    from autocad_ops import AutoCADMotor
    from rhino_ops import generar_y_ejecutar
except ImportError as e:
    import tkinter as _tk
    _tk.Tk().withdraw()
    import tkinter.messagebox as _mb
    _mb.showerror("Error de importacion", str(e))
    sys.exit(1)

# ─── paleta ───────────────────────────────────────────────────────────────────
C = {
    "bg":       "#1E2A38",
    "accent":   "#2E86AB",
    "panel":    "#263545",
    "btn_ok":   "#27AE60",
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
FLG = ("Consolas", 9)
FS  = ("Segoe UI", 8)


# ─────────────────────────────────────────────────────────────────────────────

class ArteMakerApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("AGP GROUP — Creador de Artes")
        self.configure(bg=C["bg"])
        self.resizable(True, True)
        self.minsize(700, 560)

        self._ruta_version = tk.StringVar()
        self._dwg_plano    = tk.StringVar()

        self._build_ui()
        self._centrar(760, 620)

    # ── construccion UI ───────────────────────────────────────────────────────

    def _build_ui(self):
        # ── Header ──────────────────────────────────────────────────────────
        hdr = tk.Frame(self, bg=C["accent"], pady=12)
        hdr.pack(fill="x")
        tk.Label(hdr, text="AGP GROUP  —  Creador de Artes",
                 font=FT, bg=C["accent"], fg="white").pack()

        # ── Formulario ──────────────────────────────────────────────────────
        form = tk.Frame(self, bg=C["bg"], padx=24, pady=16)
        form.pack(fill="x")
        form.columnconfigure(1, weight=1)

        # Campo 1: Ruta de la version
        tk.Label(form, text="Ruta version:", font=FL,
                 bg=C["bg"], fg=C["txt"], anchor="w", width=14
                 ).grid(row=0, column=0, sticky="w", pady=4)

        self._e_ruta = tk.Entry(form, textvariable=self._ruta_version,
                                bg=C["entry_bg"], fg=C["entry_fg"],
                                insertbackground=C["entry_fg"],
                                relief="flat", font=FL)
        self._e_ruta.grid(row=0, column=1, sticky="ew", padx=(0, 6), pady=4)

        tk.Button(form, text="Explorar...",
                  bg=C["accent"], fg="white", relief="flat", font=FL,
                  activebackground=C["accent"], activeforeground="white",
                  cursor="hand2", command=self._explorar_version
                  ).grid(row=0, column=2, pady=4)

        tk.Label(form,
                 text="Ej: \\\\servidor\\...\\ACURA MDX 4D U 2014--398\\V-06 SVM CL6",
                 font=FS, bg=C["bg"], fg=C["txt_dim"]
                 ).grid(row=1, column=1, columnspan=2, sticky="w")

        # Campo 2: Plano DWG
        tk.Label(form, text="Plano DWG:", font=FL,
                 bg=C["bg"], fg=C["txt"], anchor="w", width=14
                 ).grid(row=2, column=0, sticky="w", pady=(12, 4))

        self._e_dwg = tk.Entry(form, textvariable=self._dwg_plano,
                               bg=C["entry_bg"], fg=C["entry_fg"],
                               insertbackground=C["entry_fg"],
                               relief="flat", font=FL)
        self._e_dwg.grid(row=2, column=1, sticky="ew", padx=(0, 6), pady=(12, 4))

        tk.Button(form, text="Explorar...",
                  bg=C["accent"], fg="white", relief="flat", font=FL,
                  activebackground=C["accent"], activeforeground="white",
                  cursor="hand2", command=self._explorar_dwg
                  ).grid(row=2, column=2, pady=(12, 4))

        tk.Label(form, text="Plano de AutoCAD con perimetro, banda negra y logo",
                 font=FS, bg=C["bg"], fg=C["txt_dim"]
                 ).grid(row=3, column=1, columnspan=2, sticky="w")

        # ── Botones ──────────────────────────────────────────────────────────
        btns = tk.Frame(self, bg=C["bg"], pady=12)
        btns.pack(fill="x", padx=24)

        self._btn_ver = self._boton(btns, "  VERIFICAR  ",
                                    self._verificar, C["btn_warn"])
        self._btn_ver.pack(side="left", padx=(0, 12))

        self._btn_art = self._boton(btns, "  CREAR ARTE  ",
                                    self._crear_arte, C["btn_ok"])
        self._btn_art.pack(side="left")

        # ── Log ──────────────────────────────────────────────────────────────
        lf = tk.LabelFrame(self, text="Log",
                           bg=C["bg"], fg=C["txt_dim"], font=FL,
                           padx=8, pady=6)
        lf.pack(fill="both", expand=True, padx=24, pady=(0, 16))

        self._log_w = tk.Text(lf, bg=C["log_bg"], fg=C["txt"],
                              font=FLG, relief="flat", state="disabled",
                              wrap="word")
        self._log_w.pack(side="left", fill="both", expand=True)
        self._log_w.tag_config("ok",   foreground=C["log_ok"])
        self._log_w.tag_config("warn", foreground=C["log_warn"])
        self._log_w.tag_config("err",  foreground=C["log_err"])
        self._log_w.tag_config("dim",  foreground=C["txt_dim"])

        sb = ttk.Scrollbar(lf, command=self._log_w.yview)
        sb.pack(side="right", fill="y")
        self._log_w.configure(yscrollcommand=sb.set)

    # ── helpers ───────────────────────────────────────────────────────────────

    @staticmethod
    def _boton(parent, txt, cmd, color):
        return tk.Button(parent, text=txt, command=cmd,
                         bg=color, fg="white", relief="flat", font=FB,
                         activebackground=color, activeforeground="white",
                         padx=16, pady=9, cursor="hand2")

    def _centrar(self, w, h):
        x = (self.winfo_screenwidth()  - w) // 2
        y = (self.winfo_screenheight() - h) // 2
        self.geometry(f"{w}x{h}+{x}+{y}")

    def _log(self, msg: str, tag: str = ""):
        self._log_w.configure(state="normal")
        self._log_w.insert("end",
                           f"{time.strftime('%H:%M:%S')}  {msg}\n",
                           tag or "")
        self._log_w.see("end")
        self._log_w.configure(state="disabled")

    def _busy(self, activo: bool):
        s = "disabled" if activo else "normal"
        self._btn_ver.configure(state=s)
        self._btn_art.configure(state=s)
        self.update_idletasks()

    # ── explorar ──────────────────────────────────────────────────────────────

    def _explorar_version(self):
        """Abre un selector de carpeta para elegir la ruta de la version."""
        ruta = filedialog.askdirectory(title="Seleccionar carpeta de la version")
        if ruta:
            self._ruta_version.set(ruta.replace("/", "\\"))

    def _explorar_dwg(self):
        """Abre un selector de archivo para el plano DWG."""
        inicial = self._ruta_version.get().strip() or "/"
        ruta = filedialog.askopenfilename(
            title="Seleccionar plano DWG",
            initialdir=inicial,
            filetypes=[("AutoCAD DWG", "*.dwg"), ("Todos", "*.*")],
        )
        if ruta:
            self._dwg_plano.set(ruta.replace("/", "\\"))

    # ── verificar ─────────────────────────────────────────────────────────────

    def _verificar(self):
        if not self._validar(solo_version=True):
            return
        self._busy(True)
        threading.Thread(target=self._t_verificar, daemon=True).start()

    def _t_verificar(self):
        ruta_ver = self._ruta_version.get().strip()
        ruta_dwg = self._dwg_plano.get().strip()

        self._log("─" * 54)

        # Detectar si el usuario seleccionó la carpeta ARTES en vez de la version
        if os.path.basename(ruta_ver).upper() == "ARTES":
            self._log(
                "AVISO: seleccionaste la carpeta ARTES en vez de la version.", "warn"
            )
            self._log("       Igual se listaran los archivos que contiene.", "warn")

        self._log(f"Carpeta : {ruta_ver}", "dim")

        sufijo = extraer_sufijo(ruta_dwg) if ruta_dwg else ""
        if sufijo:
            self._log(f'Codigo buscado: "{sufijo}"', "dim")

        archivos = buscar_en_version(ruta_ver, sufijo)

        if not archivos:
            self._log("La carpeta ARTES no existe o esta vacia.", "ok")
        else:
            self._log(f"Archivos en ARTES ({len(archivos)}):", "dim")
            hay_coincidencia = False
            for it in archivos:
                if it.get("coincide"):
                    self._log(f"   [COINCIDE]  {it['archivo']}", "warn")
                    hay_coincidencia = True
                else:
                    self._log(f"   {it['archivo']}", "dim")
            if sufijo and not hay_coincidencia:
                self._log(
                    f'Ninguno coincide con el codigo "{sufijo}".', "ok"
                )

        self._busy(False)

    # ── crear arte ────────────────────────────────────────────────────────────

    def _crear_arte(self):
        if not self._validar():
            return
        self._busy(True)
        threading.Thread(target=self._t_crear_arte, daemon=True).start()

    def _t_crear_arte(self):
        ruta_ver   = self._ruta_version.get().strip()
        ruta_plano = self._dwg_plano.get().strip()

        self._log("=" * 54)
        self._log("Creando arte (degrade se detecta automaticamente en Rhino)", "ok")
        self._log(f"Version : {os.path.basename(ruta_ver)}", "dim")
        self._log(f"Plano   : {os.path.basename(ruta_plano)}", "dim")

        # Verificacion previa (no bloquea)
        sufijo = extraer_sufijo(ruta_plano)
        prev   = buscar_en_version(ruta_ver, sufijo)
        if prev:
            self._log(
                f'AVISO: ya existe arte con codigo "{sufijo}":', "warn"
            )
            for it in prev:
                self._log(f"   {it['archivo']}", "warn")

        # Paso 1 — Extraer layers en AutoCAD y guardar en ARTES
        self._log("--- [1/2] Extraccion en AutoCAD")
        try:
            motor = AutoCADMotor()
        except RuntimeError as e:
            self._log(f"ERROR AutoCAD: {e}", "err")
            self._busy(False)
            return

        nombre_base   = os.path.splitext(os.path.basename(ruta_plano))[0]
        ruta_artes    = ruta_artes_de_version(ruta_ver)
        # DWG limpio va directo a ARTES — el dibujante lo arrastra a Rhino
        ruta_filtrada = os.path.join(ruta_artes, f"{nombre_base}_PLANO.dwg")
        ruta_salida   = os.path.join(ruta_artes, f"{nombre_base}_ARTE.3dm")

        try:
            motor.extraer_layers(
                ruta_plano,
                ruta_filtrada,
                log_fn=lambda m: self._log(m, "dim"),
            )
        except Exception as e:
            self._log(f"ERROR extraccion: {e}", "err")
            motor.quit()
            self._busy(False)
            return

        motor.quit()
        self._log("Extraccion OK.", "ok")
        self._log(f"DWG limpio guardado en: {ruta_filtrada}", "ok")

        # Paso 2 — Generar script Rhino
        self._log("--- [2/2] Generando script Rhino")

        try:
            script_path = generar_y_ejecutar(
                dwg_plano   = ruta_filtrada,
                ruta_salida = ruta_salida,
                log_fn      = lambda m: self._log(m, "dim"),
            )
            self._log("Script generado.", "ok")
            self._log(f"Arte se guardara en: {ruta_salida}", "dim")
            # Mostrar panel de instrucciones Rhino
            self.after(0, lambda: self._mostrar_panel_rhino(script_path, ruta_filtrada))
        except FileNotFoundError as e:
            self._log(f"ERROR: {e}", "err")
        except Exception as e:
            self._log(f"ERROR Rhino: {e}", "err")

        self._log("=" * 54)
        self._busy(False)

    # ── panel Rhino ───────────────────────────────────────────────────────────

    def _mostrar_panel_rhino(self, script_path: str, dwg_plano: str = ""):
        """Muestra una ventana flotante con instrucciones para ejecutar en Rhino."""
        win = tk.Toplevel(self)
        win.title("Ejecutar en Rhino 8")
        win.configure(bg=C["bg"])
        win.resizable(False, False)
        win.grab_set()

        # Centrar sobre la ventana principal
        win.update_idletasks()
        x = self.winfo_x() + (self.winfo_width()  - 560) // 2
        y = self.winfo_y() + (self.winfo_height() - 320) // 2
        win.geometry(f"580x420+{x}+{y}")

        # Header
        tk.Frame(win, bg=C["accent"], height=6).pack(fill="x")
        tk.Label(win, text="Pasos para crear el arte en Rhino 8",
                 font=FB, bg=C["bg"], fg=C["txt"], pady=10).pack()

        def _bloque(titulo, valor, color_btn, accion):
            """Widget de una fila con label, campo readonly y boton."""
            f = tk.Frame(win, bg=C["panel"], padx=12, pady=8)
            f.pack(fill="x", padx=20, pady=(0, 6))
            f.columnconfigure(0, weight=1)
            tk.Label(f, text=titulo, font=FS,
                     bg=C["panel"], fg=C["txt_dim"]).grid(row=0, column=0,
                                                           columnspan=2, sticky="w")
            var = tk.StringVar(value=valor)
            tk.Entry(f, textvariable=var, bg=C["entry_bg"], fg=C["entry_fg"],
                     relief="flat", font=("Consolas", 8),
                     state="readonly").grid(row=1, column=0, sticky="ew",
                                            padx=(0, 6), pady=(2, 0))
            btn = tk.Button(f, text="Abrir", command=accion,
                            bg=color_btn, fg="white", relief="flat",
                            font=FS, padx=10, cursor="hand2")
            btn.grid(row=1, column=1, pady=(2, 0))
            return btn

        import subprocess as sp

        # Paso 1 — importar el DWG limpio
        tk.Label(win, text="PASO 1  —  Arrastra o importa este DWG en Rhino:",
                 font=FL, bg=C["bg"], fg=C["txt"], padx=20, anchor="w"
                 ).pack(fill="x", pady=(6, 0))

        def abrir_dwg():
            sp.Popen(["explorer", "/select,", dwg_plano])

        _bloque("DWG limpio (arrastralo a Rhino):", dwg_plano,
                C["btn_ok"], abrir_dwg)

        # Paso 2 — ejecutar el script
        tk.Label(win, text="PASO 2  —  Ejecuta este script en Rhino (_RunPythonScript):",
                 font=FL, bg=C["bg"], fg=C["txt"], padx=20, anchor="w"
                 ).pack(fill="x", pady=(4, 0))

        def abrir_script():
            sp.Popen(["explorer", "/select,", script_path])

        btn_scr = _bloque("Script Rhino (ruta fija — siempre la misma):",
                          script_path, C["accent"], abrir_script)

        # Boton copiar ruta del script
        def copiar_script():
            self.clipboard_clear()
            self.clipboard_append(script_path)
            btn_scr.configure(text="Copiado!")
            win.after(1500, lambda: btn_scr.configure(text="Abrir"))

        btn_scr.configure(command=copiar_script)

        # Tip toolbar
        tip = (
            "TIP: En Rhino, crea un boton en la barra con este comando:\n"
            f'_RunPythonScript "{script_path}"\n'
            "Asi desde el proximo arte solo importas el DWG y das un click."
        )
        tk.Label(win, text=tip, font=FS, bg=C["bg"], fg=C["txt_dim"],
                 justify="left", padx=20, wraplength=540
                 ).pack(anchor="w", pady=(4, 0))

        tk.Button(win, text="Cerrar", command=win.destroy,
                  bg=C["panel"], fg=C["txt"], relief="flat",
                  font=FL, padx=16, pady=6, cursor="hand2").pack(pady=10)

    # ── validacion ────────────────────────────────────────────────────────────

    def _validar(self, solo_version: bool = False) -> bool:
        ruta_ver = self._ruta_version.get().strip()
        if not ruta_ver:
            messagebox.showwarning(
                "Campo requerido",
                "Pega o explora la ruta completa de la version.\n\n"
                "Ej: \\\\servidor\\...\\ACURA MDX\\V-06 SVM CL6"
            )
            return False
        if not os.path.isdir(ruta_ver):
            messagebox.showerror(
                "Ruta no encontrada",
                f"No se puede acceder a:\n{ruta_ver}\n\n"
                "Verifica que la ruta de red este disponible."
            )
            return False
        if not solo_version:
            ruta_dwg = self._dwg_plano.get().strip()
            if not ruta_dwg:
                messagebox.showwarning("Campo requerido",
                                       "Selecciona el archivo DWG del plano.")
                return False
            if not os.path.isfile(ruta_dwg):
                messagebox.showerror("Archivo no encontrado",
                                     f"No existe:\n{ruta_dwg}")
                return False
        return True


# ─────────────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    ArteMakerApp().mainloop()
