
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

try:
    from autocad_ops import AutoCADMotor
except ImportError as e:
    tk.Tk().withdraw()
    messagebox.showerror("Error de importacion", str(e))
    sys.exit(1)


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
FLG = ("Consolas",  9)
FS  = ("Segoe UI",  8)

SCRIPT_RHINO = os.path.join(os.path.dirname(os.path.abspath(__file__)), "arte_script.py")



def _codigo_base(ruta_archivo: str) -> str:
   
    base = os.path.splitext(os.path.basename(ruta_archivo))[0].strip()
    return base.rsplit("_", 1)[0].strip() if "_" in base else base


def _buscar_artes(ruta: str, codigo: str) -> list:
   
    resultados = []
    codigo_cmp = codigo.upper().strip()
    for raiz, dirs, archivos in os.walk(ruta):
        dirs[:] = [d for d in dirs if not d.startswith(".")]
        if os.path.basename(raiz).upper() == "ARTES":
            for archivo in sorted(archivos):
                if os.path.splitext(archivo)[1].lower() not in (".dwg", ".3dm"):
                    continue
                nombre_cmp = os.path.splitext(archivo)[0].upper()
                coincide   = bool(codigo_cmp) and (codigo_cmp in nombre_cmp)
                rel = os.path.relpath(raiz, ruta)
                resultados.append({
                    "version":       rel,
                    "archivo":       archivo,
                    "ruta_completa": os.path.join(raiz, archivo),
                    "coincide":      coincide,
                })
    resultados.sort(key=lambda x: (not x["coincide"], x["version"], x["archivo"]))
    return resultados


def _ruta_artes(ruta_version: str) -> str:
    """Devuelve (y crea si no existe) la carpeta ARTES dentro de ruta_version."""
    if os.path.basename(ruta_version).upper() == "ARTES":
        return ruta_version
    ruta = os.path.join(ruta_version, "ARTES")
    os.makedirs(ruta, exist_ok=True)
    return ruta




def _overlay_autocad(ruta_arte: str, ruta_plano: str, log_fn=None):
    
    if log_fn is None:
        log_fn = print

    pythoncom.CoInitialize()
    try:
        try:
            acad = win32com.client.GetActiveObject("AutoCAD.Application")
        except Exception:
            raise RuntimeError(
                "AutoCAD no está abierto.\n"
                "Abre AutoCAD primero hpta bobo y vuelve a intentarlo."
            )

        log_fn(f"  Abriendo: {os.path.basename(ruta_arte)}")
        doc = acad.Documents.Open(os.path.abspath(ruta_arte), False, False)
        time.sleep(2)

        log_fn(f"  Superponiendo plano como XREF: {os.path.basename(ruta_plano)}")
        doc.SendCommand(
            f'-XREF A "{os.path.abspath(ruta_plano)}" 0,0,0 1 1 0\n'
        )
        time.sleep(1.5)
        doc.SendCommand("ZOOM E \n")
        time.sleep(0.5)
        log_fn("  Listo. Verifica que el perímetro del plano coincida con el arte.")
    finally:
        pythoncom.CoUninitialize()




class ArteMakerApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("AGP GROUP — Arte Maker")
        self.configure(bg=C["bg"])
        self.resizable(True, True)
        self.minsize(780, 620)

        self._ruta_version = tk.StringVar()
        self._dwg_plano    = tk.StringVar()
        self._resultados: list = []

        self._build_ui()
        self._centrar(860, 700)

    

    def _build_ui(self):
        # Header
        hdr = tk.Frame(self, bg=C["accent"], pady=12)
        hdr.pack(fill="x")
        tk.Label(hdr, text="AGP GROUP  —  Arte Maker",
                 font=FT, bg=C["accent"], fg="white").pack()

        # Formulario
        form = tk.Frame(self, bg=C["bg"], padx=24, pady=16)
        form.pack(fill="x")
        form.columnconfigure(1, weight=1)

        # Ruta versión
        tk.Label(form, text="Ruta versión:", font=FL,
                 bg=C["bg"], fg=C["txt"], anchor="w", width=14
                 ).grid(row=0, column=0, sticky="w", pady=4)
        tk.Entry(form, textvariable=self._ruta_version,
                 bg=C["entry_bg"], fg=C["entry_fg"],
                 insertbackground=C["entry_fg"], relief="flat", font=FL
                 ).grid(row=0, column=1, sticky="ew", padx=(0, 6), pady=4)
        tk.Button(form, text="Explorar…", bg=C["accent"], fg="white",
                  relief="flat", font=FL, cursor="hand2",
                  command=self._explorar_version
                  ).grid(row=0, column=2, pady=4)
        tk.Label(form,
                 text="Carpeta de la versión  (Vehiculo / Modelo / V-XX …)",
                 font=FS, bg=C["bg"], fg=C["txt_dim"]
                 ).grid(row=1, column=1, columnspan=2, sticky="w")

        # Plano DWG
        tk.Label(form, text="Plano DWG:", font=FL,
                 bg=C["bg"], fg=C["txt"], anchor="w", width=14
                 ).grid(row=2, column=0, sticky="w", pady=(12, 4))
        tk.Entry(form, textvariable=self._dwg_plano,
                 bg=C["entry_bg"], fg=C["entry_fg"],
                 insertbackground=C["entry_fg"], relief="flat", font=FL
                 ).grid(row=2, column=1, sticky="ew", padx=(0, 6), pady=(12, 4))
        tk.Button(form, text="Explorar…", bg=C["accent"], fg="white",
                  relief="flat", font=FL, cursor="hand2",
                  command=self._explorar_dwg
                  ).grid(row=2, column=2, pady=(12, 4))
        tk.Label(form, text="DWG original del plano técnico",
                 font=FS, bg=C["bg"], fg=C["txt_dim"]
                 ).grid(row=3, column=1, columnspan=2, sticky="w")

        # Botones principales
        btns = tk.Frame(self, bg=C["bg"], pady=10)
        btns.pack(fill="x", padx=24)

        self._btn_extraer = self._boton(
            btns, "  EXTRAER PLANO  ", self._extraer, C["btn_ok"])
        self._btn_extraer.pack(side="left", padx=(0, 12))

        self._btn_comprobar = self._boton(
            btns, "  COMPROBAR ARTE  ", self._comprobar, C["btn_warn"])
        self._btn_comprobar.pack(side="left")

        # Tabla de resultados (visible tras COMPROBAR)
        lf_res = tk.LabelFrame(self, text="Artes encontrados",
                               bg=C["bg"], fg=C["txt_dim"], font=FL,
                               padx=8, pady=6)
        lf_res.pack(fill="x", padx=24, pady=(6, 0))

        cols = ("estado", "ruta", "archivo")
        self._tree = ttk.Treeview(lf_res, columns=cols,
                                  show="headings", height=5)
        self._tree.heading("estado",  text="Estado")
        self._tree.heading("ruta",    text="Ruta relativa")
        self._tree.heading("archivo", text="Archivo")
        self._tree.column("estado",  width=110, anchor="center", stretch=False)
        self._tree.column("ruta",    width=340)
        self._tree.column("archivo", width=260)
        self._tree.tag_configure("match",
                                 background="#1A3A1A", foreground=C["log_ok"])
        self._tree.tag_configure("other",
                                 background=C["panel"],  foreground=C["txt_dim"])
        self._tree.pack(side="left", fill="both", expand=True)
        sb = ttk.Scrollbar(lf_res, command=self._tree.yview)
        sb.pack(side="right", fill="y")
        self._tree.configure(yscrollcommand=sb.set)
        self._tree.bind("<Double-1>", self._on_doble_click)

        tk.Label(self,
                 text="Doble clic en fila verde → abre AutoCAD con arte + plano superpuestos",
                 font=FS, bg=C["bg"], fg=C["txt_dim"]
                 ).pack(anchor="w", padx=28, pady=(2, 0))

       
        lf_log = tk.LabelFrame(self, text="Log",
                               bg=C["bg"], fg=C["txt_dim"], font=FL,
                               padx=8, pady=6)
        lf_log.pack(fill="both", expand=True, padx=24, pady=(6, 16))

        self._log_w = tk.Text(lf_log, bg=C["log_bg"], fg=C["txt"],
                              font=FLG, relief="flat", state="disabled",
                              wrap="word")
        self._log_w.pack(side="left", fill="both", expand=True)
        for tag, color in [("ok",   C["log_ok"]),
                           ("warn", C["log_warn"]),
                           ("err",  C["log_err"]),
                           ("dim",  C["txt_dim"])]:
            self._log_w.tag_config(tag, foreground=color)
        sb2 = ttk.Scrollbar(lf_log, command=self._log_w.yview)
        sb2.pack(side="right", fill="y")
        self._log_w.configure(yscrollcommand=sb2.set)

    

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
        self._log_w.insert("end", f"{time.strftime('%H:%M:%S')}  {msg}\n", tag or "")
        self._log_w.see("end")
        self._log_w.configure(state="disabled")

    def _busy(self, activo: bool):
        estado = "disabled" if activo else "normal"
        self._btn_extraer.configure(state=estado)
        self._btn_comprobar.configure(state=estado)
        self.update_idletasks()

    def _explorar_version(self):
        ruta = filedialog.askdirectory(title="Carpeta de la versión")
        if ruta:
            self._ruta_version.set(ruta.replace("/", "\\"))

    def _explorar_dwg(self):
        inicial = self._ruta_version.get().strip() or "/"
        ruta = filedialog.askopenfilename(
            title="Seleccionar plano DWG",
            initialdir=inicial,
            filetypes=[("AutoCAD DWG", "*.dwg"), ("Todos", "*.*")],
        )
        if ruta:
            self._dwg_plano.set(ruta.replace("/", "\\"))

    def _validar(self, necesita_dwg=True) -> bool:
        ruta_ver = self._ruta_version.get().strip()
        if not ruta_ver:
            messagebox.showwarning("Campo requerido",
                                   "Indica la ruta de la versión.")
            return False
        if not os.path.isdir(ruta_ver):
            messagebox.showerror("Ruta no encontrada",
                                 f"No existe o no es accesible:\n{ruta_ver}")
            return False
        if necesita_dwg:
            dwg = self._dwg_plano.get().strip()
            if not dwg:
                messagebox.showwarning("Campo requerido",
                                       "Selecciona el archivo DWG del plano.")
                return False
            if not os.path.isfile(dwg):
                messagebox.showerror("Archivo no encontrado",
                                     f"No existe:\n{dwg}")
                return False
        return True

    

    def _extraer(self):
        if not self._validar(necesita_dwg=True):
            return
        self._busy(True)
        threading.Thread(target=self._t_extraer, daemon=True).start()

    def _t_extraer(self):
        ruta_ver  = self._ruta_version.get().strip()
        ruta_plano = self._dwg_plano.get().strip()

        self._log("=" * 56)
        self._log("EXTRAER PLANO — filtrando layers en AutoCAD...", "ok")
        self._log(f"Plano : {os.path.basename(ruta_plano)}", "dim")

        nombre_base   = os.path.splitext(os.path.basename(ruta_plano))[0]
        ruta_artes    = _ruta_artes(ruta_ver)
        ruta_filtrada = os.path.join(ruta_artes, f"{nombre_base}_PLANO.dwg")

        try:
            motor = AutoCADMotor()
        except RuntimeError as e:
            self._log(f"ERROR AutoCAD: {e}", "err")
            self._busy(False)
            return

        try:
            motor.extraer_layers(
                ruta_plano,
                ruta_filtrada,
                log_fn=lambda m: self._log(m, "dim"),
            )
        except Exception as e:
            self._log(f"ERROR extracción: {e}", "err")
            motor.quit()
            self._busy(False)
            return

        motor.quit()
        self._log("Extracción completada.", "ok")
        self._log(f"DWG limpio → {ruta_filtrada}", "ok")
        self._log("─" * 56)
        self._log("SIGUIENTE PASO en Rhino:", "warn")
        self._log("  1. Arrastra el DWG limpio a Rhino", "dim")
        self._log(f"  2. Ejecuta:  _RunPythonScript  →  arte_script.py", "dim")
        self._log(f"     ({SCRIPT_RHINO})", "dim")

        
        import subprocess
        self.after(0, lambda: subprocess.Popen(
            ["explorer", "/select,", ruta_filtrada]))

        self._busy(False)

    

    def _comprobar(self):
        if not self._validar(necesita_dwg=False):
            return
        self._busy(True)
        threading.Thread(target=self._t_comprobar, daemon=True).start()

    def _t_comprobar(self):
        ruta_ver  = self._ruta_version.get().strip()
        dwg_plano = self._dwg_plano.get().strip()

        self._log("=" * 56)
        self._log("COMPROBAR ARTE — buscando artes...", "ok")
        self._log(f"Buscando en: {ruta_ver}", "dim")

        codigo = _codigo_base(dwg_plano) if dwg_plano else ""
        if codigo:
            self._log(f'Código del plano: "{codigo}"', "dim")

        resultados = _buscar_artes(ruta_ver, codigo)
        self._resultados = resultados

        self.after(0, self._poblar_tabla, resultados)

        n_match = sum(1 for r in resultados if r["coincide"])
        if not resultados:
            self._log("No se encontraron archivos en carpetas ARTES.", "warn")
        else:
            self._log(
                f"Total: {len(resultados)} archivo(s)  —  "
                f"coincidencias: {n_match}",
                "ok" if n_match else "warn",
            )
            if n_match:
                self._log(
                    "Doble clic en una fila verde para superponer en AutoCAD.", "ok")

        self._busy(False)

    def _poblar_tabla(self, resultados: list):
        for item in self._tree.get_children():
            self._tree.delete(item)
        for r in resultados:
            estado = "✔  COINCIDE" if r["coincide"] else "—"
            tag    = "match" if r["coincide"] else "other"
            self._tree.insert("", "end",
                              values=(estado, r["version"], r["archivo"]),
                              tags=(tag,))

    def _on_doble_click(self, _event):
        sel = self._tree.selection()
        if not sel:
            return
        idx = self._tree.index(sel[0])
        if idx >= len(self._resultados):
            return
        r = self._resultados[idx]

        dwg_plano = self._dwg_plano.get().strip()
        if not dwg_plano or not os.path.isfile(dwg_plano):
            messagebox.showwarning(
                "Plano requerido",
                "Indica el plano DWG original para poder superponer.")
            return
        if not r["ruta_completa"].lower().endswith(".dwg"):
            messagebox.showinfo(
                "Solo DWG",
                f"La superposición requiere un archivo DWG.\n{r['archivo']}")
            return

        self._log(f"Superponiendo: {r['archivo']}", "ok")
        self._busy(True)
        threading.Thread(
            target=self._t_overlay,
            args=(r["ruta_completa"], dwg_plano),
            daemon=True,
        ).start()

    def _t_overlay(self, ruta_arte: str, ruta_plano: str):
        try:
            _overlay_autocad(
                ruta_arte, ruta_plano,
                log_fn=lambda m: self._log(m, "dim"),
            )
            self._log("Superposición lista en AutoCAD.", "ok")
            self._log(
                "Si el perímetro del plano (XREF) coincide con el arte → ✔ correcto.",
                "ok",
            )
        except RuntimeError as e:
            self._log(str(e), "err")
        except Exception as e:
            self._log(f"ERROR: {e}", "err")
        finally:
            self._busy(False)


# ─────────────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    ArteMakerApp().mainloop()
