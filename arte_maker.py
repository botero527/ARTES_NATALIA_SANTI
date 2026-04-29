
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


# ─── PALETA ──────────────────────────────────────────────────────────────────
C = {
    "bg":        "#0A0F1E",
    "bg2":       "#0D1426",
    "panel":     "#111827",
    "panel2":    "#162033",
    "border":    "#1E3A5F",
    "accent":    "#00D4FF",
    "accent2":   "#0099CC",
    "accent3":   "#00FF88",
    "btn_ok":    "#00C876",
    "btn_ok2":   "#009955",
    "btn_warn":  "#FF8C00",
    "btn_warn2": "#CC6600",
    "txt":       "#E8F4FD",
    "txt_dim":   "#5A7A9A",
    "txt_mid":   "#8AADCC",
    "entry_bg":  "#0D1A2E",
    "entry_fg":  "#00D4FF",
    "log_bg":    "#060C18",
    "log_ok":    "#00FF88",
    "log_warn":  "#FFB800",
    "log_err":   "#FF4466",
    "log_dim":   "#4A6A8A",
}

FONT_TITLE  = ("Segoe UI", 15, "bold")
FONT_HDR    = ("Segoe UI", 11, "bold")
FONT_BODY   = ("Segoe UI", 10)
FONT_SMALL  = ("Segoe UI",  8)
FONT_LOG    = ("Consolas",  9)
FONT_MONO   = ("Consolas", 10, "bold")

SCRIPT_RHINO = os.path.join(os.path.dirname(os.path.abspath(__file__)), "arte_script.py")


# ─── HELPERS ─────────────────────────────────────────────────────────────────

def _codigo_base(ruta_archivo: str) -> str:
    base = os.path.splitext(os.path.basename(ruta_archivo))[0].strip()
    return base.rsplit("_", 1)[0].strip() if "_" in base else base


def _buscar_artes(ruta: str, codigo: str) -> list:
    """Busca recursivamente en ARTES y en todas sus subcarpetas."""
    resultados = []
    codigo_cmp = codigo.upper().strip()
    for raiz, dirs, archivos in os.walk(ruta):
        dirs[:] = [d for d in dirs if not d.startswith(".")]
        partes = raiz.replace("\\", "/").upper().split("/")

        # Estamos dentro de ARTES si alguna parte del path es "ARTES"
        if "ARTES" not in partes:
            continue

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


def _ruta_planos(ruta_base: str) -> str:
    """
    Lógica de destino para EXTRAER PLANO:
    - Si ruta_base tiene carpeta ARTES  → PLANOS va dentro de ARTES
    - Si no tiene ARTES                 → PLANOS va directo en ruta_base
    En ambos casos crea la carpeta si no existe.
    """
    artes = os.path.join(ruta_base, "ARTES")
    if os.path.isdir(artes):
        destino = os.path.join(artes, "PLANOS")
    else:
        destino = os.path.join(ruta_base, "PLANOS")
    os.makedirs(destino, exist_ok=True)
    return destino


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
                "Abre AutoCAD primero y vuelve a intentarlo."
            )
        
        log_fn(f"  Abriendo: {os.path.basename(ruta_arte)}")
        doc = acad.Documents.Open(os.path.abspath(ruta_arte), False, False)
        time.sleep(2)

        log_fn(f"  Superponiendo plano como XREF: {os.path.basename(ruta_plano)}")
        abs_plano = os.path.abspath(ruta_plano)
        try:
            pt = win32com.client.VARIANT(
                pythoncom.VT_ARRAY | pythoncom.VT_R8, [0.0, 0.0, 0.0]
            )
            msp = doc.ModelSpace
            msp.AttachExternalReference(
                abs_plano, "PLANO_REF", pt,
                1.0, 1.0, 1.0, 0.0, False,
            )
            log_fn("  XREF adjuntado via API COM.")
        except Exception as e_com:
            log_fn(f"  API COM falló ({e_com}), usando SendCommand...")
            doc.SendCommand(f'-XREF A "{abs_plano}" \n0,0,0\n1\n1\n0\n')

        time.sleep(1.5)
        doc.SendCommand("ZOOM E \n")
        time.sleep(0.5)
        log_fn("  Listo. Verifica que el perímetro del plano coincida con el arte.")
    finally:
        pythoncom.CoUninitialize()
# ─── WIDGET HELPERS ──────────────────────────────────────────────────────────

class NeonButton(tk.Frame):
    """Botón con borde de color y efecto hover — compatible Python 3.14."""
    def __init__(self, parent, text, command, color, hover_color,
                 width=180, height=40):
        super().__init__(parent, bg=color, padx=2, pady=2, cursor="hand2")
        self._cmd        = command
        self._color      = color
        self._hover      = hover_color
        self._enabled    = True

        self._lbl = tk.Label(self, text=text, font=FONT_HDR,
                             bg=color, fg="white",
                             padx=14, pady=8, cursor="hand2")
        self._lbl.pack(fill="both", expand=True)

        for w in (self, self._lbl):
            w.bind("<Enter>",    lambda e: self._on_enter())
            w.bind("<Leave>",    lambda e: self._on_leave())
            w.bind("<Button-1>", lambda e: self._click())

    def _on_enter(self):
        if self._enabled:
            self.configure(bg=self._hover)
            self._lbl.configure(bg=self._hover)

    def _on_leave(self):
        col = self._color if self._enabled else C["txt_dim"]
        self.configure(bg=col)
        self._lbl.configure(bg=col)

    def _click(self):
        if not self._enabled:
            return
        self.configure(bg="white")
        self._lbl.configure(bg="white", fg=self._color)
        self.after(130, self._restore)
        self._cmd()

    def _restore(self):
        self.configure(bg=self._color)
        self._lbl.configure(bg=self._color, fg="white")

    def configure_state(self, enabled: bool):
        self._enabled = enabled
        col = self._color if enabled else C["txt_dim"]
        self.configure(bg=col)
        self._lbl.configure(bg=col)


class GlowEntry(tk.Frame):
    """Entry con borde que brilla al tener foco."""
    def __init__(self, parent, textvariable, **kw):
        super().__init__(parent, bg=C["border"], padx=1, pady=1)
        self._var = textvariable
        self._entry = tk.Entry(self, textvariable=textvariable,
                               bg=C["entry_bg"], fg=C["entry_fg"],
                               insertbackground=C["accent"],
                               relief="flat", font=FONT_BODY,
                               bd=4, **kw)
        self._entry.pack(fill="both", expand=True)
        self._entry.bind("<FocusIn>",  lambda e: self.configure(bg=C["accent"]))
        self._entry.bind("<FocusOut>", lambda e: self.configure(bg=C["border"]))

    def get(self):
        return self._var.get()


class ScanLine(tk.Canvas):
    """Línea animada tipo 'escaneo' en el header."""
    def __init__(self, parent, **kw):
        super().__init__(parent, height=3,
                         bg=C["bg"], highlightthickness=0, **kw)
        self._x = 0
        self.bind("<Map>", self._on_map)

    def _on_map(self, _event=None):
        self.unbind("<Map>")
        self._w = self.winfo_width() or 960
        self._animate()

    def _animate(self):
        if not self.winfo_exists():
            return
        self.delete("all")
        for i in range(60):
            x0 = self._x + (i - 30) * 4
            x1 = x0 + 4
            if 0 <= x0 <= self._w:
                self.create_line(x0, 1, x1, 1, fill=C["accent"], width=2)
        self._x = (self._x + 6) % (self._w + 120)
        self.after(20, self._animate)


# ─── APP PRINCIPAL ────────────────────────────────────────────────────────────

class ArteMakerApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("AGP GROUP — Arte Maker")
        self.configure(bg=C["bg"])
        self.resizable(True, True)
        self.minsize(860, 680)

        self._ruta_base    = tk.StringVar()
        self._dwg_plano    = tk.StringVar()
        self._resultados: list = []
        self._dot_count    = 0

        self._apply_ttk_style()
        self._build_ui()
        self._centrar(960, 760)

    # ── TTK style ─────────────────────────────────────────────────────────────

    def _apply_ttk_style(self):
        style = ttk.Style(self)
        style.theme_use("clam")

        style.configure("Treeview",
                         background=C["panel"],
                         foreground=C["txt_mid"],
                         fieldbackground=C["panel"],
                         borderwidth=0,
                         font=FONT_BODY,
                         rowheight=26)
        style.configure("Treeview.Heading",
                         background=C["border"],
                         foreground=C["accent"],
                         font=("Segoe UI", 9, "bold"),
                         relief="flat")
        style.map("Treeview",
                  background=[("selected", C["border"])],
                  foreground=[("selected", C["accent"])])

        style.configure("Vertical.TScrollbar",
                         background=C["panel2"],
                         troughcolor=C["bg2"],
                         arrowcolor=C["accent"],
                         borderwidth=0)

    # ── Build UI ──────────────────────────────────────────────────────────────

    def _build_ui(self):
        # ── HEADER ──
        hdr = tk.Frame(self, bg=C["bg2"])
        hdr.pack(fill="x")

        tk.Frame(hdr, bg=C["accent"], height=2).pack(fill="x")

        inner_hdr = tk.Frame(hdr, bg=C["bg2"], pady=14, padx=24)
        inner_hdr.pack(fill="x")
        tk.Label(inner_hdr, text="AGP GROUP", font=("Segoe UI", 9, "bold"),
                 bg=C["bg2"], fg=C["txt_dim"]).pack(anchor="w")
        tk.Label(inner_hdr, text="ARTE  MAKER",
                 font=("Segoe UI", 20, "bold"),
                 bg=C["bg2"], fg=C["accent"]).pack(anchor="w")

        ScanLine(hdr).pack(fill="x")
        tk.Frame(hdr, bg=C["border"], height=1).pack(fill="x")

        # línea inferior decorativa (se empaca primero para quedar al fondo)
        tk.Frame(self, bg=C["accent"], height=2).pack(fill="x", side="bottom")

        # ── BODY ──
        body = tk.Frame(self, bg=C["bg"], padx=24, pady=16)
        body.pack(fill="both", expand=True)

        # ── TARJETA: inputs ──
        outer_in, card_in = self._card(body, "  CONFIGURACIÓN")
        outer_in.pack(fill="x", pady=(0, 10))
        card_in.columnconfigure(1, weight=1)

        # Ruta base
        self._lbl_field(card_in, "Ruta del vehiculo / modelo / version:", 0)
        self._entry_row = GlowEntry(card_in, self._ruta_base)
        self._entry_row.grid(row=1, column=0, columnspan=2, sticky="ew",
                             padx=(0, 8), pady=(2, 2))
        tk.Button(card_in, text="Explorar…",
                  bg=C["border"], fg=C["accent"], relief="flat",
                  font=FONT_SMALL, cursor="hand2",
                  activebackground=C["accent"], activeforeground=C["bg"],
                  command=self._explorar_base
                  ).grid(row=1, column=2, pady=(2, 2), padx=(4, 0))
        tk.Label(card_in,
                 text="  Puede ser la carpeta del vehículo, modelo o versión — la búsqueda es recursiva",
                 font=FONT_SMALL, bg=C["panel"], fg=C["txt_dim"]
                 ).grid(row=2, column=0, columnspan=3, sticky="w", pady=(0, 8))

        # Plano DWG
        self._lbl_field(card_in, "Plano DWG original:", 3)
        self._entry_dwg = GlowEntry(card_in, self._dwg_plano)
        self._entry_dwg.grid(row=4, column=0, columnspan=2, sticky="ew",
                              padx=(0, 8), pady=(2, 2))
        tk.Button(card_in, text="Explorar…",
                  bg=C["border"], fg=C["accent"], relief="flat",
                  font=FONT_SMALL, cursor="hand2",
                  activebackground=C["accent"], activeforeground=C["bg"],
                  command=self._explorar_dwg
                  ).grid(row=4, column=2, pady=(2, 2), padx=(4, 0))
        tk.Label(card_in, text="  Necesario para EXTRAER PLANO y para la superposición",
                 font=FONT_SMALL, bg=C["panel"], fg=C["txt_dim"]
                 ).grid(row=5, column=0, columnspan=3, sticky="w", pady=(0, 4))

        # ── TARJETA: botones ──
        card_btn = tk.Frame(body, bg=C["bg"])
        card_btn.pack(fill="x", pady=(0, 10))

        self._btn_extraer = NeonButton(
            card_btn, "▶  EXTRAER PLANO",
            self._extraer, C["btn_ok"], C["btn_ok2"], width=200, height=44)
        self._btn_extraer.pack(side="left", padx=(0, 16))

        self._btn_comprobar = NeonButton(
            card_btn, "◉  COMPROBAR ARTE",
            self._comprobar, C["btn_warn"], C["btn_warn2"], width=210, height=44)
        self._btn_comprobar.pack(side="left")

        self._lbl_status = tk.Label(card_btn, text="",
                                    font=FONT_SMALL, bg=C["bg"], fg=C["accent"])
        self._lbl_status.pack(side="left", padx=16)

        # ── TARJETA: tabla ──
        outer_tbl, card_tbl = self._card(
            body, "  ARTES ENCONTRADOS  — doble clic en verde para superponer")
        outer_tbl.pack(fill="x", pady=(0, 10))
        card_tbl.columnconfigure(0, weight=1)
        card_tbl.rowconfigure(0, weight=1)

        cols = ("estado", "ruta", "archivo")
        self._tree = ttk.Treeview(card_tbl, columns=cols, show="headings", height=6)
        self._tree.heading("estado",  text="Estado")
        self._tree.heading("ruta",    text="Ruta relativa")
        self._tree.heading("archivo", text="Archivo")
        self._tree.column("estado",  width=120, anchor="center", stretch=False)
        self._tree.column("ruta",    width=380)
        self._tree.column("archivo", width=280)
        self._tree.tag_configure("match", background="#0A2010", foreground=C["log_ok"])
        self._tree.tag_configure("other", background=C["panel2"], foreground=C["txt_dim"])
        self._tree.grid(row=0, column=0, sticky="nsew")

        sb = ttk.Scrollbar(card_tbl, orient="vertical", command=self._tree.yview)
        sb.grid(row=0, column=1, sticky="ns")
        self._tree.configure(yscrollcommand=sb.set)
        self._tree.bind("<Double-1>", self._on_doble_click)

        # ── TARJETA: log ──
        outer_log, card_log = self._card(body, "  CONSOLA")
        outer_log.pack(fill="both", expand=True)
        card_log.columnconfigure(0, weight=1)
        card_log.rowconfigure(0, weight=1)

        self._log_w = tk.Text(card_log, bg=C["log_bg"], fg=C["txt"],
                               font=FONT_LOG, relief="flat", state="disabled",
                               wrap="word", bd=0)
        self._log_w.grid(row=0, column=0, sticky="nsew")
        for tag, color in [("ok",  C["log_ok"]), ("warn", C["log_warn"]),
                            ("err", C["log_err"]), ("dim", C["log_dim"])]:
            self._log_w.tag_config(tag, foreground=color)

        sb2 = ttk.Scrollbar(card_log, orient="vertical", command=self._log_w.yview)
        sb2.grid(row=0, column=1, sticky="ns")
        self._log_w.configure(yscrollcommand=sb2.set)

    def _card(self, parent, title=""):
        """Retorna (outer, inner): outer se coloca con pack/grid, inner recibe widgets."""
        outer = tk.Frame(parent, bg=C["border"], padx=1, pady=1)
        if title:
            tk.Label(outer, text=title, font=("Segoe UI", 8, "bold"),
                     bg=C["border"], fg=C["txt_dim"]).pack(anchor="w", padx=6, pady=(3, 0))
        inner = tk.Frame(outer, bg=C["panel"], padx=12, pady=10)
        inner.pack(fill="both", expand=True)
        return outer, inner

    def _lbl_field(self, parent, text, row):
        tk.Label(parent, text=text, font=("Segoe UI", 9, "bold"),
                 bg=C["panel"], fg=C["txt_mid"], anchor="w"
                 ).grid(row=row, column=0, columnspan=3, sticky="w", pady=(8, 0))

    # ── Helpers ───────────────────────────────────────────────────────────────

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
        self._btn_extraer.configure_state(not activo)
        self._btn_comprobar.configure_state(not activo)
        if activo:
            self._dot_count = 0
            self._animar_status()
        else:
            self._lbl_status.configure(text="")
        self.update_idletasks()

    def _animar_status(self):
        if not getattr(self._btn_extraer, "_enabled", True) is False:
            return
        puntos = "●" * (self._dot_count % 4 + 1) + "○" * (3 - self._dot_count % 4)
        self._lbl_status.configure(text=f"  Procesando  {puntos}")
        self._dot_count += 1
        self.after(300, self._animar_status)

    def _explorar_base(self):
        ruta = filedialog.askdirectory(title="Seleccionar carpeta (vehiculo / modelo / version)")
        if ruta:
            self._ruta_base.set(ruta.replace("/", "\\"))

    def _explorar_dwg(self):
        inicial = self._ruta_base.get().strip() or "/"
        ruta = filedialog.askopenfilename(
            title="Seleccionar plano DWG",
            initialdir=inicial,
            filetypes=[("AutoCAD DWG", "*.dwg"), ("Todos", "*.*")],
        )
        if ruta:
            self._dwg_plano.set(ruta.replace("/", "\\"))

    def _validar(self, necesita_dwg=True) -> bool:
        ruta = self._ruta_base.get().strip()
        if not ruta:
            messagebox.showwarning("Campo requerido", "Indica la ruta base.")
            return False
        if not os.path.isdir(ruta):
            messagebox.showerror("Ruta no encontrada",
                                 f"No existe o no es accesible:\n{ruta}")
            return False
        if necesita_dwg:
            dwg = self._dwg_plano.get().strip().strip('"')
            if not dwg:
                messagebox.showwarning("Campo requerido",
                                       "Selecciona el archivo DWG del plano.")
                return False
            if not os.path.isfile(dwg):
                messagebox.showerror("Archivo no encontrado", f"No existe:\n{dwg}")
                return False
        return True

    # ── EXTRAER PLANO ─────────────────────────────────────────────────────────

    def _extraer(self):
        if not self._validar(necesita_dwg=True):
            return
        self._busy(True)
        threading.Thread(target=self._t_extraer, daemon=True).start()

    def _t_extraer(self):
        ruta_base  = self._ruta_base.get().strip()
        ruta_plano = self._dwg_plano.get().strip().strip('"')

        self._log("=" * 56)
        self._log("EXTRAER PLANO — filtrando layers en AutoCAD...", "ok")
        self._log(f"Plano : {os.path.basename(ruta_plano)}", "dim")

        nombre_base   = os.path.splitext(os.path.basename(ruta_plano))[0]
        ruta_destino  = _ruta_planos(ruta_base)           # crea PLANOS/ si no existe
        ruta_filtrada = os.path.join(ruta_destino, f"{nombre_base}_PLANO.dwg")

        self._log(f"Destino: {ruta_filtrada}", "dim")

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

    # ── COMPROBAR ARTE ────────────────────────────────────────────────────────

    def _comprobar(self):
        if not self._validar(necesita_dwg=False):
            return
        self._busy(True)
        threading.Thread(target=self._t_comprobar, daemon=True).start()

    def _t_comprobar(self):
        ruta_base = self._ruta_base.get().strip()
        dwg_plano = self._dwg_plano.get().strip().strip('"')

        self._log("=" * 56)
        self._log("COMPROBAR ARTE — buscando artes...", "ok")
        self._log(f"Buscando en: {ruta_base}", "dim")

        codigo = _codigo_base(dwg_plano) if dwg_plano else ""
        if codigo:
            self._log(f'Código del plano: "{codigo}"', "dim")

        resultados = _buscar_artes(ruta_base, codigo)
        self._resultados = resultados

        self.after(0, self._poblar_tabla, resultados)

        n_match = sum(1 for r in resultados if r["coincide"])
        if not resultados:
            self._log("No se encontraron archivos en carpetas ARTES.", "warn")
        else:
            self._log(
                f"Total: {len(resultados)} archivo(s)  —  coincidencias: {n_match}",
                "ok" if n_match else "warn",
            )
            if n_match:
                self._log("Doble clic en una fila verde para superponer en AutoCAD.", "ok")

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

        dwg_plano = self._dwg_plano.get().strip().strip('"')
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
            _overlay_autocad(ruta_arte, ruta_plano,
                             log_fn=lambda m: self._log(m, "dim"))
            self._log("Superposición lista en AutoCAD.", "ok")
            self._log(
                "Si el perímetro del plano (XREF) coincide con el arte → ✔ correcto.", "ok")
        except RuntimeError as e:
            self._log(str(e), "err")
        except Exception as e:
            self._log(f"ERROR: {e}", "err")
        finally:
            self._busy(False)


# ─────────────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    ArteMakerApp().mainloop()
