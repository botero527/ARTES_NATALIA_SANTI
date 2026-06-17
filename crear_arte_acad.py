# -*- coding: utf-8 -*-
"""
AGP Arte Maker — AutoCAD  (lógica espejo de arte_script.py para Rhino)
Requiere AutoCAD abierto con el plano (_PLANO.dwg) como documento activo.
Ejecutar:  py crear_arte_acad.py
"""

import os, sys, time, math, ctypes, datetime

try:
    import win32com.client
    import pythoncom
except ImportError:
    print("ERROR: falta pywin32.  Ejecuta:  pip install pywin32")
    input("Presiona Enter para salir...")
    sys.exit(1)

# ── Parámetros ────────────────────────────────────────────────────────────────
_DIR          = os.path.dirname(os.path.abspath(__file__))
CAJETIN_DWG   = os.path.join(_DIR, "LAYERS Y CAJETINES 1.dwg")

OFFSET_PERIM  = 0.5
OFFSET_BN_DEG = 2.5
DIVISOR_DEG   = 3
BLOQUE_25     = "25"
LAYER_PLANES  = "PLANES"
LAYER_K2      = "k2"
LAYER_K       = "k"
LAYER_K3      = "k3"
RADIO_MIN     = 15.0

PAT_PERIM = ["=PERIMETRO"]   # match exacto — no coger "PERIMETRO OEM" etc.
PAT_BN    = ["BANDA NEGRA", "BANDANEGRA", "BN", "PHANTOM", "BANDA"]
PAT_LOGO  = ["LOGO", "TRAZABILIDAD"]

LOG_FILE  = os.path.join(_DIR, "arte_acad_log.txt")

# ── Log ───────────────────────────────────────────────────────────────────────

def log(msg):
    print(msg)
    try:
        with open(LOG_FILE, "a", encoding="utf-8") as f:
            f.write(msg + "\n")
    except Exception:
        pass

# ── Helpers COM ───────────────────────────────────────────────────────────────

def pt(x, y, z=0.0):
    return win32com.client.VARIANT(
        pythoncom.VT_ARRAY | pythoncom.VT_R8, [float(x), float(y), float(z)])


def _layer_match(layer_nombre: str, patron: str) -> bool:
    """
    Si el patrón empieza con '=' → match exacto (ignorando mayúsculas).
    Si no → el patrón debe estar contenido en el nombre del layer (substring).
    Layers que contengan 'OEM' siempre se excluyen.
    """
    n = layer_nombre.upper().strip()
    if "OEM" in n:
        return False
    p = patron.upper().strip()
    if p.startswith("="):
        return n == p[1:]
    return p in n


def ents_por_patron(msp, patrones):
    res = []
    for e in msp:
        try:
            if any(_layer_match(e.Layer, p) for p in patrones):
                res.append(e)
        except Exception:
            pass
    return res


def area_bbox(ent):
    try:
        lo, hi = ent.GetBoundingBox()
        return abs(hi[0]-lo[0]) * abs(hi[1]-lo[1])
    except Exception:
        return 0.0


def centro_bbox(ents):
    xs, ys = [], []
    for e in ents:
        try:
            lo, hi = e.GetBoundingBox()
            xs += [lo[0], hi[0]]; ys += [lo[1], hi[1]]
        except Exception:
            pass
    if not xs:
        return None, None
    return (min(xs)+max(xs))/2, (min(ys)+max(ys))/2


def asegurar_layer(doc, nombre, color=None):
    """Crea el layer si no existe. Si se pasa color, siempre lo aplica."""
    try:
        lyr = doc.Layers.Item(nombre)
    except Exception:
        lyr = doc.Layers.Add(nombre)
    if color is not None:
        try:
            lyr.color = color
        except Exception:
            pass


def alerta(titulo, msg):
    ctypes.windll.user32.MessageBoxW(0, msg, titulo, 0x30)


def alerta_stop(titulo, msg):
    ctypes.windll.user32.MessageBoxW(0, msg, titulo, 0x10)


def send_cmd(doc, cmd, retries=8, delay=0.6):
    """SendCommand con retry automático para RPC_E_CALL_REJECTED (AutoCAD ocupado)."""
    for i in range(retries):
        try:
            doc.SendCommand(cmd)
            return
        except Exception as e:
            if i < retries - 1:
                time.sleep(delay)
            else:
                raise


def handles_actuales(msp):
    """Devuelve el set de handles de todos los objetos en msp."""
    h = set()
    for e in msp:
        try:
            h.add(e.Handle)
        except Exception:
            pass
    return h


def objetos_nuevos(msp, handles_antes):
    """Devuelve lista de entidades cuyos handles no estaban antes del import."""
    nuevos = []
    for e in msp:
        try:
            if e.Handle not in handles_antes:
                nuevos.append(e)
        except Exception:
            pass
    return nuevos


def corregir_colores_bylayer(ents):
    """Pone Color=ByLayer (256) en entidades que tengan ByBlock (0) o color forzado."""
    for e in ents:
        try:
            if e.color == 0:           # ByBlock → heredar del layer
                e.color = 256
        except Exception:
            pass


# ── Verificar radios usando bulge ─────────────────────────────────────────────

def verificar_radios(ent, radio_min=15.0):
    malos = []
    try:
        n = ent.NumberOfVertices
        coords = list(ent.Coordinates)
        for i in range(n):
            b = ent.GetBulge(i)
            if abs(b) < 1e-9:
                continue
            xi = (i+1) % n
            x0, y0 = coords[i*2], coords[i*2+1]
            x1, y1 = coords[xi*2], coords[xi*2+1]
            cuerda = math.hypot(x1-x0, y1-y0)
            if cuerda < 1e-9:
                continue
            radio = cuerda / (2.0 * abs(math.sin(2.0 * math.atan(abs(b)))))
            if radio < radio_min:
                malos.append(round(radio, 3))
    except Exception:
        pass
    return malos


# ── Offset hacia adentro ──────────────────────────────────────────────────────

def offset_inward(ent, dist):
    area_orig = area_bbox(ent)
    mejor = None
    mejor_area = 1e18
    todos = []
    for d in [dist, -dist]:
        try:
            results = list(ent.Offset(d))
            todos.extend(results)
        except Exception:
            pass
    for r in todos:
        a = area_bbox(r)
        if a < area_orig and a < mejor_area:
            mejor_area = a
            mejor = r
    if mejor is None and todos:
        mejor = min(todos, key=area_bbox)
    for r in todos:
        if r != mejor:
            try: r.Delete()
            except Exception: pass
    return mejor


# ── Hatch SOLID ───────────────────────────────────────────────────────────────

def hatch_solido(msp, doc, outer, inner, layer):
    try:
        h = msp.AddHatch(0, "SOLID", True)
        h.AppendOuterLoop(
            win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_DISPATCH, [outer]))
        if inner:
            h.AppendInnerLoop(
                win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_DISPATCH, [inner]))
        h.Evaluate()
        h.Layer = layer
        h.color = 256          # ByLayer — hereda color del layer
        doc.Regen(0)
        return h
    except Exception as e:
        log(f"  WARN hatch: {e}")
        return None


# ── Diálogo cajetín ───────────────────────────────────────────────────────────

# Paleta de colores
_C = {
    "bg":        "#12131A",   # fondo principal
    "panel":     "#1C1E2B",   # panel de sección
    "card":      "#22253A",   # tarjeta de campo
    "border":    "#2E3250",   # borde sutil
    "accent":    "#4D7EFF",   # azul AGP
    "accent2":   "#2D55CC",   # azul oscuro hover
    "success":   "#3ECF8E",   # verde confirmación
    "text":      "#E8EAFF",   # texto principal
    "muted":     "#6B7099",   # texto secundario / placeholder
    "auto_fg":   "#4D7EFF",   # auto-rellenado
    "entry_bg":  "#1A1C2A",   # fondo input
    "entry_act": "#1F2238",   # fondo input activo
    "sep":       "#2A2D45",   # separador
    "danger":    "#FF5757",   # cancelar
    "danger2":   "#CC3030",
}


def _rounded_btn(parent, text, cmd, color, hover_color, fg="#FFFFFF",
                 width=160, height=40, radius=10, font_size=11):
    """Botón estilizado compatible con Python 3.14 / Tk 9."""
    import tkinter as tk
    btn = tk.Button(parent, text=text, command=cmd,
                    bg=color, fg=fg, activebackground=hover_color,
                    activeforeground=fg, relief="flat", bd=0,
                    font=("Segoe UI", font_size, "bold"),
                    padx=16, pady=8, cursor="hand2")
    btn.bind("<Enter>", lambda _: btn.configure(bg=hover_color))
    btn.bind("<Leave>", lambda _: btn.configure(bg=color))
    return btn


def _section_header(parent, text):
    import tkinter as tk
    f = tk.Frame(parent, bg=_C["bg"])
    tk.Label(f, text=text, font=("Segoe UI", 8, "bold"),
             fg=_C["accent"], bg=_C["bg"]).pack(side="left", padx=8)
    tk.Frame(f, bg=_C["sep"], height=1).pack(side="left", fill="x", expand=True, pady=1)
    return f


def _make_entry(parent, width_chars=28):
    import tkinter as tk
    var = tk.StringVar()
    e = tk.Entry(parent, textvariable=var, width=width_chars,
                 bg=_C["entry_bg"], fg=_C["text"], insertbackground=_C["text"],
                 relief="flat", font=("Segoe UI", 10),
                 highlightthickness=1, highlightbackground=_C["border"],
                 highlightcolor=_C["accent"], bd=4)
    return e, var


def dialogo_cajetin(nombre_plano="", ruta_salida=""):
    import tkinter as tk
    import re as _r

    hoy = datetime.date.today().strftime("%d.%m.%Y")

    SECCIONES = [
        ("DATOS DEL PLANO", [
            ("DIBUJO",    "Dibujo",        None),
            ("VEHICULO",  "Vehículo",       None),
            ("MODELO",    "Modelo / Año",   "auto"),
            ("COD PLANO", "Código de plano",None),
            ("NAGS",      "NAGS",           "auto"),
            ("VERSION",   "Versión",        "auto"),
            ("PIEZA",     "Pieza",          "auto"),
            ("VITRO",     "Vitro",          None),
            ("MALLA",     "Malla",          None),
        ]),
        ("VALORES POR DEFECTO", [
            ("FECHA",     "Fecha",          hoy),
            ("REVISADO",  "Revisado por",   "Santiago P."),
            ("MEDIDAS",   "Medidas",        "Milimetros"),
            ("VISTA",     "Vista",          "Interna"),
            ("ESCALA",    "Escala",         "1:1"),
        ]),
    ]

    def parsear_cod(cod):
        grupos = _r.findall(r'[\d\-]+', cod.strip())
        grupos = [g.strip('-') for g in grupos if g.strip('-')]
        return (grupos[0] if grupos else "",
                grupos[1] if len(grupos) > 1 else "",
                grupos[2] if len(grupos) > 2 else "")

    def extraer_anio(texto):
        m = _r.search(r'\b(19|20)\d{2}\b', texto)
        return m.group(0) if m else ""

    resultado = [None]

    # Si ya hay una instancia Tk corriendo (ej: arte_maker.py), usar Toplevel modal
    try:
        _existing_root = tk._default_root
    except Exception:
        _existing_root = None

    if _existing_root is not None:
        root = tk.Toplevel(_existing_root)
        root.grab_set()
    else:
        root = tk.Tk()

    root.title("AGP Arte Maker")
    root.resizable(True, True)
    root.attributes("-topmost", True)
    root.configure(bg=_C["bg"])
    root.minsize(560, 400)

    # ── Título / header (fijo, fuera del scroll) ──────────────────────────────
    hdr = tk.Frame(root, bg=_C["accent"], height=4)
    hdr.pack(fill="x")

    title_frame = tk.Frame(root, bg=_C["bg"], pady=14)
    title_frame.pack(fill="x", padx=28)

    tk.Label(title_frame, text="AGP  Arte Maker",
             font=("Segoe UI", 18, "bold"), fg=_C["text"], bg=_C["bg"]).pack(anchor="w")
    tk.Label(title_frame, text=f"Cajetín  ·  {nombre_plano or 'nuevo plano'}",
             font=("Segoe UI", 9), fg=_C["muted"], bg=_C["bg"]).pack(anchor="w")

    # ── Zona scrollable ───────────────────────────────────────────────────────
    from tkinter import ttk as _ttk
    scroll_outer = tk.Frame(root, bg=_C["bg"])
    scroll_outer.pack(fill="both", expand=True, padx=0, pady=0)

    canvas_scroll = tk.Canvas(scroll_outer, bg=_C["bg"],
                              highlightthickness=0, bd=0)
    vsb = _ttk.Scrollbar(scroll_outer, orient="vertical",
                         command=canvas_scroll.yview)
    canvas_scroll.configure(yscrollcommand=vsb.set)
    vsb.pack(side="right", fill="y")
    canvas_scroll.pack(side="left", fill="both", expand=True)

    body = tk.Frame(canvas_scroll, bg=_C["bg"], padx=28)
    _body_win = canvas_scroll.create_window((0, 0), window=body, anchor="nw")

    def _on_body_configure(_e):
        canvas_scroll.configure(scrollregion=canvas_scroll.bbox("all"))
    def _on_canvas_configure(e):
        canvas_scroll.itemconfig(_body_win, width=e.width)
    body.bind("<Configure>", _on_body_configure)
    canvas_scroll.bind("<Configure>", _on_canvas_configure)

    # Scroll con rueda del ratón
    def _on_mousewheel(e):
        try:
            canvas_scroll.yview_scroll(int(-1*(e.delta/120)), "units")
        except Exception:
            pass
    root.bind_all("<MouseWheel>", _on_mousewheel)
    root.bind("<Destroy>", lambda _: root.unbind_all("<MouseWheel>"))

    entries = {}

    for sec_titulo, filas in SECCIONES:
        _section_header(body, sec_titulo).pack(fill="x", pady=8)

        for campo, etiqueta, default in filas:
            row = tk.Frame(body, bg=_C["bg"], pady=3)
            row.pack(fill="x")

            lbl = tk.Label(row, text=etiqueta, font=("Segoe UI", 9),
                           fg=_C["muted"], bg=_C["bg"], width=17, anchor="e")
            lbl.pack(side="left", padx=10)

            ent, var = _make_entry(row, width_chars=30)
            ent.pack(side="left", ipady=5)

            # badge "auto" si aplica
            if default == "auto":
                tk.Label(row, text="auto", font=("Segoe UI", 7, "bold"),
                         fg=_C["accent"], bg=_C["bg"], padx=4).pack(side="left", padx=4)

            if isinstance(default, str) and default != "auto":
                ent.insert(0, default)
                ent.configure(fg=_C["muted"])

                def _on_focus_in(e, w=ent, d=default):
                    if w.get() == d:
                        w.configure(fg=_C["text"])

                def _on_focus_out_default(e, w=ent, d=default):
                    if not w.get().strip():
                        w.delete(0, tk.END)
                        w.insert(0, d)
                        w.configure(fg=_C["muted"])

                ent.bind("<FocusIn>",  _on_focus_in)
                ent.bind("<FocusOut>", _on_focus_out_default)

            entries[campo] = ent

    # ── Auto-rellenado lógica ─────────────────────────────────────────────────
    def on_cod_plano(*_):
        n, v, p = parsear_cod(entries["COD PLANO"].get())
        for k, val in [("NAGS", n), ("VERSION", v), ("PIEZA", p)]:
            entries[k].delete(0, tk.END)
            entries[k].insert(0, val)
            entries[k].configure(fg=_C["auto_fg"])

    def on_vehiculo(*_):
        anio = extraer_anio(entries["VEHICULO"].get())
        if anio:
            entries["MODELO"].delete(0, tk.END)
            entries["MODELO"].insert(0, anio)
            entries["MODELO"].configure(fg=_C["auto_fg"])

    entries["COD PLANO"].bind("<FocusOut>", on_cod_plano)
    entries["COD PLANO"].bind("<Return>",   on_cod_plano)
    entries["VEHICULO"].bind("<FocusOut>",  on_vehiculo)
    entries["VEHICULO"].bind("<Return>",    on_vehiculo)

    # Pre-llenar
    if nombre_plano:
        entries["COD PLANO"].delete(0, tk.END)
        entries["COD PLANO"].insert(0, nombre_plano)
        entries["COD PLANO"].configure(fg=_C["text"])
        on_cod_plano()

    entries["DIBUJO"].focus_set()

    # ── Auto-asignación vitro / malla ─────────────────────────────────────────
    _section_header(body, "ASIGNACIÓN VITRO / MALLA").pack(fill="x", pady=8)

    asig_row = tk.Frame(body, bg=_C["bg"], pady=4)
    asig_row.pack(fill="x")

    asig_lbl = tk.Label(asig_row, text="Asignación automática",
                        font=("Segoe UI", 9), fg=_C["muted"], bg=_C["bg"],
                        width=17, anchor="e")
    asig_lbl.pack(side="left", padx=10)

    asig_preview = tk.Label(asig_row, text="(sin asignar)",
                            font=("Segoe UI", 9, "italic"), fg=_C["muted"], bg=_C["bg"])
    asig_preview.pack(side="left", padx=(0, 8))

    _asignacion_cache = [None]

    def _auto_asignar():
        try:
            from db_app.asignaciones import dialogo_asignacion, confirmar
        except ImportError:
            try:
                import sys as _sys, os as _os
                _sys.path.insert(0, _os.path.join(_os.path.dirname(__file__), "db_app"))
                from asignaciones import dialogo_asignacion, confirmar
            except Exception as ie:
                tk.messagebox.showerror("Error", f"No se pudo cargar módulo asignaciones:\n{ie}")
                return

        prop = dialogo_asignacion(parent=root, nombre_plano=nombre_plano)
        if not prop:
            return

        _asignacion_cache[0] = prop

        partes = []
        if prop.get("vitros"):
            partes.append("Vitros: " + ", ".join(prop["vitros"]))
            entries["VITRO"].delete(0, tk.END)
            entries["VITRO"].insert(0, " / ".join(prop["vitros"]))
            entries["VITRO"].configure(fg=_C["auto_fg"])
        if prop.get("grandes") or prop.get("pequenas"):
            mallas = prop.get("grandes", []) + [str(c) for c in prop.get("pequenas", [])]
            partes.append("Mallas: " + ", ".join(mallas))
            entries["MALLA"].delete(0, tk.END)
            entries["MALLA"].insert(0, " / ".join(mallas))
            entries["MALLA"].configure(fg=_C["auto_fg"])

        asig_preview.configure(text=" | ".join(partes) if partes else "(sin asignación)",
                               fg=_C["success"])

    btn_asig = tk.Button(
        asig_row, text="🎯  Auto-asignar",
        command=_auto_asignar,
        bg=_C["accent2"], fg="#FFFFFF",
        activebackground=_C["accent"], activeforeground="#FFFFFF",
        relief="flat", bd=0, font=("Segoe UI", 9, "bold"),
        padx=10, pady=5, cursor="hand2",
    )
    btn_asig.bind("<Enter>", lambda _: btn_asig.configure(bg=_C["accent"]))
    btn_asig.bind("<Leave>", lambda _: btn_asig.configure(bg=_C["accent2"]))
    btn_asig.pack(side="left")

    # ── Navegación Tab entre campos ───────────────────────────────────────────
    all_entries = [entries[c] for _, filas in SECCIONES for c, _, _ in filas]
    for i, e in enumerate(all_entries):
        nxt = all_entries[(i + 1) % len(all_entries)]
        e.bind("<Tab>", lambda ev, n=nxt: (n.focus_set(), n.select_range(0, tk.END), "break"))
        e.bind("<Return>", lambda ev, n=nxt: (n.focus_set(), n.select_range(0, tk.END), "break"))

    # ── Botones ───────────────────────────────────────────────────────────────
    btn_area = tk.Frame(root, bg=_C["bg"], pady=20, padx=28)
    btn_area.pack(fill="x")

    def aceptar(*_):
        vals = {c: entries[c].get().strip() for _, filas in SECCIONES for c, _, _ in filas}
        resultado[0] = vals

        # Confirmar asignación en BD si el usuario usó auto-asignar
        if _asignacion_cache[0] is not None:
            try:
                try:
                    from db_app.asignaciones import confirmar as _confirmar
                except ImportError:
                    import sys as _sys, os as _os
                    _sys.path.insert(0, _os.path.join(_os.path.dirname(__file__), "db_app"))
                    from asignaciones import confirmar as _confirmar
                # responsable desde campo DIBUJO (None si vacío → NULL en BD, nunca string vacío)
                _resp = vals.get("DIBUJO", "").strip() or None
                _confirmar(
                    _asignacion_cache[0],
                    vehiculo     = vals.get("VEHICULO", ""),
                    version      = vals.get("MODELO", ""),
                    pieza        = vals.get("PIEZA", ""),
                    cod_vehiculo = vals.get("NAGS", ""),
                    cod_completo = vals.get("COD PLANO", "") or nombre_plano,
                    ruta_archivo = ruta_salida or "",
                    responsable  = _resp,
                )
            except Exception as _e:
                print(f"[cajetin] WARN confirmar asignación: {_e}")

        root.destroy()

    def cancelar(*_):
        # Si había reserva pendiente, marcarla como CANCELADA (reciclable)
        if _asignacion_cache[0] is not None:
            try:
                try:
                    from db_app.asignaciones import cancelar as _cancelar_asig
                except ImportError:
                    import sys as _sys, os as _os
                    _sys.path.insert(0, _os.path.join(_os.path.dirname(__file__), "db_app"))
                    from asignaciones import cancelar as _cancelar_asig
                _cancelar_asig(_asignacion_cache[0])
            except Exception as _e:
                print(f"[cajetin] WARN cancelar asignación: {_e}")
        root.destroy()

    _rounded_btn(btn_area, "✔  Aplicar al cajetín", aceptar,
                 _C["accent"], _C["accent2"], width=200, height=42).pack(side="left", padx=12)
    _rounded_btn(btn_area, "✕  Cancelar", cancelar,
                 "#2A2A3A", "#3A3A4A", fg=_C["muted"], width=130, height=42).pack(side="left")

    tk.Label(btn_area, text="Enter / Tab para navegar campos  ·  Esc para cancelar",
             font=("Segoe UI", 7), fg=_C["muted"], bg=_C["bg"]).pack(
             side="right", padx=4)

    root.bind("<Escape>", cancelar)
    root.protocol("WM_DELETE_WINDOW", cancelar)

    # ── Centrar ventana en pantalla (limitar al 90% de altura) ───────────────
    root.update_idletasks()
    sw, sh  = root.winfo_screenwidth(), root.winfo_screenheight()
    w       = root.winfo_reqwidth()
    h       = min(root.winfo_reqheight(), int(sh * 0.90))
    w       = max(w, 560)
    root.geometry(f"{w}x{h}+{(sw-w)//2}+{(sh-h)//2}")

    if _existing_root is not None:
        _existing_root.wait_window(root)
    else:
        root.mainloop()
    return resultado[0]


# ── Actualizar texto del cajetín en AutoCAD ───────────────────────────────────

def actualizar_texto_cajetin(msp, valores):
    """
    Actualiza textos en los sublayers de CAJETIN 1.
    Formato exacto del layer: 'CAJETIN 1$CAMPO 1'
    """
    CAMPO_LAYER = {
        "DIBUJO":    "CAJETIN 1$DIBUJO 1",
        "VEHICULO":  "CAJETIN 1$VEHICULO 1",
        "MODELO":    "CAJETIN 1$MODELO 1",
        "COD PLANO": "CAJETIN 1$COD PLANO 1",
        "NAGS":      "CAJETIN 1$NAGS 1",
        "VERSION":   "CAJETIN 1$VERSION 1",
        "PIEZA":     "CAJETIN 1$PIEZA 1",
        "VITRO":     "CAJETIN 1$VITRO 1",
        "MALLA":     "CAJETIN 1$MALLA 1",
        "FECHA":     "CAJETIN 1$FECHA 1",
        "REVISADO":  "CAJETIN 1$REVISADO 1",
        "MEDIDAS":   "CAJETIN 1$MEDIDAS 1",
        "VISTA":     "CAJETIN 1$VISTA 1",
        "ESCALA":    "CAJETIN 1$ESCALA 1",
    }

    actualizados = 0
    for ent in msp:
        try:
            obj_name = ent.ObjectName
            if obj_name not in ("AcDbText", "AcDbMText"):
                continue
            layer = ent.Layer.upper()
            for campo, layer_target in CAMPO_LAYER.items():
                if layer == layer_target.upper():
                    val = valores.get(campo, "")
                    if val:
                        ent.TextString = val
                        actualizados += 1
                    break
        except Exception:
            pass
    log(f"  Cajetín: {actualizados} texto(s) actualizados.")


# ── Pipeline principal ────────────────────────────────────────────────────────

def pipeline(doc, log_fn=None, valores_cajetin=None, ruta_salida=None, perim_index=0, compensar=False):
    """
    Pipeline completo de creación de arte.
    doc            : AutoCAD Document COM object ya abierto.
    log_fn         : función de logging (default: log del módulo).
    valores_cajetin: dict con datos del cajetín; si None muestra el diálogo.
    ruta_salida    : ruta .dwg donde guardar; si None no guarda.
    perim_index    : índice de la pieza a procesar (0=más grande, 1=segunda, etc.)
    """
    if log_fn is None:
        log_fn = log

    msp = doc.ModelSpace
    nombre_plano = os.path.splitext(doc.Name)[0]
    log_fn(f"\n=== AGP Arte Maker AutoCAD: {doc.Name} ===")

    # ── 1. Quitar GlobalWidth ──────────────────────────────────────────────
    log_fn("[1] Limpiando GlobalWidth a 0...")
    for e in msp:
        try:
            if "Polyline" in e.ObjectName:
                e.ConstantWidth = 0.0
                e.Update()
        except Exception:
            pass

    # ── 2. Verificar contornos cerrados ───────────────────────────────────
    log_fn("[2] Verificando contornos cerrados...")
    no_cerrados = []
    for e in ents_por_patron(msp, PAT_PERIM + PAT_BN):
        try:
            if not e.Closed:
                no_cerrados.append(e.Layer)
        except Exception:
            pass
    if no_cerrados:
        msg = "ALERTA: Contornos NO cerrados en:\n" + "\n".join(set(no_cerrados)) + \
              "\n\nCorrige y vuelve a ejecutar."
        log_fn(f"ERROR: {msg}")
        alerta_stop("AGP Arte Maker — Error", msg)
        return
    log_fn("  Todos los contornos cerrados ✔")

    # ── 3. Radios mínimos ─────────────────────────────────────────────────
    log_fn(f"[3] Verificando radios (mín {RADIO_MIN} mm)...")
    radios_malos = []
    for e in ents_por_patron(msp, PAT_PERIM):
        radios_malos += verificar_radios(e, RADIO_MIN)
    if radios_malos:
        r_min = min(radios_malos)
        log_fn(f"  WARN radios: {sorted(set(radios_malos))[:8]}")
        alerta("AGP Arte Maker — Advertencia radios",
               f"Se detectaron radios menores a {RADIO_MIN} mm en el perímetro.\n"
               f"Radio mínimo encontrado: {r_min:.3f} mm\n\nEl proceso continuará.")
    else:
        log_fn("  Radios OK ✔")

    # ── 4. Todos los BN/PHANTOM cerrados ─────────────────────────────────
    def _es_poly_cerrada(e):
        try:
            return "Polyline" in e.ObjectName and e.Closed
        except Exception:
            return False
    bn_ents_todos = sorted(
        [e for e in ents_por_patron(msp, PAT_BN) if _es_poly_cerrada(e)],
        key=area_bbox, reverse=True)
    log_fn(f"[4] BN cerrados totales en doc: {len(bn_ents_todos)}")

    # ── 5. Perímetro — tomar el más grande ────────────────────────────────
    perim_ents = sorted(
        [e for e in ents_por_patron(msp, PAT_PERIM) if _es_poly_cerrada(e)],
        key=area_bbox, reverse=True)
    if not perim_ents:
        alerta_stop("AGP Arte Maker — Error", "No se encontró curva PERIMETRO cerrada.")
        return
    n_piezas = len(perim_ents)
    idx = min(perim_index, n_piezas - 1)
    perim_ent = perim_ents[idx]
    if n_piezas > 1:
        log_fn(f"  {n_piezas} piezas encontradas — procesando pieza {idx + 1}/{n_piezas} ✔")
    else:
        log_fn("  Perímetro encontrado ✔")

    # ── 5b. Compensación de perímetro (offset 3 mm hacia adentro) ────────────
    _perim_original = None
    if compensar:
        log_fn("[5b] Compensando perímetro (offset 3 mm hacia adentro)...")
        comp_perim = offset_inward(perim_ent, 3.0)
        if comp_perim is None:
            alerta_stop("AGP Arte Maker — Error",
                        "No se pudo crear el offset de compensación (3 mm).")
            return
        comp_perim.Layer = perim_ent.Layer
        _perim_original = perim_ent         # guardar para borrar al final
        perim_ent = comp_perim
        msp = doc.ModelSpace
        log_fn("  Perímetro compensado ✔")

    # Filtrar BN que pertenecen a ESTA pieza (centro dentro del bbox del perim)
    def _bbox_ent(e):
        try:
            lo, hi = e.GetBoundingBox()
            return lo[0], lo[1], hi[0], hi[1]
        except Exception:
            return None

    p_bbox = _bbox_ent(perim_ent)
    if p_bbox:
        px0, py0, px1, py1 = p_bbox
        # Ampliar un poco el bbox para tolerar entidades en el borde
        mx = (px1 - px0) * 0.1
        my = (py1 - py0) * 0.1
        def _centro_ent(e):
            b = _bbox_ent(e)
            if not b: return None, None
            return (b[0]+b[2])/2, (b[1]+b[3])/2
        bn_ents = [e for e in bn_ents_todos
                   if (lambda cx, cy: cx is not None
                       and px0 - mx <= cx <= px1 + mx
                       and py0 - my <= cy <= py1 + my)(*_centro_ent(e))]
    else:
        bn_ents = bn_ents_todos

    # 1 BN = banda sólida (sin degradé), 2+ BN = exterior+interior → degradé
    CON_DEGRADE = len(bn_ents) >= 2
    bn_ent = bn_ents[0] if bn_ents else None
    log_fn(f"  BN de esta pieza: {len(bn_ents)} → {'CON degradé' if CON_DEGRADE else 'SIN degradé'}")

    # ── Capturar handles del logo ANTES del import ────────────────────────
    _logo_handles = set()
    for _e in ents_por_patron(msp, PAT_LOGO):
        try: _logo_handles.add(_e.Handle)
        except Exception: pass
    log_fn(f"  Logo en plano: {len(_logo_handles)} objeto(s)")

    handles_antes = handles_actuales(msp)

    # ── 6. Importar cajetines ──────────────────────────────────────────────
    log_fn("[6] Importando cajetines...")
    abs_caj = os.path.abspath(CAJETIN_DWG)
    if not os.path.isfile(abs_caj):
        log_fn(f"  WARN: no se encontró {abs_caj}")
    else:
        try:
            pt_ins = win32com.client.VARIANT(
                pythoncom.VT_ARRAY | pythoncom.VT_R8, [0.0, 0.0, 0.0])
            blk_ref = msp.InsertBlock(pt_ins, abs_caj, 1.0, 1.0, 1.0, 0.0)
            log_fn("  Bloque insertado vía COM ✔")

            nivel1 = []
            try:
                nivel1 = list(blk_ref.Explode())
                log_fn(f"  Nivel 1 explosión: {len(nivel1)} objetos")
            except Exception as ex1:
                log_fn(f"  WARN explosión nivel 1: {ex1}")

            n2 = 0
            for e2 in nivel1:
                try:
                    if e2.ObjectName == "AcDbBlockReference":
                        e2.Explode()
                        n2 += 1
                except Exception:
                    pass
            if n2:
                log_fn(f"  Nivel 2 explosión: {n2} bloques anidados ✔")
            doc.Regen(0)
            time.sleep(0.25)
            msp = doc.ModelSpace

        except Exception as e_ins:
            log_fn(f"  WARN InsertBlock COM falló ({e_ins}), usando SendCommand...")
            send_cmd(doc, f'-INSERT "{abs_caj}"\n0,0,0\n1\n1\n0\n')
            time.sleep(3)
            send_cmd(doc, "EXPLODE\nL\n\n")
            time.sleep(2)
            msp = doc.ModelSpace

    nuevos = objetos_nuevos(msp, handles_antes)
    log_fn(f"  Objetos nuevos detectados: {len(nuevos)}")
    _layers_vistos = set()
    for _e in nuevos[:40]:
        try: _layers_vistos.add(_e.Layer)
        except Exception: pass
    if _layers_vistos:
        log_fn(f"  Layers detectados: {sorted(_layers_vistos)[:8]}")
    corregir_colores_bylayer(nuevos)

    # ── 7. Buscar CAJETIN y LOGO1 ─────────────────────────────────────────
    logo1_ents, cajetin_ents = [], []
    for _e in msp:
        try:
            ly = _e.Layer.upper()
            if "LOGO1"   in ly: logo1_ents.append(_e)
            if "CAJETIN" in ly: cajetin_ents.append(_e)
        except Exception:
            pass
    log_fn(f"  CAJETIN en msp: {len(cajetin_ents)} obj  |  LOGO1: {len(logo1_ents)} obj")

    borrados = 0
    for e in nuevos:
        try:
            ly = e.Layer.upper()
            if "CAJETIN 1" in ly or "LOGO1" in ly:
                continue
            e.Delete()
            borrados += 1
        except Exception:
            pass
    log_fn(f"  Limpieza: {borrados} objeto(s) sobrantes eliminados ✔")
    msp = doc.ModelSpace

    # ── 8. Reemplazar logo ────────────────────────────────────────────────
    log_fn("[8] Reemplazando logo...")
    logo_plano = []
    logo1_ents = []
    for _e in msp:
        try:
            _h = _e.Handle
            _ly = _e.Layer.upper()
            if _h in _logo_handles:
                logo_plano.append(_e)
            if "LOGO1" in _ly:
                logo1_ents.append(_e)
        except Exception:
            pass
    log_fn(f"  logo_plano: {len(logo_plano)}  logo1: {len(logo1_ents)}")
    try:
        if logo_plano and logo1_ents:
            cx_pl, cy_pl = centro_bbox(logo_plano)
            cx_l1, cy_l1 = centro_bbox(logo1_ents)
            if cx_pl is not None and cx_l1 is not None:
                dx, dy = cx_pl - cx_l1, cy_pl - cy_l1
                for e in logo1_ents:
                    try: e.Move(pt(0,0), pt(dx,dy))
                    except Exception: pass
                for e in logo_plano:
                    try: e.Delete()
                    except Exception: pass
                log_fn("  Logo reemplazado ✔")
        elif not logo1_ents:
            log_fn("  WARN: LOGO1 no encontrado en el cajetín.")
        else:
            log_fn("  WARN: no hay logo en el plano.")
    except Exception as e_logo:
        log_fn(f"  WARN logo: {e_logo}")

    # ── 9. Centrar cajetín sobre la pieza ─────────────────────────────────
    log_fn("[9] Centrando cajetín...")
    msp = doc.ModelSpace
    caj_ents = []
    for _e in msp:
        try:
            if "CAJETIN" in _e.Layer.upper():
                caj_ents.append(_e)
        except Exception:
            pass
    log_fn(f"  Buscando CAJETIN en msp: {len(caj_ents)} objetos")
    if caj_ents:
        cx_p, cy_p = centro_bbox([perim_ent])
        cx_c, cy_c = centro_bbox(caj_ents)
        log_fn(f"  Centro pieza: ({cx_p:.1f},{cy_p:.1f})  Centro cajetín: ({cx_c:.1f},{cy_c:.1f})")
        if cx_p is not None and cx_c is not None:
            dx, dy = cx_p - cx_c, cy_p - cy_c
            for e in caj_ents:
                try: e.Move(pt(0,0), pt(dx,dy))
                except Exception: pass
            log_fn("  Cajetín centrado ✔")
    else:
        log_fn("  WARN: no se encontró ningún objeto con layer CAJETIN.")

    # ── 10. Crear layers de arte con colores correctos ───────────────────
    asegurar_layer(doc, LAYER_PLANES, color=8)   # gris
    asegurar_layer(doc, LAYER_K2,     color=3)   # verde
    asegurar_layer(doc, LAYER_K,      color=5)   # azul
    asegurar_layer(doc, LAYER_K3,     color=1)   # rojo

    # ── 11. Offset perímetro 0.5 ──────────────────────────────────────────
    log_fn(f"[11] Offset perímetro {OFFSET_PERIM} mm...")
    off_perim = offset_inward(perim_ent, OFFSET_PERIM)
    if not off_perim:
        alerta_stop("AGP Arte Maker — Error", "No se pudo crear offset del perímetro.")
        return
    off_perim.Layer = LAYER_PLANES

    # Refrescar msp antes de hatches — operaciones anteriores pueden haber dejado COM stale
    msp = doc.ModelSpace

    # ── 12. Hatch k2 ──────────────────────────────────────────────────────
    log_fn("[12] Hatch k2...")
    hatch_solido(msp, doc, perim_ent, off_perim, LAYER_K2)
    time.sleep(0.3)
    msp = doc.ModelSpace

    # ── 13. Hatch k ───────────────────────────────────────────────────────
    log_fn("[13] Hatch k...")
    if bn_ent:
        hatch_solido(msp, doc, bn_ent, off_perim, LAYER_K)
        time.sleep(0.3)
        msp = doc.ModelSpace
    else:
        log_fn("  WARN: no se encontró banda negra.")

    # ── 14. Degradé ───────────────────────────────────────────────────────
    if CON_DEGRADE and bn_ent:
        log_fn(f"[14] Degradé: offset BN {OFFSET_BN_DEG} mm...")
        off_bn = offset_inward(bn_ent, OFFSET_BN_DEG)
        if off_bn:
            off_bn.Layer = LAYER_PLANES
            longitud = float(off_bn.Length)
            n_pepas  = int(round(longitud / DIVISOR_DEG))
            log_fn(f"  Longitud: {longitud:.2f} mm  pepas: {n_pepas}")

            if n_pepas > 0:
                bloque_existe = False
                try:
                    doc.Blocks.Item(BLOQUE_25)
                    bloque_existe = True
                    log_fn(f"  Bloque '{BLOQUE_25}' ya existe.")
                except Exception:
                    pass

                if not bloque_existe:
                    log_fn(f"  Importando bloque '{BLOQUE_25}' desde cajetines...")
                    send_cmd(doc, f'-INSERT "{abs_caj}"\n0,0,0\n1\n1\n0\n')
                    time.sleep(3.0)
                    send_cmd(doc, "ERASE\nL\n\n")
                    time.sleep(1.0)
                    log_fn(f"  Bloque '{BLOQUE_25}' registrado.")

                def ejecutar_divide(handle_curva):
                    send_cmd(doc, f'CLAYER\n{LAYER_K3}\n')
                    time.sleep(0.2)
                    send_cmd(doc,
                        f'DIVIDE\n'
                        f'(handent "{handle_curva}")\n'
                        f'B\n'
                        f'{BLOQUE_25}\n'
                        f'Y\n'
                        f'{n_pepas}\n'
                    )
                    espera = max(2.5, n_pepas * 0.015)
                    log_fn(f"  Esperando {espera:.1f}s (DIVIDE)...")
                    time.sleep(espera)
                    send_cmd(doc, "CLAYER\n0\n")
                    time.sleep(0.2)

                def borrar_k3():
                    _msp = doc.ModelSpace
                    for _e in _msp:
                        try:
                            if _e.Layer.upper() == LAYER_K3.upper():
                                _e.Delete()
                        except Exception:
                            pass
                    time.sleep(0.2)

                ejecutar_divide(off_bn.Handle)
                log_fn("  Degradé en layer k3 ✔")

                respuesta = ctypes.windll.user32.MessageBoxW(
                    0,
                    "¿El degradé quedó correcto?\n"
                    "(Las bolas GRANDES deben estar hacia el borde azul/BN)\n\n"
                    "Sí = quedó bien\n"
                    "No = invertir automáticamente",
                    "AGP Arte Maker — Verificar degradé",
                    0x24
                )
                if respuesta == 7:
                    log_fn("  Invirtiendo degradé...")
                    borrar_k3()
                    h_rev = off_bn.Handle
                    send_cmd(doc, f'REVERSE\n(handent "{h_rev}")\n\n')
                    time.sleep(0.4)
                    ejecutar_divide(off_bn.Handle)
                    log_fn("  Degradé invertido ✔")
                else:
                    log_fn("  Degradé confirmado por usuario ✔")
        else:
            log_fn("  WARN: no se pudo crear offset interior de BN.")
    else:
        log_fn("[14] Sin degradé — omitido.")

    # ── 15. Mover PERIMETRO/BN a PLANES ──────────────────────────────────
    log_fn("[15] Moviendo geometría original a PLANES...")
    time.sleep(0.5)
    msp = doc.ModelSpace
    for e in ents_por_patron(msp, PAT_PERIM + PAT_BN):
        try: e.Layer = LAYER_PLANES
        except Exception: pass

    # Si compensamos: devolver el offset a PERIMETRO y borrar el original
    if compensar:
        try:
            perim_ent.Layer = "PERIMETRO"
        except Exception:
            pass
        if _perim_original is not None:
            try:
                _perim_original.Delete()
                log_fn("  Perímetro original eliminado ✔")
            except Exception:
                pass

    log_fn("=== Arte base completado ✔ ===")

    # ── 17. Cajetín ───────────────────────────────────────────────────────
    log_fn("[17] Aplicando datos del cajetín...")
    if valores_cajetin is None:
        valores_cajetin = dialogo_cajetin(nombre_plano, ruta_salida=ruta_salida or "")

    if valores_cajetin:
        actualizar_texto_cajetin(doc.ModelSpace, valores_cajetin)
        log_fn("  Cajetín aplicado ✔")
    else:
        log_fn("  Cajetín: cancelado por el usuario.")

    # ── 18. Consolidar sublayers CAJETIN 1$* → CAJETIN1 ─────────────────
    log_fn("[18] Consolidando layers CAJETIN 1$* → CAJETIN1...")
    asegurar_layer(doc, "CAJETIN1")
    consolidados = 0
    for _e in doc.ModelSpace:
        try:
            if _e.Layer.upper().startswith("CAJETIN 1$") or _e.Layer.upper() == "CAJETIN 1":
                _e.Layer = "CAJETIN1"
                consolidados += 1
        except Exception:
            pass
    log_fn(f"  {consolidados} objeto(s) movidos a CAJETIN1 ✔")

    # ── 19. Purge ─────────────────────────────────────────────────────────
    log_fn("[19] Purgando layers y bloques sin usar...")
    send_cmd(doc, "-PURGE\nAll\n*\nN\n")
    time.sleep(1.0)

    # ── Guardar si se indicó ruta ─────────────────────────────────────────
    if ruta_salida:
        abs_salida = os.path.abspath(ruta_salida)
        carpeta_salida = os.path.dirname(abs_salida)
        if carpeta_salida:
            os.makedirs(carpeta_salida, exist_ok=True)
        time.sleep(0.5)
        doc.SaveAs(abs_salida)
        log_fn(f"  Guardado en: {abs_salida} ✔")
    else:
        send_cmd(doc, "QSAVE\n")
        time.sleep(0.5)

    log_fn("=== Arte completado ✔ ===")
    return n_piezas


def main():
    pythoncom.CoInitialize()
    try:
        try:
            acad = win32com.client.GetActiveObject("AutoCAD.Application")
        except Exception:
            alerta_stop("AGP Arte Maker", "AutoCAD no está abierto.\nAbre AutoCAD con el plano y vuelve a intentarlo.")
            return
        doc = acad.ActiveDocument
        nombre = doc.Name

        # Confirmar con el usuario qué archivo va a procesar
        respuesta = ctypes.windll.user32.MessageBoxW(
            0,
            f"Se va a crear el arte sobre:\n\n  {nombre}\n\n"
            f"¿Es el plano correcto?\n\n"
            f"⚠ No cambies de pestaña en AutoCAD mientras corre el proceso.",
            "AGP Arte Maker — Confirmar",
            0x24   # MB_YESNO | MB_ICONQUESTION
        )
        if respuesta != 6:   # IDYES
            return

        doc.Activate()
        pipeline(doc)
        ctypes.windll.user32.MessageBoxW(0,
            "Arte creado correctamente.\nRevisa el resultado en AutoCAD.",
            "AGP Arte Maker", 0x40)
    except Exception as e:
        log(f"ERROR FATAL: {e}")
        alerta_stop("AGP Arte Maker — Error fatal", str(e))
    finally:
        pythoncom.CoUninitialize()


if __name__ == "__main__":
    main()
    input("\nPresiona Enter para cerrar...")
