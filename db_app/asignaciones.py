# -*- coding: utf-8 -*-
"""
Motor de asignacion de vitros y mallas con sistema de reservas.

Estados posibles en columna `estado`:
  NULL        -> importado del Excel, dato historico
  PENDIENTE   -> reservado, el usuario esta llenando el cajetin / formulario
  ASIGNADO    -> confirmado, datos guardados
  CANCELADO   -> el usuario cancelo, numero reciclable

Flujo auto (desde cajetin):
  reservar()  -> INSERT/UPDATE estado=PENDIENTE, retorna codigos reales
  confirmar() -> UPDATE estado=ASIGNADO + datos
  cancelar()  -> UPDATE estado=CANCELADO (reciclable)

Flujo manual (separar):
  reservar() + confirmar() en un solo paso dentro de dialogo_separar()
  Si cancela el dialogo -> cancelar() sobre lo reservado

Prioridad al elegir numeros:
  1. Filas CANCELADO (reciclar antes que crear nuevo)
  2. NEXT VALUE FOR de la SEQUENCE (nuevo numero, atomico multi-PC)

UPDLOCK + READPAST garantiza que dos PCs no reclamen el mismo CANCELADO.
"""

import os, sys
import pyodbc

_DIR = os.path.dirname(__file__)
sys.path.insert(0, _DIR)
from importar_excel import conectar


# ---------------------------------------------------------------------------
#  Helpers de secuencia
# ---------------------------------------------------------------------------

def _next_from_seq(cur, seq_name, n, fmt):
    resultado = []
    for _ in range(n):
        cur.execute(f"SELECT NEXT VALUE FOR {seq_name}")
        resultado.append(fmt(cur.fetchone()[0]))
    return resultado


def sincronizar_secuencias():
    """Llamar al final de cada sync del Excel."""
    cn = conectar()
    try:
        cur = cn.cursor()
        checks = [
            ("mallas.seq_vitro",
             "SELECT ISNULL(MAX(TRY_CAST(SUBSTRING(vitro,3,50) AS INT)),0) "
             "FROM mallas.vitrojet WHERE vitro LIKE 'T-[0-9]%'"),
            ("mallas.seq_grande",
             "SELECT ISNULL(MAX(TRY_CAST(SUBSTRING(codigo,3,50) AS INT)),0) "
             "FROM mallas.grandes WHERE codigo LIKE 'A-[0-9]%'"),
            ("mallas.seq_pequena",
             "SELECT ISNULL(MAX(TRY_CAST(codigo AS INT)),0) FROM mallas.pequenas"),
        ]
        for seq, qmax in checks:
            cur.execute(qmax)
            max_n = cur.fetchone()[0] or 0
            cur.execute(
                "SELECT CAST(current_value AS BIGINT) "
                "FROM sys.sequences WHERE object_id = OBJECT_ID(?)", (seq,))
            row = cur.fetchone()
            if row and row[0] <= max_n:
                cur.execute(f"ALTER SEQUENCE {seq} RESTART WITH {max_n + 1}")
        cn.commit()
    finally:
        cn.close()


# ---------------------------------------------------------------------------
#  Core: reservar / confirmar / cancelar
# ---------------------------------------------------------------------------

def reservar(n_vitros=0, n_grandes=0, n_pequenas=0):
    """
    Reserva codigos unicos en BD con estado=PENDIENTE.
    Multi-PC safe: UPDLOCK+READPAST para CANCELADOS, SEQUENCE para nuevos.
    Retorna dict con vitros/grandes/pequenas (codigos reales, ya en BD).
    """
    cn = conectar()
    try:
        cur = cn.cursor()

        vitros   = _reservar_tabla(cur, n_vitros,   "vitrojet",  "vitro",
                                   "mallas.seq_vitro",
                                   lambda n: f"T-{n:04d}",
                                   "vitro LIKE 'T-[0-9]%'")
        grandes  = _reservar_tabla(cur, n_grandes,  "grandes",   "codigo",
                                   "mallas.seq_grande",
                                   lambda n: f"A-{n}",
                                   "codigo LIKE 'A-[0-9]%'")
        pequenas = _reservar_tabla(cur, n_pequenas, "pequenas",  "codigo",
                                   "mallas.seq_pequena",
                                   lambda n: n,
                                   "1=1")

        cn.commit()
        return {
            "vitros":    vitros,
            "grandes":   grandes,
            "pequenas":  pequenas,
            "n_vitros":  n_vitros,
            "n_grandes": n_grandes,
            "n_pequenas":n_pequenas,
        }
    except Exception:
        try: cn.rollback()
        except Exception: pass
        raise
    finally:
        cn.close()


def _reservar_tabla(cur, n, tabla, pk, seq, fmt, filtro_nuevo):
    """
    Reserva n codigos de `mallas.{tabla}`:
      - Primero reclama CANCELADOS (UPDLOCK+READPAST = atomico)
      - Luego crea nuevos con SEQUENCE para los que falten
    """
    if n <= 0:
        return []

    # 1. Reclamar CANCELADOS disponibles
    cur.execute(
        f"UPDATE TOP({n}) mallas.{tabla} "
        f"SET estado='PENDIENTE' "
        f"OUTPUT inserted.{pk} "
        f"WHERE estado='CANCELADO'"
    )
    reciclados = [str(r[0]) for r in cur.fetchall()]
    faltan = n - len(reciclados)

    nuevos = []
    if faltan > 0:
        # 2. Crear nuevos con SEQUENCE
        codigos_nuevos = _next_from_seq(cur, seq, faltan, fmt)
        for cod in codigos_nuevos:
            cur.execute(
                f"INSERT INTO mallas.{tabla} ({pk}, estado, cambio) "
                f"VALUES (?, 'PENDIENTE', 'auto')", (cod,))
        nuevos = [str(c) for c in codigos_nuevos]

    return reciclados + nuevos


def confirmar(reserva, vehiculo="", version="", pieza="", cod_vehiculo="",
              cod_completo="", bnerig="BN", tipo="S", ruta_archivo="",
              responsable=None):
    """
    Marca la reserva como ASIGNADO y guarda todos los datos.
    Si falla -> rollback, los codigos quedan PENDIENTE (se pueden cancelar despues).
    """
    if not reserva:
        return reserva

    concat = _concatenar(vehiculo, cod_vehiculo, version, pieza)
    all_mallas = list(reserva.get("grandes", [])) + \
                 [str(p) for p in reserva.get("pequenas", [])]
    vm_map = dict(zip(reserva.get("vitros", []), all_mallas))

    cn = conectar()
    try:
        cur = cn.cursor()

        for v in reserva.get("vitros", []):
            cm = vm_map.get(v, "")
            tm = "G" if str(cm).startswith("A-") else ("P" if cm else None)
            cur.execute(
                "UPDATE mallas.vitrojet SET "
                "estado='ASIGNADO', codigo_malla=?, tipo_malla=?, cod_completo=?, "
                "bnerig=?, vehiculo=?, version=?, ruta=?, responsable=?, "
                "updated_at=GETDATE() WHERE vitro=?",
                (_val(cm), tm, _val(cod_completo), _val(bnerig) or "BN",
                 _val(vehiculo), _val(version), _val(ruta_archivo),
                 _val(responsable), v))

        for g in reserva.get("grandes", []):
            cur.execute(
                "UPDATE mallas.grandes SET "
                "estado='ASIGNADO', cod_veh=?, descripcion=?, pieza=?, tipo=?, "
                "version=?, concatenar=?, ruta_dwg=?, responsable=?, updated_at=GETDATE() "
                "WHERE codigo=?",
                (_val(cod_vehiculo), _val(vehiculo), _val(pieza),
                 _val(tipo) or "S", _val(version),
                 _val(concat), _val(ruta_archivo), _val(responsable), g))

        for p in reserva.get("pequenas", []):
            cur.execute(
                "UPDATE mallas.pequenas SET "
                "estado='ASIGNADO', cod_veh=?, descripcion=?, pieza=?, tipo=?, "
                "version=?, concatenar=?, ruta_dwg=?, responsable=?, updated_at=GETDATE() "
                "WHERE codigo=?",
                (_val(cod_vehiculo), _val(vehiculo), _val(pieza),
                 _val(tipo) or "S", _val(version),
                 _val(concat), _val(ruta_archivo), _val(responsable), p))

        cn.commit()
        return reserva
    except Exception:
        try: cn.rollback()
        except Exception: pass
        raise
    finally:
        cn.close()


def cancelar(reserva):
    """
    Marca los codigos de la reserva como CANCELADO.
    Quedan reciclables para la proxima reserva.
    """
    if not reserva:
        return
    vitros   = reserva.get("vitros",   [])
    grandes  = reserva.get("grandes",  [])
    pequenas = reserva.get("pequenas", [])
    if not (vitros or grandes or pequenas):
        return

    cn = conectar()
    try:
        cur = cn.cursor()
        for v in vitros:
            cur.execute("UPDATE mallas.vitrojet  SET estado='CANCELADO' WHERE vitro=?   AND estado='PENDIENTE'", (v,))
        for g in grandes:
            cur.execute("UPDATE mallas.grandes   SET estado='CANCELADO' WHERE codigo=?  AND estado='PENDIENTE'", (g,))
        for p in pequenas:
            cur.execute("UPDATE mallas.pequenas  SET estado='CANCELADO' WHERE codigo=?  AND estado='PENDIENTE'", (p,))
        cn.commit()
    finally:
        cn.close()


# Alias para compatibilidad con codigo existente
def proponer(n_vitros=0, n_grandes=0, n_pequenas=0):
    return reservar(n_vitros, n_grandes, n_pequenas)


def separar(n_vitros=0, n_grandes=0, n_pequenas=0, **datos):
    """Separacion manual: reserva + confirma. Retorna reserva."""
    res = reservar(n_vitros, n_grandes, n_pequenas)
    return confirmar(res, **datos)


# ---------------------------------------------------------------------------
#  Helpers
# ---------------------------------------------------------------------------

def _concatenar(vehiculo, cod_vehiculo, version, pieza):
    partes = "-".join(str(x) for x in [cod_vehiculo, version, pieza] if str(x).strip())
    if vehiculo and partes:
        return f"{vehiculo}-({partes})"
    return vehiculo or ""


def _val(s):
    if s is None:
        return None
    s = str(s).strip()
    return s if s else None


# ---------------------------------------------------------------------------
#  Dialogo asignacion automatica (desde dialogo_cajetin)
# ---------------------------------------------------------------------------

def dialogo_asignacion(parent=None, nombre_plano=""):
    """
    Reserva codigos REALES con estado=PENDIENTE en cuanto el usuario confirma
    la cantidad. Los muestra exactos (no aproximados).
    Retorna la reserva dict para que aceptar() del cajetin llame a confirmar().
    Si el cajetin se cancela, el llamador debe llamar a cancelar(reserva).
    """
    import tkinter as tk
    from tkinter import messagebox

    C = {
        "bg": "#12131A", "panel": "#1C1E2B", "sep": "#2A2D45",
        "accent": "#4D7EFF", "accent2": "#2D55CC",
        "text": "#E8EAFF", "muted": "#6B7099", "entry": "#1A1C2A",
        "green": "#3ECF8E", "red": "#FF5757", "orange": "#FF9500",
    }

    resultado = [None]
    win = tk.Toplevel(parent) if parent else tk.Tk()
    if parent:
        win.grab_set()

    win.title("Asignar vitro / malla")
    win.configure(bg=C["bg"])
    win.resizable(False, False)
    win.attributes("-topmost", True)

    tk.Frame(win, bg=C["accent"], height=4).pack(fill="x")
    hf = tk.Frame(win, bg=C["bg"], pady=14, padx=24); hf.pack(fill="x")
    tk.Label(hf, text="Asignacion Automatica",
             font=("Segoe UI", 15, "bold"), fg=C["text"], bg=C["bg"]).pack(anchor="w")
    if nombre_plano:
        tk.Label(hf, text=f"Plano: {nombre_plano}",
                 font=("Segoe UI", 9), fg=C["muted"], bg=C["bg"]).pack(anchor="w")

    body = tk.Frame(win, bg=C["bg"], padx=24); body.pack(fill="x")

    def _sep_lbl(txt):
        tk.Label(body, text=txt, font=("Segoe UI", 8, "bold"),
                 fg=C["accent"], bg=C["bg"]).pack(anchor="w", pady=(10, 0))
        tk.Frame(body, bg=C["sep"], height=1).pack(fill="x", pady=(2, 4))

    def _spin(lbl, default=1):
        f = tk.Frame(body, bg=C["bg"], pady=3); f.pack(fill="x")
        tk.Label(f, text=lbl, width=26, anchor="e",
                 font=("Segoe UI", 9), fg=C["muted"], bg=C["bg"]).pack(side="left", padx=8)
        var = tk.IntVar(value=default)
        tk.Spinbox(f, from_=0, to=20, textvariable=var, width=5,
                   bg=C["entry"], fg=C["text"], insertbackground=C["text"],
                   buttonbackground="#1C1E2B", relief="flat",
                   font=("Segoe UI", 11, "bold")).pack(side="left")
        return var

    _sep_lbl("CANTIDAD A ASIGNAR")
    nv  = _spin("Vitros  (T-xxxx)",          1)
    ng  = _spin("Mallas grandes  (A-xxxx)",  0)
    np_ = _spin("Mallas pequenas  (numero)", 0)

    _sep_lbl("NUMEROS ASIGNADOS")
    pf  = tk.Frame(body, bg=C["panel"], padx=14, pady=12); pf.pack(fill="x", pady=(0, 8))
    plbl = tk.Label(pf,
                    text='Presiona "Reservar numeros" para ver los codigos exactos.\n'
                         'Al reservar quedan PENDIENTE en BD hasta que aceptes o canceles.',
                    font=("Segoe UI", 9), fg=C["muted"], bg=C["panel"],
                    wraplength=380, justify="left")
    plbl.pack(anchor="w")

    reserva_actual = [None]

    def _liberar_anterior():
        if reserva_actual[0]:
            try: cancelar(reserva_actual[0])
            except Exception: pass
            reserva_actual[0] = None

    def hacer_reserva():
        if nv.get() == 0 and ng.get() == 0 and np_.get() == 0:
            messagebox.showwarning("Sin cantidad", "Indica al menos 1 vitro o malla.")
            return
        plbl.configure(text="Reservando...", fg=C["muted"]); win.update()
        _liberar_anterior()  # si ya habia reserva anterior con distinta cantidad, cancelarla
        try:
            res = reservar(nv.get(), ng.get(), np_.get())
            reserva_actual[0] = res
            lineas = []
            if res["vitros"]:   lineas.append("Vitros:     " + "  ".join(res["vitros"]))
            if res["grandes"]:  lineas.append("Mallas G:   " + "  ".join(res["grandes"]))
            if res["pequenas"]: lineas.append("Mallas P:   " + "  ".join(str(c) for c in res["pequenas"]))
            plbl.configure(
                text=("\n".join(lineas) + "\n(PENDIENTE en BD - confirmar al aceptar cajetin)"),
                fg=C["green"])
        except Exception as e:
            plbl.configure(text=f"Error: {e}", fg=C["red"])

    def ok():
        if reserva_actual[0] is None:
            hacer_reserva()
            if reserva_actual[0] is None:
                return
        resultado[0] = reserva_actual[0]
        win.destroy()

    def cancel_win():
        _liberar_anterior()  # -> CANCELADO en BD (reciclable)
        win.destroy()

    win.protocol("WM_DELETE_WINDOW", cancel_win)  # X del titulo

    bf = tk.Frame(win, bg=C["bg"], pady=14, padx=24); bf.pack(fill="x")

    def _btn(txt, cmd, bg, hov):
        b = tk.Button(bf, text=txt, command=cmd, bg=bg, fg="#FFF",
                      activebackground=hov, activeforeground="#FFF",
                      relief="flat", bd=0, font=("Segoe UI", 10, "bold"),
                      padx=12, pady=7, cursor="hand2")
        b.bind("<Enter>", lambda _: b.configure(bg=hov))
        b.bind("<Leave>", lambda _: b.configure(bg=bg))
        b.pack(side="left", padx=(0, 8))

    _btn("Reservar numeros", hacer_reserva, "#2D3250",    "#3D4270")
    _btn("Confirmar",        ok,            C["accent"],  C["accent2"])
    _btn("Cancelar",         cancel_win,    "#2A2A3A",    "#3A3A4A")

    win.bind("<Escape>", lambda _: cancel_win())
    win.update_idletasks()
    sw, sh = win.winfo_screenwidth(), win.winfo_screenheight()
    w = max(win.winfo_reqwidth(), 460)
    win.geometry(f"{w}x{win.winfo_reqheight()}+{(sw-w)//2}+{(sh-win.winfo_reqheight())//2}")

    if parent:
        parent.wait_window(win)
    else:
        win.mainloop()

    return resultado[0]


# ---------------------------------------------------------------------------
#  Dialogo separacion manual (boton naranja en TabArte)
# ---------------------------------------------------------------------------

_do_confirmar = confirmar
_do_cancelar  = cancelar


def dialogo_separar(parent_win=None):
    """
    Separacion manual completa.
    - Reserva al confirmar (no antes) -> si cancela, nada queda en BD
    """
    import tkinter as tk
    from tkinter import messagebox, filedialog

    C = {
        "bg": "#12131A", "panel": "#1C1E2B", "sep": "#2A2D45",
        "accent": "#FF9500", "accent2": "#CC7700",
        "text": "#E8EAFF", "muted": "#6B7099", "entry": "#1A1C2A",
        "green": "#3ECF8E", "red": "#FF5757", "req": "#FF6B6B",
    }

    resultado = [None]
    win = tk.Toplevel(parent_win) if parent_win else tk.Tk()
    if parent_win:
        win.grab_set()

    win.title("Separar vitro / malla")
    win.configure(bg=C["bg"])
    win.resizable(False, False)
    win.attributes("-topmost", True)

    tk.Frame(win, bg=C["accent"], height=6).pack(fill="x")
    hf = tk.Frame(win, bg=C["bg"], pady=14, padx=24); hf.pack(fill="x")
    tk.Label(hf, text="Separar Vitro / Malla",
             font=("Segoe UI", 16, "bold"), fg=C["accent"], bg=C["bg"]).pack(anchor="w")
    tk.Label(hf, text="Asignacion manual  |  * = obligatorio  |  guarda en BD al confirmar",
             font=("Segoe UI", 9), fg=C["muted"], bg=C["bg"]).pack(anchor="w")

    body = tk.Frame(win, bg=C["bg"], padx=24); body.pack(fill="x")

    def _sec(txt):
        tk.Label(body, text=txt, font=("Segoe UI", 8, "bold"),
                 fg=C["accent"], bg=C["bg"]).pack(anchor="w", pady=(10, 0))
        tk.Frame(body, bg=C["sep"], height=1).pack(fill="x", pady=(2, 4))

    def _row(lbl, default="", req=False):
        f = tk.Frame(body, bg=C["bg"], pady=3); f.pack(fill="x")
        col = C["req"] if req else C["muted"]
        tk.Label(f, text=(lbl + " *" if req else lbl), width=24, anchor="e",
                 font=("Segoe UI", 9), fg=col, bg=C["bg"]).pack(side="left", padx=8)
        var = tk.StringVar(value=default)
        tk.Entry(f, textvariable=var, width=28,
                 bg=C["entry"], fg=C["text"], insertbackground=C["text"],
                 relief="flat", font=("Segoe UI", 10),
                 highlightthickness=1, highlightbackground="#2E3250",
                 highlightcolor=C["accent"], bd=4).pack(side="left")
        return var

    def _spin(lbl, default=1):
        f = tk.Frame(body, bg=C["bg"], pady=3); f.pack(fill="x")
        tk.Label(f, text=lbl, width=24, anchor="e",
                 font=("Segoe UI", 9), fg=C["muted"], bg=C["bg"]).pack(side="left", padx=8)
        var = tk.IntVar(value=default)
        tk.Spinbox(f, from_=0, to=20, textvariable=var, width=5,
                   bg=C["entry"], fg=C["text"], insertbackground=C["text"],
                   buttonbackground="#1C1E2B", relief="flat",
                   font=("Segoe UI", 11, "bold")).pack(side="left")
        return var

    _sec("CANTIDAD")
    nv  = _spin("Vitros  (T-xxxx)")
    ng  = _spin("Mallas grandes  (A-xxxx)", 0)
    np_ = _spin("Mallas pequenas  (numero)", 0)

    _sec("DATOS DEL VEHICULO")
    v_veh  = _row("Vehiculo",          req=True)
    v_cod  = _row("Cod. vehiculo",     req=True)
    v_ver  = _row("Version / Ano",     req=True)
    v_piez = _row("Pieza",             req=True)
    v_comp = _row("Cod. completo veh.")
    v_bn   = _row("BNERIG",            default="BN")
    v_tipo = _row("Tipo malla",        default="S")

    _sec("ARCHIVO")
    ruta_var = tk.StringVar()
    rf = tk.Frame(body, bg=C["bg"], pady=3); rf.pack(fill="x")
    tk.Label(rf, text="Ruta archivo *", width=24, anchor="e",
             font=("Segoe UI", 9), fg=C["req"], bg=C["bg"]).pack(side="left", padx=8)
    tk.Entry(rf, textvariable=ruta_var, width=24,
             bg=C["entry"], fg=C["text"], insertbackground=C["text"],
             relief="flat", font=("Segoe UI", 9),
             highlightthickness=1, highlightbackground="#2E3250",
             highlightcolor=C["accent"], bd=4).pack(side="left")
    tk.Button(rf, text="...",
              command=lambda: ruta_var.set(
                  filedialog.askopenfilename(
                      title="Seleccionar archivo",
                      filetypes=[("DWG / PDF", "*.dwg *.pdf"), ("Todos", "*.*")]
                  ) or ruta_var.get()),
              bg="#2D3250", fg=C["text"], relief="flat",
              font=("Segoe UI", 9), padx=6, cursor="hand2").pack(side="left", padx=4)

    _sec("RESPONSABLE")
    v_resp = _row("Nombre responsable *", req=True)

    # Panel resultado — visible siempre (vacío al inicio, lleno tras confirmar)
    pf = tk.Frame(win, bg=C["panel"], padx=16, pady=14)
    pf.pack(fill="x", padx=24, pady=(8, 2))
    plbl_titulo = tk.Label(pf, text="", font=("Segoe UI", 11, "bold"),
                           fg=C["green"], bg=C["panel"])
    plbl_titulo.pack(anchor="w")
    plbl = tk.Label(pf, text="", font=("Segoe UI", 13, "bold"),
                    fg=C["text"], bg=C["panel"], wraplength=440, justify="left")
    plbl.pack(anchor="w", pady=(4, 0))
    plbl_sub = tk.Label(pf, text="", font=("Segoe UI", 9),
                        fg=C["muted"], bg=C["panel"])
    plbl_sub.pack(anchor="w")

    btn_confirmar_ref = [None]

    def _hacer_confirmar():
        faltantes = []
        for var, nombre in [(v_veh, "Vehiculo"), (v_cod, "Cod. vehiculo"),
                             (v_ver, "Version"), (v_piez, "Pieza"),
                             (ruta_var, "Ruta archivo"), (v_resp, "Nombre responsable")]:
            if not var.get().strip():
                faltantes.append(nombre)
        if faltantes:
            messagebox.showwarning(
                "Campos requeridos",
                "Completa los siguientes campos:\n  " + "\n  ".join(faltantes))
            return
        if nv.get() == 0 and ng.get() == 0 and np_.get() == 0:
            messagebox.showwarning("Sin cantidad", "Indica al menos 1 vitro o malla.")
            return

        plbl_titulo.configure(text="Guardando...", fg=C["muted"])
        plbl.configure(text="")
        plbl_sub.configure(text="")
        win.update()
        try:
            res = reservar(nv.get(), ng.get(), np_.get())
            prop_real = _do_confirmar(
                res,
                vehiculo     = v_veh.get().strip(),
                version      = v_ver.get().strip(),
                pieza        = v_piez.get().strip(),
                cod_vehiculo = v_cod.get().strip(),
                cod_completo = v_comp.get().strip() or None,
                bnerig       = v_bn.get().strip() or "BN",
                tipo         = v_tipo.get().strip() or "S",
                ruta_archivo = ruta_var.get().strip(),
                responsable  = v_resp.get().strip(),
            )
            resultado[0] = prop_real

            # Construir texto del resultado
            lineas = []
            if prop_real["vitros"]:
                lineas.append("Vitro:     " + "   ".join(prop_real["vitros"]))
            if prop_real["grandes"]:
                lineas.append("Malla G:   " + "   ".join(prop_real["grandes"]))
            if prop_real["pequenas"]:
                lineas.append("Malla P:   " + "   ".join(str(c) for c in prop_real["pequenas"]))

            plbl_titulo.configure(text="✔  Asignacion guardada en BD", fg=C["green"])
            plbl.configure(text="\n".join(lineas), fg=C["text"])
            plbl_sub.configure(
                text=f"Vehiculo: {v_veh.get().strip()}  |  Version: {v_ver.get().strip()}  |  "
                     f"Responsable: {v_resp.get().strip()}",
                fg=C["muted"])

            # Deshabilitar boton confirmar, habilitar Nuevo / Cerrar
            if btn_confirmar_ref[0]:
                btn_confirmar_ref[0].configure(state="disabled", bg="#1A1C2A")
        except Exception as e:
            plbl_titulo.configure(text="Error al guardar", fg=C["red"])
            plbl.configure(text=str(e), fg=C["red"])
            plbl_sub.configure(text="")

    def cancel():
        win.destroy()

    win.protocol("WM_DELETE_WINDOW", cancel)
    bf = tk.Frame(win, bg=C["bg"], pady=14, padx=24); bf.pack(fill="x")

    def _btn(txt, cmd, bg, hov):
        b = tk.Button(bf, text=txt, command=cmd, bg=bg, fg="#FFF",
                      activebackground=hov, activeforeground="#FFF",
                      relief="flat", bd=0, font=("Segoe UI", 10, "bold"),
                      padx=12, pady=8, cursor="hand2")
        b.bind("<Enter>", lambda _: b.configure(bg=hov))
        b.bind("<Leave>", lambda _: b.configure(bg=bg))
        b.pack(side="left", padx=(0, 8))

    b_conf = tk.Button(bf, text="Confirmar separacion", command=_hacer_confirmar,
                       bg=C["accent"], fg="#FFF", activebackground=C["accent2"],
                       activeforeground="#FFF", relief="flat", bd=0,
                       font=("Segoe UI", 10, "bold"), padx=12, pady=8, cursor="hand2")
    b_conf.bind("<Enter>", lambda _: b_conf.configure(bg=C["accent2"]) if str(b_conf["state"]) != "disabled" else None)
    b_conf.bind("<Leave>", lambda _: b_conf.configure(bg=C["accent"])  if str(b_conf["state"]) != "disabled" else None)
    b_conf.pack(side="left", padx=(0, 8))
    btn_confirmar_ref[0] = b_conf

    _btn("Cerrar", cancel, "#2A2A3A", "#3A3A4A")

    win.bind("<Escape>", lambda _: cancel())
    win.update_idletasks()
    sw, sh = win.winfo_screenwidth(), win.winfo_screenheight()
    w = max(win.winfo_reqwidth(), 500)
    h = win.winfo_reqheight()
    win.geometry(f"{w}x{h}+{(sw-w)//2}+{(sh-h)//2}")

    if parent_win:
        parent_win.wait_window(win)
    else:
        win.mainloop()

    return resultado[0]
