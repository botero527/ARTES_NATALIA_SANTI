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
DEGRADE_INVERTIR = True   # True = REVERSE antes del DIVIDE (bolas grandes hacia BN)

PAT_PERIM = ["PERIMETRO"]
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


def ents_por_patron(msp, patrones):
    res = []
    for e in msp:
        try:
            if any(p in e.Layer.upper() for p in patrones):
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


def asegurar_layer(doc, nombre):
    try:
        doc.Layers.Item(nombre)
    except Exception:
        doc.Layers.Add(nombre)


def alerta(titulo, msg):
    ctypes.windll.user32.MessageBoxW(0, msg, titulo, 0x30)


def alerta_stop(titulo, msg):
    ctypes.windll.user32.MessageBoxW(0, msg, titulo, 0x10)


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
    for d in [dist, -dist]:
        try:
            results = list(ent.Offset(d))
            for r in results:
                a = area_bbox(r)
                if a < area_orig and a < mejor_area:
                    mejor_area = a
                    if mejor and mejor != r:
                        try: mejor.Delete()
                        except Exception: pass
                    mejor = r
                else:
                    try: r.Delete()
                    except Exception: pass
        except Exception:
            pass
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

def dialogo_cajetin(nombre_plano=""):
    """
    Muestra el diálogo de campos del cajetín.
    Retorna dict {campo: valor} o None si canceló.
    """
    import tkinter as tk
    from tkinter import ttk
    import re as _r

    hoy = datetime.date.today().strftime("%d.%m.%Y")

    FILAS = [
        ("DIBUJO",    "Dibujo",           None),
        ("VEHICULO",  "Vehículo",          None),
        ("MODELO",    "Modelo (año)",      "auto"),
        ("COD PLANO", "Cód. plano",        None),
        ("NAGS",      "NAGS",              "auto"),
        ("VERSION",   "Versión",           "auto"),
        ("PIEZA",     "Pieza",             "auto"),
        ("VITRO",     "Vitro",             None),
        ("MALLA",     "Malla",             None),
        ("FECHA",     "Fecha",             hoy),
        ("REVISADO",  "Revisado",          "Santiago P."),
        ("MEDIDAS",   "Medidas",           "Milimetros"),
        ("VISTA",     "Vista",             "Interna"),
        ("ESCALA",    "Escala",            "1:1"),
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

    root = tk.Tk()
    root.title("Rellenar Cajetín — AGP Arte Maker")
    root.resizable(False, False)
    root.attributes("-topmost", True)

    frame = ttk.Frame(root, padding=16)
    frame.grid(row=0, column=0, sticky="nsew")

    entries = {}
    row_w = 0
    sep_puesto = False

    for campo, etiqueta, default in FILAS:
        if campo == "FECHA" and not sep_puesto:
            ttk.Separator(frame, orient="horizontal").grid(
                row=row_w, column=0, columnspan=2, sticky="ew", pady=(8, 4))
            ttk.Label(frame, text="— valores por defecto (editables) —",
                      foreground="gray").grid(row=row_w+1, column=0,
                      columnspan=2, pady=(0, 6))
            row_w += 2
            sep_puesto = True

        ttk.Label(frame, text=etiqueta + ":", anchor="e", width=18).grid(
            row=row_w, column=0, sticky="e", pady=3, padx=(0, 8))
        ent = ttk.Entry(frame, width=36)
        ent.grid(row=row_w, column=1, sticky="w", pady=3)

        if isinstance(default, str) and default not in ("auto",):
            ent.insert(0, default)
            ent.configure(foreground="gray")

        if default == "auto":
            ent.configure(foreground="gray")

        entries[campo] = ent
        row_w += 1

    def on_cod_plano(*_):
        n, v, p = parsear_cod(entries["COD PLANO"].get())
        for k, val in [("NAGS", n), ("VERSION", v), ("PIEZA", p)]:
            entries[k].delete(0, tk.END)
            entries[k].insert(0, val)
            entries[k].configure(foreground="black")

    def on_vehiculo(*_):
        anio = extraer_anio(entries["VEHICULO"].get())
        if anio:
            entries["MODELO"].delete(0, tk.END)
            entries["MODELO"].insert(0, anio)
            entries["MODELO"].configure(foreground="black")

    entries["COD PLANO"].bind("<FocusOut>", on_cod_plano)
    entries["COD PLANO"].bind("<Return>",   on_cod_plano)
    entries["VEHICULO"].bind("<FocusOut>",  on_vehiculo)
    entries["VEHICULO"].bind("<Return>",    on_vehiculo)

    # Pre-llenar COD PLANO con el nombre del plano
    if nombre_plano:
        entries["COD PLANO"].delete(0, tk.END)
        entries["COD PLANO"].insert(0, nombre_plano)
        on_cod_plano()

    list(entries.values())[0].focus_set()

    def aceptar(*_):
        resultado[0] = {c: e.get().strip() for c, e in entries.items()}
        root.destroy()

    def cancelar(*_):
        root.destroy()

    root.bind("<Escape>", cancelar)

    btn_frame = ttk.Frame(frame)
    btn_frame.grid(row=row_w, column=0, columnspan=2, pady=(14, 0))
    ttk.Button(btn_frame, text="  Aceptar  ", command=aceptar).pack(side="left", padx=8)
    ttk.Button(btn_frame, text="  Cancelar ", command=cancelar).pack(side="left", padx=8)

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

def main():
    pythoncom.CoInitialize()
    try:
        try:
            acad = win32com.client.GetActiveObject("AutoCAD.Application")
        except Exception:
            alerta_stop("AGP Arte Maker", "AutoCAD no está abierto.\nAbre AutoCAD con el plano y vuelve a intentarlo.")
            return

        doc = acad.ActiveDocument
        msp = doc.ModelSpace
        nombre_plano = os.path.splitext(doc.Name)[0]
        log(f"\n=== AGP Arte Maker AutoCAD: {doc.Name} ===")

        # ── 1. Quitar GlobalWidth ──────────────────────────────────────────────
        log("[1] Limpiando GlobalWidth a 0...")
        for e in msp:
            try:
                if "Polyline" in e.ObjectName:
                    e.ConstantWidth = 0.0
                    e.Update()
            except Exception:
                pass

        # ── 2. Verificar contornos cerrados ───────────────────────────────────
        log("[2] Verificando contornos cerrados...")
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
            log(f"ERROR: {msg}")
            alerta_stop("AGP Arte Maker — Error", msg)
            return
        log("  Todos los contornos cerrados ✔")

        # ── 3. Radios mínimos ─────────────────────────────────────────────────
        log(f"[3] Verificando radios (mín {RADIO_MIN} mm)...")
        radios_malos = []
        for e in ents_por_patron(msp, PAT_PERIM):
            radios_malos += verificar_radios(e, RADIO_MIN)
        if radios_malos:
            r_min = min(radios_malos)
            log(f"  WARN radios: {sorted(set(radios_malos))[:8]}")
            alerta("AGP Arte Maker — Advertencia radios",
                   f"Se detectaron radios menores a {RADIO_MIN} mm en el perímetro.\n"
                   f"Radio mínimo encontrado: {r_min:.3f} mm\n\nEl proceso continuará.")
        else:
            log("  Radios OK ✔")

        # ── 4. Detectar degradé ───────────────────────────────────────────────
        bn_ents = sorted(
            [e for e in ents_por_patron(msp, PAT_BN)
             if e.Closed and "Polyline" in e.ObjectName],
            key=area_bbox, reverse=True)
        CON_DEGRADE = len(bn_ents) >= 2
        bn_ent = bn_ents[0] if bn_ents else None
        log(f"[4] BN encontrados: {len(bn_ents)} → {'CON degradé' if CON_DEGRADE else 'SIN degradé'}")

        # ── 5. Perímetro ──────────────────────────────────────────────────────
        perim_ents = sorted(
            [e for e in ents_por_patron(msp, PAT_PERIM)
             if e.Closed and "Polyline" in e.ObjectName],
            key=area_bbox, reverse=True)
        if not perim_ents:
            alerta_stop("AGP Arte Maker — Error", "No se encontró curva PERIMETRO cerrada.")
            return
        perim_ent = perim_ents[0]
        log("  Perímetro encontrado ✔")

        # ── Capturar handles del logo ANTES del import (para refetchear después) ─
        _logo_handles = set()
        for _e in ents_por_patron(msp, PAT_LOGO):
            try: _logo_handles.add(_e.Handle)
            except Exception: pass
        log(f"  Logo en plano: {len(_logo_handles)} objeto(s)")

        # ── Registrar handles existentes ANTES del import ─────────────────────
        handles_antes = handles_actuales(msp)

        # ── 6. Importar cajetines vía COM InsertBlock (síncrono) ──────────────
        log("[6] Importando cajetines...")
        abs_caj = os.path.abspath(CAJETIN_DWG)
        if not os.path.isfile(abs_caj):
            log(f"  WARN: no se encontró {abs_caj}")
        else:
            try:
                pt_ins = win32com.client.VARIANT(
                    pythoncom.VT_ARRAY | pythoncom.VT_R8, [0.0, 0.0, 0.0])
                blk_ref = msp.InsertBlock(pt_ins, abs_caj, 1.0, 1.0, 1.0, 0.0)
                log("  Bloque insertado vía COM ✔")

                # Nivel 1: explotar bloque raíz del DWG
                nivel1 = []
                try:
                    nivel1 = list(blk_ref.Explode())
                    log(f"  Nivel 1 explosión: {len(nivel1)} objetos")
                except Exception as ex1:
                    log(f"  WARN explosión nivel 1: {ex1}")

                # Nivel 2: explotar bloques anidados (CAJETIN 1, CAJETIN 2…)
                n2 = 0
                for e2 in nivel1:
                    try:
                        if e2.ObjectName == "AcDbBlockReference":
                            e2.Explode()
                            n2 += 1
                    except Exception:
                        pass
                if n2:
                    log(f"  Nivel 2 explosión: {n2} bloques anidados ✔")
                time.sleep(0.5)
                # Refrescar referencia de msp — el COM object no se actualiza automáticamente
                msp = doc.ModelSpace

            except Exception as e_ins:
                log(f"  WARN InsertBlock COM falló ({e_ins}), usando SendCommand...")
                doc.SendCommand(f'-INSERT "{abs_caj}"\n0,0,0\n1\n1\n0\n')
                time.sleep(3)
                doc.SendCommand("EXPLODE\nL\n\n")
                time.sleep(2)
                msp = doc.ModelSpace  # refrescar

        # Identificar objetos nuevos y corregir colores ByBlock
        nuevos = objetos_nuevos(msp, handles_antes)
        log(f"  Objetos nuevos detectados: {len(nuevos)}")
        # Debug: mostrar primeros 8 layers para diagnóstico
        _layers_vistos = set()
        for _e in nuevos[:40]:
            try: _layers_vistos.add(_e.Layer)
            except Exception: pass
        if _layers_vistos:
            log(f"  Layers detectados: {sorted(_layers_vistos)[:8]}")
        corregir_colores_bylayer(nuevos)

        # ── 7. Buscar CAJETIN y LOGO1 directamente por layer (robusto) ───────
        # Buscar en todo msp — funciona aunque el handle-tracking falle
        logo1_ents   = [e for e in msp if "LOGO1"   in e.Layer.upper()]
        cajetin_ents = [e for e in msp if "CAJETIN" in e.Layer.upper()]
        log(f"  CAJETIN en msp: {len(cajetin_ents)} obj  |  LOGO1: {len(logo1_ents)} obj")

        # Borrar objetos nuevos que no son CAJETIN ni LOGO1 (AYUDAS, etc.)
        conservar_layers = {"CAJETIN", "LOGO"}
        borrados = 0
        for e in nuevos:
            try:
                ly = e.Layer.upper()
                if not any(pat in ly for pat in conservar_layers):
                    e.Delete()
                    borrados += 1
            except Exception:
                pass
        log(f"  Limpieza: {borrados} objeto(s) sobrantes eliminados ✔")

        # ── 8. Reemplazar logo ────────────────────────────────────────────────
        log("[8] Reemplazando logo...")
        # Refetchear logo del plano desde el msp fresco (por handle)
        logo_plano = [e for e in msp
                      if hasattr(e, 'Handle') and e.Handle in _logo_handles]
        logo1_ents = [e for e in msp if "LOGO1" in e.Layer.upper()]
        log(f"  logo_plano: {len(logo_plano)}  logo1: {len(logo1_ents)}")
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
                    log("  Logo reemplazado ✔")
            elif not logo1_ents:
                log("  WARN: LOGO1 no encontrado en el cajetín.")
            else:
                log("  WARN: no hay logo en el plano.")
        except Exception as e_logo:
            log(f"  WARN logo: {e_logo}")

        # ── 9. Centrar cajetín sobre la pieza ─────────────────────────────────
        log("[9] Centrando cajetín...")
        caj_ents = [e for e in msp if "CAJETIN" in e.Layer.upper()]
        log(f"  Buscando CAJETIN en msp: {len(caj_ents)} objetos")
        if caj_ents:
            cx_p, cy_p = centro_bbox([perim_ent])
            cx_c, cy_c = centro_bbox(caj_ents)
            log(f"  Centro pieza: ({cx_p:.1f},{cy_p:.1f})  Centro cajetín: ({cx_c:.1f},{cy_c:.1f})")
            if cx_p is not None and cx_c is not None:
                dx, dy = cx_p - cx_c, cy_p - cy_c
                for e in caj_ents:
                    try: e.Move(pt(0,0), pt(dx,dy))
                    except Exception: pass
                log("  Cajetín centrado ✔")
        else:
            log("  WARN: no se encontró ningún objeto con layer CAJETIN.")

        # ── 10. Crear layers de arte ──────────────────────────────────────────
        for lyr in [LAYER_PLANES, LAYER_K2, LAYER_K, LAYER_K3]:
            asegurar_layer(doc, lyr)

        # ── 11. Offset perímetro 0.5 ──────────────────────────────────────────
        log(f"[11] Offset perímetro {OFFSET_PERIM} mm...")
        off_perim = offset_inward(perim_ent, OFFSET_PERIM)
        if not off_perim:
            alerta_stop("AGP Arte Maker — Error", "No se pudo crear offset del perímetro.")
            return
        off_perim.Layer = LAYER_PLANES

        # ── 12. Hatch k2 (perímetro → offset 0.5) ────────────────────────────
        log("[12] Hatch k2...")
        hatch_solido(msp, doc, perim_ent, off_perim, LAYER_K2)

        # ── 13. Hatch k (BN → offset 0.5) ────────────────────────────────────
        log("[13] Hatch k...")
        if bn_ent:
            hatch_solido(msp, doc, bn_ent, off_perim, LAYER_K)
        else:
            log("  WARN: no se encontró banda negra.")

        # ── 14. Degradé ───────────────────────────────────────────────────────
        if CON_DEGRADE and bn_ent:
            log(f"[14] Degradé: offset BN {OFFSET_BN_DEG} mm...")
            off_bn = offset_inward(bn_ent, OFFSET_BN_DEG)
            if off_bn:
                off_bn.Layer = LAYER_PLANES
                longitud = float(off_bn.Length)
                n_pepas  = int(round(longitud / DIVISOR_DEG))
                log(f"  Longitud: {longitud:.2f} mm  pepas: {n_pepas}")

                if n_pepas > 0:
                    # Verificar que el bloque "25" esté definido
                    bloque_existe = False
                    try:
                        doc.Blocks.Item(BLOQUE_25)
                        bloque_existe = True
                        log(f"  Bloque '{BLOQUE_25}' ya existe.")
                    except Exception:
                        pass

                    if not bloque_existe:
                        log(f"  Importando bloque '{BLOQUE_25}' desde cajetines...")
                        doc.SendCommand(f'-INSERT "{abs_caj}"\n0,0,0\n1\n1\n0\n')
                        time.sleep(4)
                        doc.SendCommand("ERASE\nL\n\n")
                        time.sleep(1.5)
                        log(f"  Bloque '{BLOQUE_25}' registrado.")

                    # Orientar el degradé: bolas grandes hacia el BN
                    if DEGRADE_INVERTIR:
                        h_rev = off_bn.Handle
                        doc.SendCommand(f'REVERSE\n(handent "{h_rev}")\n\n')
                        time.sleep(0.5)
                        log("  REVERSE aplicado (DEGRADE_INVERTIR=True) ✔")

                    # Poner k3 como layer activo → DIVIDE inserta bloques en k3
                    doc.SendCommand(f'CLAYER\n{LAYER_K3}\n')
                    time.sleep(0.3)

                    handle = off_bn.Handle
                    doc.SendCommand(
                        f'DIVIDE\n'
                        f'(handent "{handle}")\n'
                        f'B\n'
                        f'{BLOQUE_25}\n'
                        f'Y\n'
                        f'{n_pepas}\n'
                    )
                    espera = max(4.0, n_pepas * 0.02)
                    log(f"  Esperando {espera:.1f}s (DIVIDE)...")
                    time.sleep(espera)
                    doc.SendCommand("CLAYER\n0\n")
                    time.sleep(0.3)
                    log("  Degradé en layer k3 ✔")
            else:
                log("  WARN: no se pudo crear offset interior de BN.")
        else:
            log("[14] Sin degradé — omitido.")

        # ── 15. Mover PERIMETRO/BN a PLANES ──────────────────────────────────
        log("[15] Moviendo geometría original a PLANES...")
        for e in ents_por_patron(msp, PAT_PERIM + PAT_BN):
            try: e.Layer = LAYER_PLANES
            except Exception: pass

        log("=== Arte base completado — revisa en AutoCAD antes de guardar ✔ ===")

        # ── 17. Diálogo del cajetín ───────────────────────────────────────────
        log("[17] Abriendo diálogo de cajetín...")
        valores = dialogo_cajetin(nombre_plano)

        if valores:
            actualizar_texto_cajetin(msp, valores)
            log("  Cajetín aplicado ✔ — guarda manualmente cuando estés conforme.")
        else:
            log("  Cajetín: cancelado por el usuario.")

        # ── 18. Purge — elimina layers vacíos (AYUDAS, etc.) ─────────────────
        log("[18] Purgando layers y bloques sin usar...")
        doc.SendCommand("-PURGE\nAll\n*\nN\n")
        time.sleep(2)

        log("=== Arte completado ✔ ===")
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
