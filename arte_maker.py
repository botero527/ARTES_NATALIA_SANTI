
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

SCRIPT_RHINO  = os.path.join(os.path.dirname(os.path.abspath(__file__)), "arte_script.py")
CAJETIN_DWG   = os.path.join(os.path.dirname(os.path.abspath(__file__)), "LAYERS Y CAJETINES 1.dwg")

# ── Parámetros arte (espejo de config.py) ─────────────────────────────────────
_OFFSET_PERIM   = 0.5
_OFFSET_BN_DEG  = 2.5
_DIVISOR_DEG    = 3
_BLOQUE_25      = "25"
_LAYER_PLANES   = "PLANES"
_LAYER_K2       = "k2"
_LAYER_K        = "k"
_PAT_PERIM      = ["PERIMETRO"]
_PAT_BN         = ["BANDA NEGRA", "BANDANEGRA", "BN", "PHANTOM", "BANDA"]
_PAT_LOGO       = ["LOGO", "TRAZABILIDAD"]
_RADIO_MIN      = 15.0

import math as _math2
import re   as _re


# ═══════════════════════════════════════════════════════════════════════════════
#  CREAR ARTE EN AUTOCAD
# ═══════════════════════════════════════════════════════════════════════════════

def _acad_connect():
    """Devuelve la instancia activa de AutoCAD o lanza RuntimeError."""
    pythoncom.CoInitialize()
    try:
        return win32com.client.GetActiveObject("AutoCAD.Application")
    except Exception:
        raise RuntimeError("AutoCAD no está abierto. Ábrelo primero.")


def _pt(x, y, z=0.0):
    return win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, [float(x), float(y), float(z)])


def _ents_por_patron(msp, patrones):
    """Entidades cuyo layer contenga alguno de los patrones."""
    res = []
    for ent in msp:
        try:
            if any(p in ent.Layer.upper() for p in patrones):
                res.append(ent)
        except Exception:
            pass
    return res


def _bbox(ent):
    lo, hi = ent.GetBoundingBox()
    return lo[0], lo[1], hi[0], hi[1]


def _area_bbox(ent):
    try:
        x0, y0, x1, y1 = _bbox(ent)
        return abs(x1-x0) * abs(y1-y0)
    except Exception:
        return 0.0


def _centro_bbox(ents):
    """Centro del bounding box de una lista de entidades."""
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


def _asegurar_layer(doc, nombre):
    try:
        doc.Layers.Item(nombre)
    except Exception:
        doc.Layers.Add(nombre)


def _verificar_radios_acad(ent, radio_min=15.0):
    """
    Verifica radios en una LWPolyline de AutoCAD usando el bulge.
    Bulge = tan(angulo_incluido / 4). Radio = cuerda / (2*sin(2*atan(bulge)))
    """
    radios_chicos = []
    try:
        n_verts = ent.NumberOfVertices
        coords  = list(ent.Coordinates)   # [x0,y0, x1,y1, ...]
        for i in range(n_verts):
            b = ent.GetBulge(i)
            if abs(b) < 1e-9:
                continue
            x0 = coords[i*2];     y0 = coords[i*2+1]
            xi = (i+1) % n_verts
            x1 = coords[xi*2];    y1 = coords[xi*2+1]
            cuerda = _math2.hypot(x1-x0, y1-y0)
            if cuerda < 1e-9:
                continue
            radio = cuerda / (2.0 * abs(_math2.sin(2.0 * _math2.atan(abs(b)))))
            if radio < radio_min:
                radios_chicos.append(round(radio, 3))
    except Exception:
        pass
    return radios_chicos


def _offset_inward(doc, msp, ent, dist):
    """
    Offset hacia adentro de una polyline cerrada.
    Devuelve la nueva entidad (la de menor área bbox).
    """
    try:
        results = ent.Offset(dist)
        if not results:
            results = ent.Offset(-dist)
        candidates = list(results) if results else []
        if len(candidates) < 2:
            # probar dirección contraria
            r2 = ent.Offset(-dist)
            candidates += list(r2) if r2 else []
        if not candidates:
            return None
        # Elegir el de menor área (el de adentro)
        area_orig = _area_bbox(ent)
        inward = [c for c in candidates if _area_bbox(c) < area_orig]
        if not inward:
            inward = candidates
        inward.sort(key=_area_bbox)
        return inward[0]
    except Exception:
        return None


def _hatch_solido(doc, msp, outer_ent, inner_ent, layer):
    """Crea hatch SOLID entre outer_ent e inner_ent (si inner_ent es None, rellena outer)."""
    try:
        h = msp.AddHatch(0, "SOLID", True)   # patternType=0=predefined, assoc=True
        outer_loop = win32com.client.VARIANT(
            pythoncom.VT_ARRAY | pythoncom.VT_DISPATCH, [outer_ent])
        h.AppendOuterLoop(outer_loop)
        if inner_ent is not None:
            inner_loop = win32com.client.VARIANT(
                pythoncom.VT_ARRAY | pythoncom.VT_DISPATCH, [inner_ent])
            h.AppendInnerLoop(inner_loop)
        h.Evaluate()
        h.Layer = layer
        doc.Regen(0)
        return h
    except Exception as e:
        return None


def _length_polyline(ent):
    try:
        return float(ent.Length)
    except Exception:
        return 0.0


def _crear_arte_autocad(ruta_dwg: str, log_fn=None):
    """
    Pipeline completo de creación de arte en AutoCAD.
    ruta_dwg: ruta al _PLANO.dwg ya extraído.
    """
    if log_fn is None:
        log_fn = print

    pythoncom.CoInitialize()
    try:
        try:
            acad = win32com.client.GetActiveObject("AutoCAD.Application")
        except Exception:
            raise RuntimeError("AutoCAD no está abierto. Ábrelo primero.")

        log_fn(f"  Abriendo: {os.path.basename(ruta_dwg)}")
        doc = acad.Documents.Open(os.path.abspath(ruta_dwg), False, False)
        time.sleep(2)
        msp = doc.ModelSpace

        # ── 1. Quitar GlobalWidth (ConstantWidth) en todas las polylines ─────
        log_fn("  [1] Limpiando GlobalWidth a 0...")
        for ent in msp:
            try:
                if "Polyline" in ent.ObjectName:
                    ent.ConstantWidth = 0.0
                    ent.Update()
            except Exception:
                pass

        # ── 2. Verificar que todos los contornos PERIMETRO/BN estén cerrados ─
        log_fn("  [2] Verificando contornos cerrados...")
        _no_cerrados = []
        for ent in _ents_por_patron(msp, _PAT_PERIM + _PAT_BN):
            try:
                if not ent.Closed:
                    _no_cerrados.append(ent.Layer)
            except Exception:
                pass
        if _no_cerrados:
            msg = "ALERTA: Los siguientes layers tienen contornos NO cerrados:\n" + \
                  "\n".join(set(_no_cerrados)) + \
                  "\n\nCorrige y vuelve a ejecutar."
            log_fn(f"  ERROR contornos abiertos: {_no_cerrados}")
            raise RuntimeError(msg)
        log_fn("  Todos los contornos están cerrados ✔")

        # ── 3. Verificar radios mínimos en PERIMETRO ──────────────────────────
        log_fn("  [3] Verificando radios (mín 15 mm)...")
        _radios_malos = []
        for ent in _ents_por_patron(msp, _PAT_PERIM):
            _radios_malos += _verificar_radios_acad(ent, _RADIO_MIN)
        if _radios_malos:
            _min_r = min(_radios_malos)
            log_fn(f"  WARN radios < 15mm: {sorted(set(_radios_malos))[:8]} — continuando.")
            # mostrar messagebox en hilo principal
            import ctypes
            ctypes.windll.user32.MessageBoxW(
                0,
                f"Radios menores a 15 mm detectados.\nMínimo: {_min_r:.3f} mm\n\nEl proceso continuará.",
                "AGP Arte Maker — Advertencia radios",
                0x30
            )
        else:
            log_fn("  Radios OK ✔")

        # ── 4. Detectar modo degradé: contar polylines cerradas en BN/PHANTOM ─
        _bn_ents = [e for e in _ents_por_patron(msp, _PAT_BN)
                    if e.Closed and "Polyline" in e.ObjectName]
        _bn_ents.sort(key=_area_bbox, reverse=True)   # mayor área = exterior
        CON_DEGRADE = len(_bn_ents) >= 2
        bn_ent = _bn_ents[0] if _bn_ents else None
        log_fn(f"  [4] BN encontrados: {len(_bn_ents)}  → {'CON degradé' if CON_DEGRADE else 'SIN degradé'}")

        # ── 5. Obtener perímetro ──────────────────────────────────────────────
        _perim_ents = [e for e in _ents_por_patron(msp, _PAT_PERIM)
                       if e.Closed and "Polyline" in e.ObjectName]
        if not _perim_ents:
            raise RuntimeError("No se encontró curva de PERÍMETRO cerrada.")
        perim_ent = sorted(_perim_ents, key=_area_bbox, reverse=True)[0]
        log_fn("  [5] Perímetro encontrado ✔")

        # Asegurar layers de arte
        for lyr in [_LAYER_PLANES, _LAYER_K2, _LAYER_K]:
            _asegurar_layer(doc, lyr)

        # ── 6. Offset perímetro 0.5mm hacia adentro ───────────────────────────
        log_fn(f"  [6] Offset perímetro {_OFFSET_PERIM} mm...")
        off_perim = _offset_inward(doc, msp, perim_ent, _OFFSET_PERIM)
        if not off_perim:
            raise RuntimeError("No se pudo crear offset del perímetro.")
        off_perim.Layer = _LAYER_PLANES

        # ── 7. Hatch SOLID perímetro → offset 0.5 (layer k2) ─────────────────
        log_fn("  [7] Hatch k2 (borde perímetro)...")
        h_k2 = _hatch_solido(doc, msp, perim_ent, off_perim, _LAYER_K2)
        if not h_k2:
            log_fn("  WARN: hatch k2 no se pudo crear.")

        # ── 8. Hatch SOLID BN → offset 0.5 (layer k) ─────────────────────────
        log_fn("  [8] Hatch k (banda negra)...")
        if bn_ent:
            h_k = _hatch_solido(doc, msp, bn_ent, off_perim, _LAYER_K)
            if not h_k:
                log_fn("  WARN: hatch k no se pudo crear.")
        else:
            log_fn("  WARN: no se encontró banda negra.")

        # ── 9. Degradé ────────────────────────────────────────────────────────
        if CON_DEGRADE and bn_ent:
            log_fn(f"  [9] Degradé: offset BN {_OFFSET_BN_DEG} mm...")
            off_bn = _offset_inward(doc, msp, bn_ent, _OFFSET_BN_DEG)
            if off_bn:
                off_bn.Layer = _LAYER_PLANES
                longitud = _length_polyline(off_bn)
                n_pepas  = int(round(longitud / _DIVISOR_DEG)) if longitud > 0 else 0
                log_fn(f"  Longitud offset BN: {longitud:.2f} mm  pepas: {n_pepas}")

                if n_pepas > 0:
                    # Insertar bloque 25 desde CAJETIN_DWG si no existe
                    try:
                        doc.Blocks.Item(_BLOQUE_25)
                        log_fn(f"  Bloque '{_BLOQUE_25}' ya existe en el doc.")
                    except Exception:
                        log_fn(f"  Importando bloque '{_BLOQUE_25}' desde cajetines...")
                        doc.SendCommand(
                            f'-INSERT "{os.path.abspath(CAJETIN_DWG)}" \n'
                            f'C \n0,0,0\n1\n1\n0\n'
                        )
                        time.sleep(1.5)
                        doc.SendCommand("U \n")   # deshacer la inserción del bloque completo
                        time.sleep(0.5)

                    # Usar DIVIDE con bloque 25 sobre el offset BN
                    log_fn(f"  Aplicando DIVIDE con bloque '{_BLOQUE_25}' x{n_pepas}...")
                    handle = off_bn.Handle
                    doc.SendCommand(
                        f'(handent "{handle}") \n'
                    )
                    time.sleep(0.3)
                    doc.SendCommand(
                        f'DIVIDE \n'
                        f'(handent "{handle}") \n'
                        f'B \n'
                        f'{_BLOQUE_25} \n'
                        f'Y \n'
                        f'{n_pepas} \n'
                    )
                    time.sleep(1.0)
                    log_fn(f"  Degradé aplicado ✔")
            else:
                log_fn("  WARN: no se pudo crear offset interior de BN.")
        else:
            log_fn("  [9] Sin degradé — omitido.")

        # ── 10. Importar cajetines ────────────────────────────────────────────
        log_fn("  [10] Importando cajetines...")
        _ids_antes = set()
        for e in msp:
            try: _ids_antes.add(e.Handle)
            except Exception: pass

        abs_caj = os.path.abspath(CAJETIN_DWG)
        doc.SendCommand(f'-INSERT "{abs_caj}" \n0,0,0\n1\n1\n0\n')
        time.sleep(2)
        doc.SendCommand("EXPLODE \nL \n \n")
        time.sleep(1)

        # ── 11. Reemplazar logo ───────────────────────────────────────────────
        log_fn("  [11] Reemplazando logo...")
        _logo_plano = [e for e in msp
                       if any(p in e.Layer.upper() for p in _PAT_LOGO)]
        _logo1 = []
        for e in msp:
            try:
                if "LOGO1" in e.Layer.upper():
                    _logo1.append(e)
            except Exception:
                pass

        if _logo_plano and _logo1:
            cx_pl, cy_pl = _centro_bbox(_logo_plano)
            cx_l1, cy_l1 = _centro_bbox(_logo1)
            if cx_pl and cx_l1:
                dx, dy = cx_pl - cx_l1, cy_pl - cy_l1
                for e in _logo1:
                    try:
                        e.Move(_pt(0,0), _pt(dx, dy))
                    except Exception:
                        pass
                for e in _logo_plano:
                    try: e.Delete()
                    except Exception: pass
                log_fn("  Logo reemplazado ✔")

        # ── 12. Centrar cajetín sobre la pieza ────────────────────────────────
        log_fn("  [12] Centrando cajetín...")
        _caj_ents = [e for e in msp if "CAJETIN" in e.Layer.upper()]
        if _caj_ents:
            cx_p, cy_p = _centro_bbox([perim_ent])
            cx_c, cy_c = _centro_bbox(_caj_ents)
            if cx_p and cx_c:
                dx, dy = cx_p - cx_c, cy_p - cy_c
                for e in _caj_ents:
                    try: e.Move(_pt(0,0), _pt(dx, dy))
                    except Exception: pass
                log_fn("  Cajetín centrado ✔")

        # ── 13. Mover PERIMETRO/BN al layer PLANES ────────────────────────────
        log_fn("  [13] Moviendo geometría original a PLANES...")
        for ent in _ents_por_patron(msp, _PAT_PERIM + _PAT_BN):
            try: ent.Layer = _LAYER_PLANES
            except Exception: pass

        # ── 14. Guardar ───────────────────────────────────────────────────────
        log_fn("  [14] Guardando...")
        doc.SendCommand("QSAVE \n")
        time.sleep(1)

        # ── 15. Purge ─────────────────────────────────────────────────────────
        log_fn("  [15] Purgando capas vacías...")
        doc.SendCommand("-PURGE \nA \n \nN \n")
        time.sleep(1)

        log_fn("  ✔ Arte creado correctamente.")

    finally:
        pythoncom.CoUninitialize()


# ─── HELPERS ─────────────────────────────────────────────────────────────────

import re as _re

def _extraer_codigos(ruta_archivo: str) -> list:
    """
    Extrae los códigos numéricos del final del nombre del plano.
    Lee dígitos de derecha a izquierda hasta acumular 6, ignorando letras.
    Ej: '1576 00 001'     → ['001']
        '1795 003 001-002' → ['001', '002']
        '1576 00 00'       → ['00', '00']
    """
    base = os.path.splitext(os.path.basename(ruta_archivo))[0]
    grupos = _re.findall(r'\d+', base)   # todos los grupos numéricos
    if not grupos:
        return []
    codigos = []
    total   = 0
    for g in reversed(grupos):           # de derecha a izquierda
        if total + len(g) > 6:
            break
        codigos.insert(0, g)
        total += len(g)
    return codigos


def _buscar_artes(ruta: str, codigos: list) -> list:
    """
    Busca recursivamente dentro de carpetas ARTES (y sus subcarpetas).
    Solo retorna archivos que coincidan con alguno de los códigos.
    Un archivo coincide si contiene el código exacto como grupo numérico.
    """
    resultados = []
    for raiz, dirs, archivos in os.walk(ruta):
        dirs[:] = [d for d in dirs if not d.startswith(".")]
        partes = raiz.replace("\\", "/").upper().split("/")
        if "ARTES" not in partes:
            continue
        for archivo in sorted(archivos):
            if os.path.splitext(archivo)[1].lower() not in (".dwg", ".3dm"):
                continue
            nombre_sin_ext = os.path.splitext(archivo)[0]
            nums_archivo   = _re.findall(r'\d+', nombre_sin_ext)
            coincide = bool(codigos) and any(c in nums_archivo for c in codigos)
            if not coincide:
                continue                 # solo mostrar coincidencias
            rel = os.path.relpath(raiz, ruta)
            resultados.append({
                "version":       rel,
                "archivo":       archivo,
                "ruta_completa": os.path.join(raiz, archivo),
                "coincide":      True,
            })
    resultados.sort(key=lambda x: (x["version"], x["archivo"]))
    return resultados


def _ruta_planos(ruta_dwg: str) -> str:
    """
    Crea y devuelve la carpeta PLANOS junto al DWG del plano.
    Ej: ...\\V-00 EPEL\\1774 001.dwg  →  ...\\V-00 EPEL\\PLANOS\\
    """
    carpeta_version = os.path.dirname(os.path.abspath(ruta_dwg))
    destino = os.path.join(carpeta_version, "PLANOS")
    os.makedirs(destino, exist_ok=True)
    return destino


import math as _math

_CAPAS_COMP = [
    ("PERIMETRO",   ["PERIMETRO"],                                   ["PERIMETRO"]),
    ("BANDA NEGRA", ["BANDA NEGRA","BANDANEGRA","BN","PHANTOM"],     ["BANDA NEGRA","BANDANEGRA","BN","PHANTOM"]),
    ("LOGO",        ["LOGO","TRAZABILIDAD"],                          ["LOGO","TRAZABILIDAD"]),
]
_TOL = 0.012   # 1.2 % tolerancia en dimensiones


def _bbox_entidades(coleccion, patrones):
    mn = [1e18, 1e18]; mx = [-1e18, -1e18]; ok = False
    for ent in coleccion:
        try:
            if not any(p in ent.Layer.upper() for p in patrones):
                continue
            lo, hi = ent.GetBoundingBox()
            mn[0]=min(mn[0],lo[0]); mn[1]=min(mn[1],lo[1])
            mx[0]=max(mx[0],hi[0]); mx[1]=max(mx[1],hi[1])
            ok = True
        except Exception:
            pass
    return (mn[0],mn[1],mx[0],mx[1]) if ok else None


def _puntos_entidades(coleccion, patrones, max_pts=300):
    """Extrae puntos de muestra de entidades en las capas indicadas."""
    pts = []
    for ent in coleccion:
        try:
            if not any(p in ent.Layer.upper() for p in patrones):
                continue
            n = ent.ObjectName
            if n in ("AcDbPolyline",):
                c = list(ent.Coordinates)
                for i in range(0, len(c)-1, 2):
                    pts.append((c[i], c[i+1]))
            elif n == "AcDb2dPolyline":
                c = list(ent.Coordinates)
                for i in range(0, len(c)-2, 3):
                    pts.append((c[i], c[i+1]))
            elif n == "AcDbLine":
                sp=ent.StartPoint; ep=ent.EndPoint
                pts.append((sp[0],sp[1])); pts.append((ep[0],ep[1]))
            elif n in ("AcDbCircle","AcDbArc"):
                ce=ent.Center; r=ent.Radius
                if n=="AcDbArc":
                    a0,a1=ent.StartAngle,ent.EndAngle
                    if a1<a0: a1+=2*_math.pi
                    angs=[a0+(a1-a0)*i/16 for i in range(17)]
                else:
                    angs=[2*_math.pi*i/16 for i in range(16)]
                for a in angs:
                    pts.append((ce[0]+r*_math.cos(a), ce[1]+r*_math.sin(a)))
            elif n=="AcDbSpline":
                fp=list(ent.FitPoints)
                for i in range(0,len(fp)-2,3):
                    pts.append((fp[i],fp[i+1]))
        except Exception:
            pass
    if len(pts)>max_pts:
        step=max(1,len(pts)//max_pts)
        pts=pts[::step]
    return pts


def _transformar(pts, rot_deg, mirror, cx, cy):
    """Aplica espejo+rotación a puntos centrados en (cx,cy)."""
    rad=_math.radians(rot_deg); cos_r=_math.cos(rad); sin_r=_math.sin(rad)
    res=[]
    for x,y in pts:
        x-=cx; y-=cy
        if mirror: x=-x
        res.append((x*cos_r-y*sin_r, x*sin_r+y*cos_r))
    return res


def _score_transform(pts_arte, pts_plano, rot_deg, mirror, cx_p, cy_p, cx_a, cy_a):
    """Distancia media mínima entre pts_arte y pts_plano transformados."""
    if not pts_arte or not pts_plano:
        return 1e9
    tp = _transformar(pts_plano, rot_deg, mirror, cx_p, cy_p)
    # desplazar al centro del arte
    total = 0.0
    for ax, ay in pts_arte:
        ax -= cx_a; ay -= cy_a
        d = min((ax-px)**2+(ay-py)**2 for px,py in tp)
        total += d**0.5
    return total / len(pts_arte)


def _dims(bbox):
    if bbox is None: return None, None
    return abs(bbox[2]-bbox[0]), abs(bbox[3]-bbox[1])


def _centro(bbox):
    if bbox is None: return None
    return (bbox[0]+bbox[2])/2, (bbox[1]+bbox[3])/2


def _dims_ok(w1,h1,w2,h2):
    def pct(a,b): return abs(a-b)/max(a,b,1e-6)
    if pct(w1,w2)<_TOL and pct(h1,h2)<_TOL: return True
    if pct(w1,h2)<_TOL and pct(h1,w2)<_TOL: return True
    return False


def _overlay_autocad(ruta_arte: str, ruta_plano: str, log_fn=None):
    if log_fn is None:
        log_fn = print

    pythoncom.CoInitialize()
    try:
        try:
            acad = win32com.client.GetActiveObject("AutoCAD.Application")
        except Exception:
            raise RuntimeError("AutoCAD no está abierto.\nAbre AutoCAD primero y vuelve a intentarlo.")

        log_fn(f"  Abriendo: {os.path.basename(ruta_arte)}")
        doc = acad.Documents.Open(os.path.abspath(ruta_arte), False, False)
        time.sleep(2)
        msp = doc.ModelSpace

        # ── Adjuntar XREF ────────────────────────────────────────────────────
        abs_plano = os.path.abspath(ruta_plano)
        log_fn(f"  Adjuntando plano XREF: {os.path.basename(abs_plano)}")
        xref_ref = None
        try:
            pt = win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, [0.0,0.0,0.0])
            xref_ref = msp.AttachExternalReference(abs_plano, "PLANO_REF", pt, 1.0,1.0,1.0, 0.0, False)
            log_fn("  XREF adjuntado.")
        except Exception as e_com:
            log_fn(f"  API COM falló ({e_com}), usando SendCommand...")
            doc.SendCommand(f'-XREF A "{abs_plano}" \nPLANO_REF\n0,0,0\n1\n1\n0\n')
            time.sleep(1.5)
            # buscar la referencia creada
            for obj in msp:
                try:
                    if obj.ObjectName == "AcDbBlockReference" and "PLANO_REF" in obj.Name.upper():
                        xref_ref = obj
                        break
                except Exception:
                    pass

        time.sleep(1.0)

        # ── Leer entidades del bloque XREF ───────────────────────────────────
        xref_blk = None
        try:
            xref_blk = doc.Blocks.Item("PLANO_REF")
        except Exception:
            pass

        # ── Verificar dimensiones por capa ───────────────────────────────────
        log_fn("─" * 48)
        log_fn("  COMPARACIÓN DE CAPAS:")
        resumen = []
        dims_ok_perim = False

        for nombre_capa, pat_arte, pat_plano in _CAPAS_COMP:
            ba = _bbox_entidades(msp,      pat_arte)
            bp = _bbox_entidades(xref_blk, pat_plano) if xref_blk else None
            wa, ha = _dims(ba)
            wp, hp = _dims(bp)

            if wa is None and wp is None:
                log_fn(f"  [{nombre_capa}]  —  no encontrado en ninguno"); resumen.append((nombre_capa, None)); continue
            if wa is None:
                log_fn(f"  [{nombre_capa}]  —  no encontrado en el ARTE"); resumen.append((nombre_capa, None)); continue
            if wp is None:
                log_fn(f"  [{nombre_capa}]  —  no encontrado en el PLANO"); resumen.append((nombre_capa, None)); continue

            ok = _dims_ok(wa, ha, wp, hp)
            if nombre_capa == "PERIMETRO":
                dims_ok_perim = ok
            estado = "✔  COINCIDE" if ok else "✘  NO COINCIDE"
            log_fn(f"  [{nombre_capa}]  {estado}  arte {wa:.1f}×{ha:.1f}  plano {wp:.1f}×{hp:.1f} mm")
            resumen.append((nombre_capa, ok))

        # ── Buscar mejor transformación (rot + espejo) con puntos ─────────────
        mejor_rot, mejor_mirror, mejor_score = 0, False, 1e9
        desc_transform = "0° sin espejo"

        bbox_arte_p  = _bbox_entidades(msp,      ["PERIMETRO"])
        bbox_plano_p = _bbox_entidades(xref_blk, ["PERIMETRO"]) if xref_blk else None

        if bbox_arte_p and bbox_plano_p:
            pts_arte  = _puntos_entidades(msp,      ["PERIMETRO"])
            pts_plano = _puntos_entidades(xref_blk, ["PERIMETRO"]) if xref_blk else []
            cx_a, cy_a = _centro(bbox_arte_p)
            cx_p, cy_p = _centro(bbox_plano_p)

            if pts_arte and pts_plano:
                log_fn("  Probando 8 transformaciones (4 rotaciones × espejo)...")
                for rot in [0, 90, 180, 270]:
                    for mirror in [False, True]:
                        sc = _score_transform(pts_arte, pts_plano, rot, mirror, cx_p, cy_p, cx_a, cy_a)
                        if sc < mejor_score:
                            mejor_score = sc; mejor_rot = rot; mejor_mirror = mirror
                desc_transform = f"{mejor_rot}°{'  + espejo' if mejor_mirror else ''}"
                log_fn(f"  Mejor transformación: {desc_transform}  (error promedio {mejor_score:.2f} mm)")

        # ── Aplicar transformación al XREF ────────────────────────────────────
        if xref_ref and bbox_arte_p:
            cx_arte, cy_arte = _centro(bbox_arte_p)
            try:
                xref_ref.Rotation = _math.radians(mejor_rot)
                xref_ref.XScaleFactor = -1.0 if mejor_mirror else 1.0
                time.sleep(0.3)
                lo2, hi2 = xref_ref.GetBoundingBox()
                cx_x = (lo2[0] + hi2[0]) / 2
                cy_x = (lo2[1] + hi2[1]) / 2
                ins   = xref_ref.InsertionPoint
                nuevo_ins = win32com.client.VARIANT(
                    pythoncom.VT_ARRAY | pythoncom.VT_R8,
                    [ins[0] + (cx_arte - cx_x),
                     ins[1] + (cy_arte - cy_x),
                     0.0]
                )
                xref_ref.InsertionPoint = nuevo_ins
                log_fn(f"  XREF posicionado: {desc_transform}")
            except Exception as e_pos:
                log_fn(f"  Posicionamiento automático falló: {e_pos}")

        # ── Resultado final ───────────────────────────────────────────────────
        capas_ok = [r for r in resumen if r[1] is True]
        capas_no = [r for r in resumen if r[1] is False]
        log_fn("─" * 48)
        if capas_no:
            log_fn(f"  RESULTADO: ✘ NO COINCIDE en: {', '.join(c[0] for c in capas_no)}")
        elif capas_ok:
            log_fn(f"  RESULTADO: ✔ ARTE CORRECTO  ({desc_transform})")
        else:
            log_fn("  RESULTADO: no se pudo comparar (capas no encontradas)")

        doc.SendCommand("ZOOM E \n")
        time.sleep(0.5)
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
        self._btn_extraer.pack(side="left", padx=(0, 12))

        self._btn_crear = NeonButton(
            card_btn, "✦  CREAR ARTE",
            self._crear_arte, "#7B2FBE", "#5A1F9A", width=185, height=44)
        self._btn_crear.pack(side="left", padx=(0, 12))

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
        self._btn_crear.configure_state(not activo)
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

    # ── CREAR ARTE ────────────────────────────────────────────────────────────

    def _crear_arte(self):
        dwg = self._dwg_plano.get().strip().strip('"')
        if not dwg:
            messagebox.showwarning("Campo requerido",
                                   "Selecciona el archivo DWG del plano extraído (_PLANO.dwg).")
            return
        if not os.path.isfile(dwg):
            messagebox.showerror("Archivo no encontrado", f"No existe:\n{dwg}")
            return
        self._busy(True)
        threading.Thread(target=self._t_crear_arte, args=(dwg,), daemon=True).start()

    def _t_crear_arte(self, ruta_dwg: str):
        self._log("=" * 56)
        self._log("CREAR ARTE — iniciando pipeline AutoCAD...", "ok")
        self._log(f"Plano: {os.path.basename(ruta_dwg)}", "dim")
        try:
            _crear_arte_autocad(ruta_dwg, log_fn=lambda m: self._log(m, "dim"))
            self._log("✔ Arte completado.", "ok")
        except RuntimeError as e:
            self._log(str(e), "err")
            self.after(0, lambda: messagebox.showerror("Error", str(e)))
        except Exception as e:
            self._log(f"ERROR inesperado: {e}", "err")
        finally:
            self._busy(False)

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
        ruta_destino  = _ruta_planos(ruta_plano)           # crea PLANOS/ junto al DWG
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

        codigos = _extraer_codigos(dwg_plano) if dwg_plano else []
        if codigos:
            self._log(f'Códigos buscados: {" | ".join(codigos)}', "dim")
        else:
            self._log("Sin código de plano — se mostrarán todos los artes.", "warn")

        resultados = _buscar_artes(ruta_base, codigos)
        self._resultados = resultados

        self.after(0, self._poblar_tabla, resultados)

        if not resultados:
            self._log("No se encontraron artes coincidentes.", "warn")
        else:
            self._log(f"Se encontraron {len(resultados)} arte(s) coincidente(s).", "ok")
            self._log("Doble clic en una fila para superponer en AutoCAD.", "ok")

        self._busy(False)

    def _poblar_tabla(self, resultados: list):
        for item in self._tree.get_children():
            self._tree.delete(item)
        for r in resultados:
            estado = "✔  COINCIDE"
            tag    = "match"
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
