
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
    "bg":        "#0B0D17",
    "bg2":       "#0F1220",
    "panel":     "#131626",
    "panel2":    "#171B2E",
    "border":    "#1E2540",
    "accent":    "#4FACFF",
    "accent2":   "#2980CC",
    "accent3":   "#50FA7B",
    "btn_ok":    "#27AE60",
    "btn_ok2":   "#1E8449",
    "btn_warn":  "#E67E22",
    "btn_warn2": "#CA6F1E",
    "txt":       "#ECF0FF",
    "txt_dim":   "#4A5A7A",
    "txt_mid":   "#7A90B0",
    "entry_bg":  "#0D1020",
    "entry_fg":  "#4FACFF",
    "log_bg":    "#080A14",
    "log_ok":    "#50FA7B",
    "log_warn":  "#FFB86C",
    "log_err":   "#FF5555",
    "log_dim":   "#44557A",
    "purple":    "#8B5CF6",
    "purple2":   "#6D28D9",
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

# Importar pipeline y diálogo desde crear_arte_acad.py
try:
    from crear_arte_acad import dialogo_cajetin as _dialogo_cajetin
    from crear_arte_acad import actualizar_texto_cajetin as _actualizar_texto_cajetin
    from crear_arte_acad import pipeline as _pipeline_acad
    _CAJETIN_DIALOG_OK = True
except Exception:
    _CAJETIN_DIALOG_OK = False
    _pipeline_acad = None


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


def _crear_arte_autocad(ruta_dwg: str, log_fn=None, valores_cajetin=None,
                        ruta_salida: str = None, perim_index: int = 0) -> int:
    """
    Abre el DWG y ejecuta el pipeline. Devuelve el nº de piezas encontradas.
    perim_index: índice de pieza a procesar (0=más grande).
    """
    if log_fn is None:
        log_fn = print
    if _pipeline_acad is None:
        raise RuntimeError("No se pudo importar el pipeline desde crear_arte_acad.py")

    pythoncom.CoInitialize()
    try:
        try:
            acad = win32com.client.GetActiveObject("AutoCAD.Application")
        except Exception:
            raise RuntimeError("AutoCAD no está abierto. Ábrelo primero.")

        log_fn(f"  Abriendo: {os.path.basename(ruta_dwg)}")
        doc = acad.Documents.Open(os.path.abspath(ruta_dwg), False, False)
        time.sleep(2)
        try:
            doc.Activate()
            time.sleep(0.5)
        except Exception:
            pass

        n = _pipeline_acad(doc, log_fn=log_fn,
                           valores_cajetin=valores_cajetin,
                           ruta_salida=ruta_salida,
                           perim_index=perim_index)
        return n or 1
    finally:
        pythoncom.CoUninitialize()


# ─── HELPERS ─────────────────────────────────────────────────────────────────

import re as _re

def _extraer_codigos(ruta_archivo: str) -> list:
    """
    Extrae el sufijo numérico exacto del plano (los últimos grupos, hasta 6 dígitos total).
    Ej: '1795 003 001-002' → ['001', '002']
        '1795 003 007'     → ['007']
        '1576 00 001'      → ['001']
    """
    base = os.path.splitext(os.path.basename(ruta_archivo))[0]
    grupos = _re.findall(r'\d+', base)
    if not grupos:
        return []
    codigos = []
    total = 0
    for g in reversed(grupos):
        if total + len(g) > 6:
            break
        codigos.insert(0, g)
        total += len(g)
    return codigos


def _buscar_artes(ruta: str, codigos: list) -> list:
    """
    Busca archivos de arte con matching estricto:
    - Solo en carpetas cuyo path contenga 'ARTES' (cualquier subcarpeta dentro)
    - Solo archivos que empiecen con 'P' (mayúscula o minúscula)
    - El archivo debe terminar EXACTAMENTE con los mismos grupos numéricos que el plano
      (en orden). Ej: códigos ['007','008'] → el archivo debe tener ...007-008 al final,
      NOT solo 007 ni 007-009.
    """
    resultados = []
    for raiz, dirs, archivos in os.walk(ruta):
        dirs[:] = [d for d in dirs if not d.startswith(".")]
        partes = raiz.replace("\\", "/").upper().split("/")
        # Buscar dentro de cualquier carpeta ARTES y sus subcarpetas
        if not any(p == "ARTES" or p.startswith("ARTES") for p in partes):
            continue
        for archivo in sorted(archivos):
            ext = os.path.splitext(archivo)[1].lower()
            if ext not in (".dwg", ".3dm"):
                continue
            # Solo archivos que empiecen con P
            if not archivo.upper().startswith("P"):
                continue
            nombre_sin_ext = os.path.splitext(archivo)[0]
            nums_archivo   = _re.findall(r'\d+', nombre_sin_ext)
            # Matching estricto: los últimos len(codigos) grupos numéricos
            # deben ser EXACTAMENTE iguales a codigos, en ese orden
            if not codigos or len(nums_archivo) < len(codigos):
                continue
            if nums_archivo[-len(codigos):] != codigos:
                continue
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
    """Devuelve (y crea) la carpeta PLANOS junto al DWG original."""
    carpeta_version = os.path.dirname(os.path.abspath(ruta_dwg))
    destino = os.path.join(carpeta_version, "PLANOS")
    os.makedirs(destino, exist_ok=True)
    return destino


def _ruta_arte_salida(ruta_dwg_original: str, malla: str = "", pieza: str = "") -> str:
    """
    Devuelve la ruta de salida del arte:
      - Nombre: P {malla} {pieza}.dwg   (o P {nombre_original}.dwg si no hay malla/pieza)
      - Carpeta: ARTES/BN/ si existe subcarpeta llamada exactamente 'BN' (sin más),
                 ARTES/    en caso contrario.
    Crea las carpetas necesarias.
    """
    carpeta      = os.path.dirname(os.path.abspath(ruta_dwg_original))
    artes_dir    = os.path.join(carpeta, "ARTES")
    os.makedirs(artes_dir, exist_ok=True)

    # Buscar subcarpeta exactamente "BN" (case-insensitive, solo si es exactamente eso)
    destino = artes_dir
    try:
        for entry in os.listdir(artes_dir):
            full = os.path.join(artes_dir, entry)
            if os.path.isdir(full) and entry.upper() == "BN":
                destino = full
                break
    except Exception:
        pass

    # Construir nombre del archivo: P {malla} {pieza}.dwg
    partes_nombre = [p.strip() for p in [malla, pieza] if p.strip()]
    if partes_nombre:
        nombre_archivo = "P " + " ".join(partes_nombre) + ".dwg"
    else:
        nombre_base    = os.path.splitext(os.path.basename(ruta_dwg_original))[0]
        nombre_archivo = f"P {nombre_base}.dwg"
    # Eliminar dobles espacios por si algún campo tiene espacios internos raros
    while "  " in nombre_archivo:
        nombre_archivo = nombre_archivo.replace("  ", " ")

    return os.path.join(destino, nombre_archivo)


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
                log_fn(f"  [{nombre_capa}]  —  no encontrado en ninguno:0"); resumen.append((nombre_capa, None)); continue
            if wa is None:
                log_fn(f"  [{nombre_capa}]  —  no encontrado en el ARTE:)"); resumen.append((nombre_capa, None)); continue
            if wp is None:
                log_fn(f"  [{nombre_capa}]  —  no encontrado en el PLANO"); resumen.append((nombre_capa, None)); continue

            ok = _dims_ok(wa, ha, wp, hp)
            if nombre_capa == "PERIMETRO":
                dims_ok_perim = ok
            estado = "OK  COINCIDE" if ok else "X  NO COINCIDE"
            log_fn(f"  [{nombre_capa}]  {estado}  arte {wa:.1f}x{ha:.1f}  plano {wp:.1f}x{hp:.1f} mm")
            resumen.append((nombre_capa, ok))

        # ── Buscar mejor transformación (rot + espejo) con puntos ─────────────
        mejor_rot, mejor_mirror, mejor_score = 0, False, 1e9
        desc_transform = "0 sin espejo"

        bbox_arte_p  = _bbox_entidades(msp,      ["PERIMETRO"])
        bbox_plano_p = _bbox_entidades(xref_blk, ["PERIMETRO"]) if xref_blk else None

        if bbox_arte_p and bbox_plano_p:
            pts_arte  = _puntos_entidades(msp,      ["PERIMETRO"])
            pts_plano = _puntos_entidades(xref_blk, ["PERIMETRO"]) if xref_blk else []
            cx_a, cy_a = _centro(bbox_arte_p)
            cx_p, cy_p = _centro(bbox_plano_p)

            if pts_arte and pts_plano:
                log_fn("  Probando 8 transformaciones (4 rotaciones x espejo)...")
                for rot in [0, 90, 180, 270]:
                    for mirror in [False, True]:
                        sc = _score_transform(pts_arte, pts_plano, rot, mirror, cx_p, cy_p, cx_a, cy_a)
                        if sc < mejor_score:
                            mejor_score = sc; mejor_rot = rot; mejor_mirror = mirror
                desc_transform = f"{mejor_rot}{'  + espejo' if mejor_mirror else ''}"
                log_fn(f"  Mejor transformacion: {desc_transform}  (error promedio {mejor_score:.2f} mm)")

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
                log_fn(f"  Posicionamiento automatico fallo: {e_pos}")

        # ── Resultado final ───────────────────────────────────────────────────
        capas_ok = [r for r in resumen if r[1] is True]
        capas_no = [r for r in resumen if r[1] is False]
        log_fn("─" * 48)
        if capas_no:
            log_fn(f"  RESULTADO: X NO COINCIDE en: {', '.join(c[0] for c in capas_no)}")
        elif capas_ok:
            log_fn(f"  RESULTADO: OK ARTE CORRECTO  ({desc_transform})")
        else:
            log_fn("  RESULTADO: no se pudo comparar (capas no encontradas)")

        doc.SendCommand("ZOOM E \n")
        time.sleep(0.5)
    finally:
        pythoncom.CoUninitialize()


# ─── WIDGET HELPERS ──────────────────────────────────────────────────────────

class NeonButton(tk.Frame):
    """Botón premium con borde colored, subtítulo opcional y efecto hover."""
    def __init__(self, parent, text, command, color, hover_color,
                 width=180, height=48, subtitle=""):
        super().__init__(parent, bg=C["bg"], cursor="hand2")
        self._cmd        = command
        self._color      = color
        self._hover      = hover_color
        self._enabled    = True

        # Contenedor interior con el color del botón
        self._inner = tk.Frame(self, bg=color, padx=2, pady=2)
        self._inner.pack(fill="both", expand=True)

        # Texto principal
        self._lbl = tk.Label(self._inner, text=text, font=FONT_HDR,
                             bg=color, fg="white",
                             padx=18, pady=6, cursor="hand2")
        self._lbl.pack(fill="x")

        # Subtítulo opcional
        self._sub = None
        if subtitle:
            self._sub = tk.Label(self._inner, text=subtitle,
                                 font=("Segoe UI", 7),
                                 bg=color, fg="white",
                                 padx=18, pady=0, cursor="hand2")
            self._sub.pack(fill="x")

        # Barra inferior de 3px (accent bar)
        self._bar = tk.Frame(self, bg=hover_color, height=3)
        self._bar.pack(fill="x", side="bottom")

        for w in (self, self._inner, self._lbl) + ((self._sub,) if self._sub else ()):
            w.bind("<Enter>",    lambda e: self._on_enter())
            w.bind("<Leave>",    lambda e: self._on_leave())
            w.bind("<Button-1>", lambda e: self._click())

    def _on_enter(self):
        if self._enabled:
            self._inner.configure(bg=self._hover)
            self._lbl.configure(bg=self._hover)
            if self._sub:
                self._sub.configure(bg=self._hover)

    def _on_leave(self):
        col = self._color if self._enabled else C["txt_dim"]
        self._inner.configure(bg=col)
        self._lbl.configure(bg=col)
        if self._sub:
            self._sub.configure(bg=col)

    def _click(self):
        if not self._enabled:
            return
        self._inner.configure(bg="white")
        self._lbl.configure(bg="white", fg=self._color)
        if self._sub:
            self._sub.configure(bg="white", fg=self._color)
        self.after(130, self._restore)
        self._cmd()

    def _restore(self):
        self._inner.configure(bg=self._color)
        self._lbl.configure(bg=self._color, fg="white")
        if self._sub:
            self._sub.configure(bg=self._color, fg="white")

    def configure_state(self, enabled: bool):
        self._enabled = enabled
        col = self._color if enabled else C["txt_dim"]
        self._inner.configure(bg=col)
        self._lbl.configure(bg=col)
        if self._sub:
            self._sub.configure(bg=col)
        self._bar.configure(bg=self._hover if enabled else C["border"])


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


class ScanLine(tk.Frame):
    """Barra decorativa en el header — compatible Python 3.14 / Tk 9."""
    def __init__(self, parent, **kw):
        super().__init__(parent, bg=C["accent"], height=2, **kw)


class PulseBar(ttk.Progressbar):
    """Barra de progreso indeterminada — compatible Python 3.14 / Tk 9."""
    def __init__(self, parent, **kw):
        super().__init__(parent, mode="indeterminate", length=0,
                         style="Pulse.Horizontal.TProgressbar", **kw)

    def start(self):
        super().start(12)

    def stop(self):
        super().stop()
        self["value"] = 0


# ─── APP PRINCIPAL ────────────────────────────────────────────────────────────

class ArteMakerApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("AGP GROUP — Arte Maker")
        self.configure(bg=C["bg"])
        self.resizable(True, True)
        self.minsize(860, 700)

        self._ruta_base    = tk.StringVar()
        self._dwg_plano    = tk.StringVar()
        self._resultados: list = []
        self._dot_count    = 0

        # Escuchar cambios en el campo DWG para actualizar el badge
        self._dwg_plano.trace_add("write", self._on_dwg_changed)

        self._apply_ttk_style()
        self._build_ui()
        self._centrar(980, 800)

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
                         rowheight=28)
        style.configure("Treeview.Heading",
                         background=C["panel2"],
                         foreground=C["accent"],
                         font=("Segoe UI", 9, "bold"),
                         relief="flat",
                         padding=(6, 4))
        style.map("Treeview",
                  background=[("selected", C["accent2"])],
                  foreground=[("selected", "#FFFFFF")])

        style.configure("Pulse.Horizontal.TProgressbar",
                         troughcolor=C["bg"],
                         background=C["accent"],
                         borderwidth=0,
                         thickness=4)

        style.configure("Vertical.TScrollbar",
                         background=C["panel2"],
                         troughcolor=C["bg2"],
                         arrowcolor=C["accent"],
                         borderwidth=0)

    # ── Build UI ──────────────────────────────────────────────────────────────

    def _build_ui(self):
        # ═══════════════════════════════
        # HEADER
        # ═══════════════════════════════
        hdr = tk.Frame(self, bg=C["bg2"])
        hdr.pack(fill="x")

        # Barra top 3px accent
        tk.Frame(hdr, bg=C["accent"], height=3).pack(fill="x")

        inner_hdr = tk.Frame(hdr, bg=C["bg2"], pady=12, padx=24)
        inner_hdr.pack(fill="x")
        inner_hdr.columnconfigure(1, weight=1)

        # Columna izquierda: logo + título
        left_hdr = tk.Frame(inner_hdr, bg=C["bg2"])
        left_hdr.grid(row=0, column=0, sticky="w")

        tk.Label(left_hdr, text="AGP", font=("Segoe UI", 9, "bold"),
                 bg=C["bg2"], fg=C["txt_mid"]).pack(side="left", padx=8)
        tk.Label(left_hdr, text="ARTE MAKER",
                 font=("Segoe UI", 22, "bold"),
                 bg=C["bg2"], fg=C["accent"]).pack(side="left")

        # Columna derecha: badge plano activo
        right_hdr = tk.Frame(inner_hdr, bg=C["bg2"])
        right_hdr.grid(row=0, column=2, sticky="e", padx=(0, 4))

        tk.Label(right_hdr, text="PLANO ACTIVO",
                 font=("Segoe UI", 7, "bold"),
                 bg=C["bg2"], fg=C["txt_dim"]).pack(anchor="e")

        self._badge_frame = tk.Frame(right_hdr, bg=C["panel2"],
                                     padx=10, pady=3)
        self._badge_frame.pack(anchor="e")
        self._badge_lbl = tk.Label(self._badge_frame,
                                   text="— sin seleccionar —",
                                   font=("Segoe UI", 9, "bold"),
                                   bg=C["panel2"], fg=C["txt_dim"])
        self._badge_lbl.pack()

        # ScanLine animada
        ScanLine(hdr).pack(fill="x")
        tk.Frame(hdr, bg=C["border"], height=1).pack(fill="x")

        # Franja bottom decorativa
        tk.Frame(self, bg=C["accent2"], height=2).pack(fill="x", side="bottom")

        # ═══════════════════════════════
        # BODY
        # ═══════════════════════════════
        body = tk.Frame(self, bg=C["bg"], padx=24, pady=14)
        body.pack(fill="both", expand=True)

        # ─── TARJETA: inputs ────────────────────────────────────────────────
        outer_in, card_in = self._card(body, "CONFIGURACION")
        outer_in.pack(fill="x", pady=10)
        card_in.columnconfigure(1, weight=1)

        # Ruta base
        self._lbl_field(card_in, "Ruta del vehiculo / modelo / version:", 0)
        self._entry_row = GlowEntry(card_in, self._ruta_base)
        self._entry_row.grid(row=1, column=0, columnspan=2, sticky="ew",
                             padx=(0, 8), pady=(2, 2))
        self._btn_exp_base = self._explore_btn(card_in, self._explorar_base)
        self._btn_exp_base.grid(row=1, column=2, pady=(2, 2), padx=(4, 0))
        tk.Label(card_in,
                 text="  Puede ser la carpeta del vehiculo, modelo o version — la busqueda es recursiva",
                 font=FONT_SMALL, bg=C["panel"], fg=C["txt_dim"]
                 ).grid(row=2, column=0, columnspan=3, sticky="w", pady=(0, 8))

        # Plano DWG
        self._lbl_field(card_in, "Plano DWG original:", 3)
        self._entry_dwg = GlowEntry(card_in, self._dwg_plano)
        self._entry_dwg.grid(row=4, column=0, columnspan=2, sticky="ew",
                              padx=(0, 8), pady=(2, 2))
        self._btn_exp_dwg = self._explore_btn(card_in, self._explorar_dwg)
        self._btn_exp_dwg.grid(row=4, column=2, pady=(2, 2), padx=(4, 0))
        tk.Label(card_in, text="  Necesario para EXTRAER PLANO y para la superposicion",
                 font=FONT_SMALL, bg=C["panel"], fg=C["txt_dim"]
                 ).grid(row=5, column=0, columnspan=3, sticky="w", pady=(0, 4))

        # ─── TARJETA: workflow buttons ───────────────────────────────────────
        outer_wf, card_wf = self._card(body, "WORKFLOW")
        outer_wf.pack(fill="x", pady=10)

        btn_row = tk.Frame(card_wf, bg=C["panel"])
        btn_row.pack(fill="x", pady=6)

        self._btn_extraer = NeonButton(
            btn_row,
            text="▶  EXTRAER PLANO",
            command=self._extraer,
            color=C["btn_ok"],
            hover_color=C["btn_ok2"],
            width=200, height=48,
            subtitle="Filtra layers del DWG original")
        self._btn_extraer.pack(side="left", padx=10)

        self._btn_crear = NeonButton(
            btn_row,
            text="✦  CREAR ARTE",
            command=self._crear_arte,
            color=C["purple"],
            hover_color=C["purple2"],
            width=185, height=48,
            subtitle="Pipeline AutoCAD completo")
        self._btn_crear.pack(side="left", padx=10)

        self._btn_comprobar = NeonButton(
            btn_row,
            text="⊕  BUSCAR ARTE",
            command=self._comprobar,
            color=C["btn_warn"],
            hover_color=C["btn_warn2"],
            width=200, height=48,
            subtitle="Encuentra artes existentes")
        self._btn_comprobar.pack(side="left", padx=10)

        # Separador vertical
        tk.Frame(btn_row, bg=C["border"], width=2).pack(side="left", fill="y", padx=10)

        # Botón TODO EN UNO
        self._btn_todo = NeonButton(
            btn_row,
            text="⚡  TODO EN UNO",
            command=self._todo_en_uno,
            color="#E63946",
            hover_color="#B71C2E",
            width=200, height=48,
            subtitle="Extrae plano + crea arte de un click")
        self._btn_todo.pack(side="left")

        # PulseBar — debajo de los botones, oculta hasta que haya actividad
        self._pulse = PulseBar(card_wf)
        self._pulse.pack(fill="x", pady=2)

        # ─── TARJETA: resultados ─────────────────────────────────────────────
        outer_tbl = tk.Frame(body, bg=C["border"], padx=1, pady=1)
        outer_tbl.pack(fill="x", pady=10)

        # Header de la tarjeta con badge dinámico
        tbl_hdr = tk.Frame(outer_tbl, bg=C["panel2"], padx=12, pady=6)
        tbl_hdr.pack(fill="x")
        tbl_hdr.columnconfigure(1, weight=1)

        self._lbl_tbl_titulo = tk.Label(
            tbl_hdr,
            text="ARTES ENCONTRADOS — 0 resultados",
            font=("Segoe UI", 8, "bold"),
            bg=C["panel2"], fg=C["txt_mid"])
        self._lbl_tbl_titulo.grid(row=0, column=0, sticky="w")

        # Badge de estado
        self._badge_estado = tk.Label(
            tbl_hdr,
            text="— busca primero",
            font=("Segoe UI", 8, "bold"),
            bg=C["panel2"], fg=C["txt_dim"],
            padx=8, pady=2)
        self._badge_estado.grid(row=0, column=1, sticky="e")

        # Botón "Abrir carpeta"
        self._btn_abrir_carpeta = tk.Button(
            tbl_hdr,
            text="Abrir carpeta",
            bg=C["border"], fg=C["accent"],
            relief="flat", font=FONT_SMALL,
            cursor="hand2",
            activebackground=C["accent"], activeforeground=C["bg"],
            command=self._abrir_carpeta_arte)
        self._btn_abrir_carpeta.grid(row=0, column=2, sticky="e", padx=(8, 0))

        card_tbl = tk.Frame(outer_tbl, bg=C["panel"], padx=12, pady=8)
        card_tbl.pack(fill="both", expand=True)
        card_tbl.columnconfigure(0, weight=1)
        card_tbl.rowconfigure(0, weight=1)

        cols = ("estado", "ruta", "archivo", "tipo")
        self._tree = ttk.Treeview(card_tbl, columns=cols, show="headings", height=6)
        self._tree.heading("estado",  text="Estado")
        self._tree.heading("ruta",    text="Ruta relativa")
        self._tree.heading("archivo", text="Archivo")
        self._tree.heading("tipo",    text="Tipo")
        self._tree.column("estado",  width=110, anchor="center", stretch=False)
        self._tree.column("ruta",    width=350)
        self._tree.column("archivo", width=260)
        self._tree.column("tipo",    width=70, anchor="center", stretch=False)
        self._tree.tag_configure("match",
                                 background="#0A2015",
                                 foreground=C["log_ok"])
        self._tree.tag_configure("other",
                                 background=C["panel2"],
                                 foreground=C["txt_dim"])
        self._tree.grid(row=0, column=0, sticky="nsew")

        sb = ttk.Scrollbar(card_tbl, orient="vertical", command=self._tree.yview)
        sb.grid(row=0, column=1, sticky="ns")
        self._tree.configure(yscrollcommand=sb.set)
        self._tree.bind("<Double-1>", self._on_doble_click)

        # Hint doble clic
        tk.Label(card_tbl,
                 text="  Doble clic en una fila verde para superponer en AutoCAD",
                 font=FONT_SMALL, bg=C["panel"], fg=C["txt_dim"]
                 ).grid(row=1, column=0, columnspan=2, sticky="w", pady=(4, 0))

        # ─── TARJETA: consola ────────────────────────────────────────────────
        outer_log = tk.Frame(body, bg=C["border"], padx=1, pady=1)
        outer_log.pack(fill="both", expand=True)

        log_hdr = tk.Frame(outer_log, bg=C["panel2"], padx=12, pady=6)
        log_hdr.pack(fill="x")
        log_hdr.columnconfigure(1, weight=1)

        tk.Label(log_hdr, text="CONSOLA",
                 font=("Segoe UI", 8, "bold"),
                 bg=C["panel2"], fg=C["txt_mid"]
                 ).grid(row=0, column=0, sticky="w")

        tk.Button(log_hdr,
                  text="Limpiar",
                  bg=C["border"], fg=C["txt_dim"],
                  relief="flat", font=FONT_SMALL,
                  cursor="hand2",
                  activebackground=C["log_err"], activeforeground="white",
                  command=self._limpiar_log
                  ).grid(row=0, column=2, sticky="e")

        card_log = tk.Frame(outer_log, bg=C["log_bg"], padx=10, pady=8)
        card_log.pack(fill="both", expand=True)
        card_log.columnconfigure(0, weight=1)
        card_log.rowconfigure(0, weight=1)

        self._log_w = tk.Text(card_log, bg=C["log_bg"], fg=C["txt"],
                               font=FONT_LOG, relief="flat", state="disabled",
                               wrap="word", bd=0)
        self._log_w.grid(row=0, column=0, sticky="nsew")

        for tag, color in [
            ("ok",   C["log_ok"]),
            ("warn", C["log_warn"]),
            ("err",  C["log_err"]),
            ("dim",  C["log_dim"]),
            ("ts",   C["txt_dim"]),
        ]:
            self._log_w.tag_config(tag, foreground=color)

        sb2 = ttk.Scrollbar(card_log, orient="vertical", command=self._log_w.yview)
        sb2.grid(row=0, column=1, sticky="ns")
        self._log_w.configure(yscrollcommand=sb2.set)

    # ── Widgets helpers ───────────────────────────────────────────────────────

    def _card(self, parent, title=""):
        """Retorna (outer, inner): outer se coloca con pack/grid, inner recibe widgets."""
        outer = tk.Frame(parent, bg=C["border"], padx=1, pady=1)
        if title:
            hdr = tk.Frame(outer, bg=C["panel2"], padx=12, pady=5)
            hdr.pack(fill="x")
            tk.Label(hdr, text=title, font=("Segoe UI", 8, "bold"),
                     bg=C["panel2"], fg=C["txt_mid"]).pack(anchor="w")
        inner = tk.Frame(outer, bg=C["panel"], padx=12, pady=10)
        inner.pack(fill="both", expand=True)
        return outer, inner

    def _lbl_field(self, parent, text, row):
        tk.Label(parent, text=text, font=("Segoe UI", 9, "bold"),
                 bg=C["panel"], fg=C["txt_mid"], anchor="w"
                 ).grid(row=row, column=0, columnspan=3, sticky="w", pady=(8, 0))

    def _explore_btn(self, parent, cmd):
        return tk.Button(parent, text="Explorar...",
                         bg=C["border"], fg=C["accent"], relief="flat",
                         font=FONT_SMALL, cursor="hand2",
                         activebackground=C["accent"], activeforeground=C["bg"],
                         command=cmd)
    
    # ── Helpers ───────────────────────────────────────────────────────────────

    def _centrar(self, w, h):
        x = (self.winfo_screenwidth()  - w) // 2
        y = (self.winfo_screenheight() - h) // 2
        self.geometry(f"{w}x{h}+{x}+{y}")

    def _log(self, msg: str, tag: str = ""):
        self._log_w.configure(state="normal")
        ts = time.strftime("%H:%M:%S")
        self._log_w.insert("end", f"{ts}  ", "ts")
        self._log_w.insert("end", f"{msg}\n", tag or "")
        self._log_w.see("end")
        self._log_w.configure(state="disabled")

    def _limpiar_log(self):
        self._log_w.configure(state="normal")
        self._log_w.delete("1.0", "end")
        self._log_w.configure(state="disabled")

    def _busy(self, activo: bool):
        self._btn_extraer.configure_state(not activo)
        self._btn_crear.configure_state(not activo)
        self._btn_comprobar.configure_state(not activo)
        self._btn_todo.configure_state(not activo)
        if activo:
            self._dot_count = 0
            self.configure(cursor="watch")
            self._pulse.start()
            self._animar_status()
        else:
            self.configure(cursor="")
            self._pulse.stop()
        self.update_idletasks()

    def _animar_status(self):
        # Solo anima si los botones están deshabilitados (= procesando)
        if getattr(self._btn_extraer, "_enabled", True):
            return
        self._dot_count += 1
        self.after(300, self._animar_status)

    def _on_dwg_changed(self, *_):
        """Actualiza el badge del header cuando cambia el campo DWG."""
        dwg = self._dwg_plano.get().strip().strip('"')
        if dwg:
            nombre = os.path.basename(dwg)
            self._badge_lbl.configure(text=nombre, fg=C["accent"])
            self._badge_frame.configure(bg=C["panel2"])
            self._badge_lbl.configure(bg=C["panel2"])
        else:
            self._badge_lbl.configure(text="— sin seleccionar —", fg=C["txt_dim"])

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

    # ── CREAR ARTE ────────────────────────────────────────────────────────────

    def _crear_arte(self):
        dwg = self._dwg_plano.get().strip().strip('"')
        if not dwg:
            messagebox.showwarning("Campo requerido",
                                   "Selecciona el archivo DWG del plano extraido (_PLANO.dwg).")
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
            self._log("Arte completado.", "ok")
        except RuntimeError as e:
            self._log(str(e), "err")
            self.after(0, lambda: messagebox.showerror("Error", str(e)))
        except Exception as e:
            self._log(f"ERROR inesperado: {e}", "err")
        finally:
            self._busy(False)

    # ── TODO EN UNO ───────────────────────────────────────────────────────────

    def _todo_en_uno(self):
        if not self._validar(necesita_dwg=True):
            return

        # Mostrar diálogo del cajetín PRIMERO en el hilo principal
        ruta_plano   = self._dwg_plano.get().strip().strip('"')
        nombre_plano = os.path.splitext(os.path.basename(ruta_plano))[0]

        if _CAJETIN_DIALOG_OK:
            valores = _dialogo_cajetin(nombre_plano)
        else:
            valores = {}

        if valores is None:
            return  # usuario canceló

        self._busy(True)
        threading.Thread(target=self._t_todo_en_uno,
                         args=(ruta_plano, valores), daemon=True).start()

    def _t_todo_en_uno(self, ruta_plano: str, valores: dict):
        malla = valores.get("MALLA", "").strip()
        pieza = valores.get("PIEZA", "").strip()

        self._log("=" * 56)
        self._log("TODO EN UNO — Extrayendo plano + Creando arte...", "ok")
        self._log(f"Plano: {os.path.basename(ruta_plano)}", "dim")
        if malla or pieza:
            self._log(f"Malla: {malla}   Pieza: {pieza}", "dim")

        # ─ Paso 1: Extraer plano ─────────────────────────────────────────────
        self._log("─" * 40)
        self._log("PASO 1/2 — Extrayendo layers...", "warn")
        nombre_base   = os.path.splitext(os.path.basename(ruta_plano))[0]
        ruta_filtrada = os.path.join(_ruta_planos(ruta_plano), f"{nombre_base}_PLANO.dwg")
        self._log(f"  → {ruta_filtrada}", "dim")

        try:
            motor = AutoCADMotor()
        except RuntimeError as e:
            self._log(f"ERROR AutoCAD: {e}", "err")
            self._busy(False)
            return
        try:
            motor.extraer_layers(ruta_plano, ruta_filtrada,
                                 log_fn=lambda m: self._log(m, "dim"))
        except Exception as e:
            self._log(f"ERROR extraccion: {e}", "err")
            motor.quit()
            self._busy(False)
            return
        motor.quit()
        time.sleep(2.0)   # dar tiempo a AutoCAD para cerrar completamente el doc
        self._log(f"  Plano extraido ✔  {os.path.basename(ruta_filtrada)}", "ok")

        # ─ Paso 2: Crear arte — detectar cuántas piezas hay ──────────────────
        self._log("─" * 40)
        self._log("PASO 2/2 — Creando arte en AutoCAD...", "warn")

        import shutil, traceback as _tb

        # Procesamos pieza 0 primero para descubrir n_piezas
        ruta_arte_0 = _ruta_arte_salida(ruta_plano, malla, pieza)
        self._log(f"  → {ruta_arte_0}", "dim")

        try:
            n_piezas = _crear_arte_autocad(
                ruta_filtrada,
                log_fn=lambda m: self._log(m, "dim"),
                valores_cajetin=valores if valores else None,
                ruta_salida=ruta_arte_0,
                perim_index=0)
            self._log(f"Arte pieza 1 guardado ✔  {os.path.basename(ruta_arte_0)}", "ok")
        except Exception as e:
            self._log(f"ERROR en creacion de arte: {e}", "err")
            self._log(_tb.format_exc(), "err")
            self.after(0, lambda: messagebox.showerror("Error creando arte", str(e)))
            self._busy(False)
            return

        # Si hay más piezas: copiar el plano extraído y procesar cada una
        artes_creados = [ruta_arte_0]
        for i in range(1, n_piezas):
            self._log(f"─" * 40)
            self._log(f"Procesando pieza {i+1}/{n_piezas}...", "warn")
            pieza_sufijo = f"{pieza} {i+1}".strip() if pieza else str(i+1)
            ruta_copia = ruta_filtrada.replace("_PLANO.dwg", f"_PLANO_p{i+1}.dwg")
            try:
                shutil.copy2(ruta_filtrada, ruta_copia)
                ruta_arte_i = _ruta_arte_salida(ruta_plano, malla, pieza_sufijo)
                self._log(f"  → {ruta_arte_i}", "dim")
                _crear_arte_autocad(
                    ruta_copia,
                    log_fn=lambda m: self._log(m, "dim"),
                    valores_cajetin=valores if valores else None,
                    ruta_salida=ruta_arte_i,
                    perim_index=i)
                self._log(f"Arte pieza {i+1} guardado ✔  {os.path.basename(ruta_arte_i)}", "ok")
                artes_creados.append(ruta_arte_i)
            except Exception as e:
                self._log(f"WARN pieza {i+1}: {e}", "warn")

        self._log("=" * 56)
        self._log(f"TODO EN UNO completado — {len(artes_creados)} arte(s).", "ok")
        self._log(f"  Plano  → {ruta_filtrada}", "ok")
        for a in artes_creados:
            self._log(f"  Arte   → {a}", "ok")

        import subprocess
        self.after(0, lambda: subprocess.Popen(["explorer", os.path.dirname(ruta_arte_0)]))
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
        ruta_destino  = _ruta_planos(ruta_plano)
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
            self._log(f"ERROR extraccion: {e}", "err")
            motor.quit()
            self._busy(False)
            return

        motor.quit()
        self._log("Extraccion completada.", "ok")
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

    # ── BUSCAR ARTE ───────────────────────────────────────────────────────────

    def _comprobar(self):
        if not self._validar(necesita_dwg=False):
            return
        self._busy(True)
        threading.Thread(target=self._t_comprobar, daemon=True).start()

    def _t_comprobar(self):
        ruta_base = self._ruta_base.get().strip()
        dwg_plano = self._dwg_plano.get().strip().strip('"')

        self._log("=" * 56)
        self._log("BUSCAR ARTE — buscando artes...", "ok")
        self._log(f"Buscando en: {ruta_base}", "dim")

        codigos = _extraer_codigos(dwg_plano) if dwg_plano else []
        if codigos:
            self._log(f'Codigos buscados: {" | ".join(codigos)}', "dim")
        else:
            self._log("Sin codigo de plano — se mostraran todos los artes.", "warn")

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
            ext = os.path.splitext(r["archivo"])[1].upper().lstrip(".")
            tipo = ext if ext else "?"
            self._tree.insert("", "end",
                              values=("OK  COINCIDE", r["version"], r["archivo"], tipo),
                              tags=("match",))

        # Actualizar header count
        n = len(resultados)
        self._lbl_tbl_titulo.configure(
            text=f"ARTES ENCONTRADOS — {n} resultado{'s' if n != 1 else ''}")

        # Actualizar badge de estado
        if n > 0:
            self._badge_estado.configure(
                text=f"OK  EXISTE ({n})",
                fg=C["log_ok"],
                bg=C["panel2"])
        elif self._dwg_plano.get().strip():
            self._badge_estado.configure(
                text="X  NO ENCONTRADO",
                fg=C["log_err"],
                bg=C["panel2"])
        else:
            self._badge_estado.configure(
                text="— busca primero",
                fg=C["txt_dim"],
                bg=C["panel2"])

    def _abrir_carpeta_arte(self):
        """Abre el explorador en la carpeta del arte seleccionado."""
        sel = self._tree.selection()
        if not sel:
            # Sin selección, intentar con el primer resultado
            if not self._resultados:
                messagebox.showinfo("Sin seleccion",
                                    "Selecciona un arte de la tabla primero.")
                return
            r = self._resultados[0]
        else:
            idx = self._tree.index(sel[0])
            if idx >= len(self._resultados):
                return
            r = self._resultados[idx]

        carpeta = os.path.dirname(r["ruta_completa"])
        import subprocess
        try:
            subprocess.Popen(["explorer", carpeta])
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo abrir el explorador:\n{e}")

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
                f"La superposicion requiere un archivo DWG.\n{r['archivo']}")
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
            self._log("Superposicion lista en AutoCAD.", "ok")
            self._log(
                "Si el perimetro del plano (XREF) coincide con el arte → OK correcto.", "ok")
        except RuntimeError as e:
            self._log(str(e), "err")
        except Exception as e:
            self._log(f"ERROR: {e}", "err")
        finally:
            self._busy(False)


# ─────────────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    ArteMakerApp().mainloop()
