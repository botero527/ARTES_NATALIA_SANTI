"""
Genera el script Python que se ejecuta DENTRO de Rhino 8 (rhinoscriptsyntax)
y lanza Rhino para que lo procese.
"""
import os
import sys
import subprocess
from string import Template

from config import (
    RHINO_EXE, CAJETINES_DWG,
    LAYER_PLANES, LAYER_K2, LAYER_K,
    OFFSET_PERIMETRO, OFFSET_BN_DEGRADE, DIVISOR_DEGRADE,
    LOGO_MARGEN_1, LOGO_MARGEN_2,
    NOMBRE_BLOQUE_25,
)

# ─────────────────────────────────────────────────────────────────────────────
# TEMPLATE DEL SCRIPT QUE CORRE DENTRO DE RHINO
# Usa $variable para los parámetros que se inyectan desde Python externo.
# ─────────────────────────────────────────────────────────────────────────────

_SCRIPT_BASE = r'''# -*- coding: utf-8 -*-
# Script generado por AGP Arte Maker — NO editar manualmente
import rhinoscriptsyntax as rs  # type: ignore
import Rhino  # type: ignore
import scriptcontext as sc  # type: ignore
import math
import os

# ── parámetros inyectados ────────────────────────────────────────────────────
DWG_PLANO       = r"$dwg_plano"
DWG_CAJETIN     = r"$dwg_cajetin"
RUTA_SALIDA     = r"$ruta_salida"
LAYER_PLANES    = "$layer_planes"
LAYER_K2        = "$layer_k2"
LAYER_K         = "$layer_k"
OFFSET_PERIM    = $offset_perimetro
OFFSET_BN_DEG   = $offset_bn_degrade
DIVISOR_DEG     = $divisor_degrade
LOGO_M1         = $logo_margen_1
LOGO_M2         = $logo_margen_2
BLOQUE_25       = "$nombre_bloque_25"

LOG_FILE        = r"$log_file"

# ── helpers ──────────────────────────────────────────────────────────────────

def _log(msg):
    print(msg)
    try:
        with open(LOG_FILE, "a") as f:
            f.write(str(msg) + "\n")
    except Exception:
        pass


def asegurar_layer(nombre, color_rgb=None):
    if not rs.IsLayer(nombre):
        rs.AddLayer(nombre)
    if color_rgb:
        import System.Drawing  # type: ignore
        c = System.Drawing.Color.FromArgb(*color_rgb)
        rs.LayerColor(nombre, c)


def objetos_en_patron(patron):
    """Devuelve todos los objetos cuyo layer contenga el patrón (case-insensitive)."""
    resultado = []
    layers = rs.LayerNames()
    if layers is None:
        return resultado
    for lname in layers:
        if patron.upper() in lname.upper():
            objs = rs.ObjectsByLayer(lname)
            if objs:
                resultado.extend(objs)
    return resultado


def primera_curva_cerrada_en_patron(patron):
    """Primera curva cerrada en layers que contengan el patrón."""
    for obj in objetos_en_patron(patron):
        if rs.IsCurve(obj) and rs.IsCurveClosed(obj):
            return obj
    return None


def hatch_solido(curvas_ids, layer_nombre):
    """
    Crea hatch sólido usando las curvas dadas como borde.
    Si son dos curvas (una dentro de otra) crea el anillo.
    """
    doc = sc.doc
    solid_idx = doc.HatchPatterns.Find("Solid", True)
    if solid_idx < 0:
        solid_idx = 0

    curvas_geo = []
    for cid in curvas_ids:
        crv = rs.coercecurve(cid)
        if crv:
            curvas_geo.append(crv)

    if not curvas_geo:
        return

    hatches = Rhino.Geometry.Hatch.Create(curvas_geo, solid_idx, 0.0, 1.0)
    if not hatches:
        _log("  WARN: no se genero hatch para layer {}".format(layer_nombre))
        return

    layer_idx = doc.Layers.FindByFullPath(layer_nombre, True)
    if layer_idx < 0:
        layer_idx = rs.LayerIndex(layer_nombre, True)

    for h in hatches:
        attrs = Rhino.DocObjects.ObjectAttributes()
        attrs.LayerIndex = max(layer_idx, 0)
        doc.Objects.AddHatch(h, attrs)

    doc.Views.Redraw()


def distribuir_bloque_freeform(bloque_id, curva_id, n_items):
    """
    Copia n_items veces el bloque a lo largo de la curva con orientacion libre.
    Retorna lista de nuevos IDs.
    """
    if n_items < 1:
        return []
    longitud = rs.CurveLength(curva_id)
    if not longitud or longitud == 0:
        return []

    espaciado = longitud / n_items

    # Punto de referencia = centro del bounding box del bloque
    bbox = rs.BoundingBox([bloque_id])
    if not bbox:
        return []
    ref_x = (bbox[0][0] + bbox[6][0]) / 2.0
    ref_y = (bbox[0][1] + bbox[6][1]) / 2.0
    ref_pt = [ref_x, ref_y, 0.0]

    nuevos = []
    for i in range(n_items):
        distancia = i * espaciado
        # CurveArcLengthPoint devuelve Point3d directamente (no un parametro t)
        punto = rs.CurveArcLengthPoint(curva_id, distancia)
        if punto is None:
            continue
        # Obtener el parametro t en ese punto para calcular la tangente
        t = rs.CurveClosestPoint(curva_id, punto)
        if t is None:
            continue
        tangente = rs.CurveTangent(curva_id, t)

        traslacion = [
            punto[0] - ref_pt[0],
            punto[1] - ref_pt[1],
            0.0,
        ]
        nuevo_id = rs.CopyObject(bloque_id, traslacion)
        if nuevo_id:
            angulo = math.degrees(math.atan2(tangente[1], tangente[0]))
            rs.RotateObject(nuevo_id, rs.coerce3dpoint(punto), angulo)
            nuevos.append(nuevo_id)

    return nuevos


def centrar_objeto_en_bbox(obj_id, referencia_ids):
    """Centra obj_id sobre el bounding box de referencia_ids."""
    bbox_ref = rs.BoundingBox(referencia_ids)
    bbox_obj = rs.BoundingBox([obj_id])
    if not bbox_ref or not bbox_obj:
        return
    cx_ref = (bbox_ref[0][0] + bbox_ref[6][0]) / 2.0
    cy_ref = (bbox_ref[0][1] + bbox_ref[6][1]) / 2.0
    cx_obj = (bbox_obj[0][0] + bbox_obj[6][0]) / 2.0
    cy_obj = (bbox_obj[0][1] + bbox_obj[6][1]) / 2.0
    rs.MoveObject(obj_id, [cx_ref - cx_obj, cy_ref - cy_obj, 0.0])


# ── script principal ─────────────────────────────────────────────────────────

def main():
    rs.EnableRedraw(False)
    _log("=== Arte Maker: iniciando script Rhino ===")

    # 1b. Detectar geometria del plano ANTES de importar cajetines
    #     (los cajetines tienen layers "BN INT" que confunden la deteccion)
    _log("  [1b] Detectando geometria del plano...")

    perim_id = primera_curva_cerrada_en_patron("PERIMETRO")
    if perim_id is None:
        _log("  ERROR: No se encontro curva de PERIMETRO.")
        _all_layers = rs.LayerNames() or []
        _log("  Layers en doc: {}".format(", ".join(str(_l) for _l in _all_layers)))
        rs.EnableRedraw(True)
        return

    bn_id = (
        primera_curva_cerrada_en_patron("BN")
        or primera_curva_cerrada_en_patron("PHANTOM")
        or primera_curva_cerrada_en_patron("BANDA")
    )

    logo_objs = objetos_en_patron("LOGO") + objetos_en_patron("TRAZABILIDAD")

    _bn_cerradas = set()
    for _pat in ["BN", "PHANTOM", "BANDA NEGRA", "BANDANEGRA", "BANDA"]:
        for _obj in objetos_en_patron(_pat):
            if rs.IsCurve(_obj) and rs.IsCurveClosed(_obj):
                _bn_cerradas.add(_obj)
    CON_DEGRADE = len(_bn_cerradas) >= 2
    _log("  --> {} curvas cerradas BN: {}".format(
        len(_bn_cerradas), "CON degrade" if CON_DEGRADE else "SIN degrade"))

    # 1. Importar cajetines DESPUES de detectar (para no contaminar la deteccion)
    _log("  [1/7] Importando cajetines...")
    rs.Command('! _-Import "{}" _Enter _Enter'.format(DWG_CAJETIN), False)

    # 2. Crear layers de arte (los objetos originales NO se mueven) ───────────
    # Sin color forzado: si el layer ya existe (vino del cajetin importado)
    # conserva su color original. Solo se crea si no existe.
    _log("  [2/7] Preparando layers de arte...")
    for lyr in [LAYER_PLANES, LAYER_K2, LAYER_K]:
        asegurar_layer(lyr)

    # 4. Offset del perímetro (0.5 mm) ───────────────────────────────────────
    _log("  [4/8] Offset perimetro {:.1f} mm...".format(OFFSET_PERIM))
    _bbox_p    = rs.BoundingBox([perim_id])
    _inside_pt = rs.coerce3dpoint([
        (_bbox_p[0][0] + _bbox_p[6][0]) / 2.0,
        (_bbox_p[0][1] + _bbox_p[6][1]) / 2.0,
        0.0
    ])
    offset_ids = rs.OffsetCurve(perim_id, _inside_pt, OFFSET_PERIM)
    if not offset_ids:
        _log("  ERROR: No se pudo crear offset del perímetro")
        rs.EnableRedraw(True)
        return
    offset_perim_id = offset_ids[0] if isinstance(offset_ids, list) else offset_ids
    _area_orig = rs.CurveArea(perim_id)
    _area_off  = rs.CurveArea(offset_perim_id)
    if _area_orig and _area_off and _area_off[0] > _area_orig[0]:
        _log("  corrigiendo direccion offset perimetro...")
        rs.DeleteObject(offset_perim_id)
        _rev = rs.CopyObject(perim_id)
        rs.ReverseCurve(_rev)
        offset_ids = rs.OffsetCurve(_rev, _inside_pt, OFFSET_PERIM)
        rs.DeleteObject(_rev)
        if not offset_ids:
            _log("  ERROR: No se pudo crear offset del perímetro")
            rs.EnableRedraw(True)
            return
        offset_perim_id = offset_ids[0] if isinstance(offset_ids, list) else offset_ids
    rs.ObjectLayer(offset_perim_id, LAYER_PLANES)

    # 5. Hatch entre perímetro y offset → k2 ─────────────────────────────────
    _log("  [5/8] Hatch k2 (borde perímetro)...")
    hatch_solido([perim_id, offset_perim_id], LAYER_K2)

    # 6. Hatch entre BN y offset → k ─────────────────────────────────────────
    _log("  [6/8] Hatch k (banda negra)...")
    if bn_id:
        hatch_solido([bn_id, offset_perim_id], LAYER_K)
    else:
        _log("  WARN: No se encontró curva de banda negra")

    # 7. Recuadro del logo ────────────────────────────────────────────────────
    _log("  [7/8] Construyendo recuadro de logo...")
    if logo_objs:
        bbox = rs.BoundingBox(logo_objs)
        if bbox:
            xmin = bbox[0][0]
            xmax = bbox[6][0]
            ymin = min(bbox[0][1], bbox[6][1])
            # y1 = offset 3 mm bajo el logo
            # y2 = offset 3.8 mm bajo y1 (offset del offset)
            y1   = ymin - LOGO_M1
            y2   = y1   - LOGO_M2
            pts  = [[xmin, y1, 0], [xmax, y1, 0], [xmax, y2, 0], [xmin, y2, 0]]
            lineas = [
                rs.AddLine(pts[0], pts[1]),
                rs.AddLine(pts[1], pts[2]),
                rs.AddLine(pts[2], pts[3]),
                rs.AddLine(pts[3], pts[0]),
            ]
            joined = rs.JoinCurves(lineas, True)
            for j in (joined if joined else []):
                rs.ObjectLayer(j, LAYER_PLANES)
    else:
        _log("  WARN: No se encontraron objetos de logo")

    # 7b. Centrar cajetin sobre area de logo/trazabilidad ───────────────────
    _log("  [7b] Centrando cajetin sobre logo...")
    cajetin_objs = objetos_en_patron("CAJETIN 1")
    if not cajetin_objs:
        cajetin_objs = objetos_en_patron("CAJETIN")
    if cajetin_objs and logo_objs:
        bbox_logo = rs.BoundingBox(logo_objs)
        bbox_caj  = rs.BoundingBox(cajetin_objs)
        if bbox_logo and bbox_caj:
            cx_logo = (bbox_logo[0][0] + bbox_logo[6][0]) / 2.0
            cy_logo = (bbox_logo[0][1] + bbox_logo[6][1]) / 2.0
            cx_caj  = (bbox_caj[0][0]  + bbox_caj[6][0])  / 2.0
            cy_caj  = (bbox_caj[0][1]  + bbox_caj[6][1])  / 2.0
            rs.MoveObjects(cajetin_objs, [cx_logo - cx_caj, cy_logo - cy_caj, 0.0])
            _log("  Cajetin centrado.")
    else:
        _log("  WARN: no se encontro cajetin o logo para centrar.")

    # 8. Degradé ─────────────────────────────────────────────────────────────
    if CON_DEGRADE:
        _log("  [8/8] Degradé: offset BN interior...")
        if bn_id:
            # 8a. Offset de BN hacia adentro
            _bbox_bn   = rs.BoundingBox([bn_id])
            _bn_inside = rs.coerce3dpoint([
                (_bbox_bn[0][0] + _bbox_bn[6][0]) / 2.0,
                (_bbox_bn[0][1] + _bbox_bn[6][1]) / 2.0,
                0.0
            ])
            off_bn_ids = rs.OffsetCurve(bn_id, _bn_inside, OFFSET_BN_DEG)
            if off_bn_ids:
                _off_bn_tmp  = off_bn_ids[0] if isinstance(off_bn_ids, list) else off_bn_ids
                _area_bn     = rs.CurveArea(bn_id)
                _area_off_bn = rs.CurveArea(_off_bn_tmp)
                if _area_bn and _area_off_bn and _area_off_bn[0] > _area_bn[0]:
                    _log("  corrigiendo direccion offset BN...")
                    rs.DeleteObject(_off_bn_tmp)
                    _rev_bn = rs.CopyObject(bn_id)
                    rs.ReverseCurve(_rev_bn)
                    off_bn_ids = rs.OffsetCurve(_rev_bn, _bn_inside, OFFSET_BN_DEG)
                    rs.DeleteObject(_rev_bn)
            if off_bn_ids:
                off_bn_id = off_bn_ids[0] if isinstance(off_bn_ids, list) else off_bn_ids
                rs.ObjectLayer(off_bn_id, LAYER_PLANES)

                # 8b. Calcular número de pepas
                longitud  = rs.CurveLength(off_bn_id) or 0
                n_pepas   = int(round(longitud / DIVISOR_DEG)) if longitud > 0 else 0
                _log("     longitud: {:.2f} mm  pepas: {}".format(longitud, n_pepas))

                # 8c. Buscar bloque 25 en documento
                bloque_id = None
                for obj in rs.AllObjects() or []:
                    try:
                        if rs.IsBlockInstance(obj):
                            if rs.BlockInstanceName(obj) == BLOQUE_25:
                                bloque_id = obj
                                break
                    except Exception:
                        pass

                if bloque_id and n_pepas > 0:
                    _log("     Distribuyendo bloque '{}' x{}...".format(BLOQUE_25, n_pepas))
                    nuevos = distribuir_bloque_freeform(bloque_id, off_bn_id, n_pepas)
                else:
                    _log("  WARN: bloque '{}' no encontrado o n_pepas=0".format(BLOQUE_25))
            else:
                _log("  WARN: No se pudo crear offset interior de BN")

    else:
        _log("  [8/8] Sin degradé — omitido")

    # Guardar ────────────────────────────────────────────────────────────────
    if RUTA_SALIDA and RUTA_SALIDA != "":
        _log("  Guardando en: {}".format(RUTA_SALIDA))
        rs.Command('_-Save "{}" _Enter'.format(RUTA_SALIDA), False)

    rs.EnableRedraw(True)
    doc = sc.doc
    if doc:
        doc.Views.Redraw()

    _log("=== Arte Maker: COMPLETADO ===")


main()
'''

_TEMPLATE = Template(_SCRIPT_BASE)


# ─── función pública ──────────────────────────────────────────────────────────

def generar_y_ejecutar(
    dwg_plano: str,
    ruta_salida: str = "",
    log_fn=None,
) -> str:
    """
    Genera el script de Rhino, lo guarda en una ruta fija y abre Rhino.
    Devuelve la ruta del script generado (siempre la misma).
    """
    if log_fn is None:
        log_fn = print

    # Ruta FIJA — el usuario configura Rhino una sola vez apuntando aquí
    _DIR     = os.path.dirname(os.path.abspath(__file__))
    script_p = os.path.join(_DIR, "arte_script.py")
    log_p    = os.path.join(_DIR, "arte_log.txt")

    cajetines = os.path.abspath(CAJETINES_DWG)
    if not os.path.isfile(cajetines):
        raise FileNotFoundError(
            f"No se encontró el archivo de cajetines:\n{cajetines}\n"
            "Ajusta CAJETINES_DWG en config.py"
        )

    contenido = _TEMPLATE.substitute(
        dwg_plano        = dwg_plano.replace("\\", "\\\\"),
        dwg_cajetin      = cajetines.replace("\\", "\\\\"),
        ruta_salida      = ruta_salida.replace("\\", "\\\\"),
        layer_planes     = LAYER_PLANES,
        layer_k2         = LAYER_K2,
        layer_k          = LAYER_K,
        offset_perimetro = OFFSET_PERIMETRO,
        offset_bn_degrade= OFFSET_BN_DEGRADE,
        divisor_degrade  = DIVISOR_DEGRADE,
        logo_margen_1    = LOGO_MARGEN_1,
        logo_margen_2    = LOGO_MARGEN_2,
        nombre_bloque_25 = NOMBRE_BLOQUE_25,
        log_file         = log_p.replace("\\", "\\\\"),
    )

    with open(script_p, "w", encoding="utf-8") as f:
        f.write(contenido)

    log_fn(f"  Script Rhino generado: {script_p}")

    if not os.path.isfile(RHINO_EXE):
        raise FileNotFoundError(
            f"No se encontró Rhino 8 en:\n{RHINO_EXE}\n"
            "Ajusta RHINO_EXE en config.py"
        )

    log_fn(f"  Script listo para ejecutar en Rhino.")
    log_fn(f"  Ruta del script: {script_p}")

    # Intentar abrir Rhino si no está abierto
    _abrir_rhino_si_cerrado(log_fn)

    return log_p


def _abrir_rhino_si_cerrado(log_fn):
    """Abre Rhino si no hay ninguna instancia corriendo."""
    import win32gui

    rhino_abierto = False

    def _check(hwnd, _):
        nonlocal rhino_abierto
        if win32gui.IsWindowVisible(hwnd):
            titulo = win32gui.GetWindowText(hwnd)
            if "Rhinoceros" in titulo or "Rhino 8" in titulo:
                rhino_abierto = True

    try:
        win32gui.EnumWindows(_check, None)
    except Exception:
        pass

    if not rhino_abierto:
        log_fn("  Abriendo Rhino 8...")
        subprocess.Popen([RHINO_EXE, "/nosplash"])
    else:
        log_fn("  Rhino ya estaba abierto.")
