# -*- coding: utf-8 -*-
# Script generado por AGP Arte Maker — NO editar manualmente
import rhinoscriptsyntax as rs  # type: ignore
import Rhino  # type: ignore
import scriptcontext as sc  # type: ignore
import math
import os

# ── parámetros inyectados ────────────────────────────────────────────────────
DWG_PLANO       = r"\\\\192.168.2.37\\ingenieria\\PRODUCCION\\AGP PLANOS TECNICOS\\MBZ\\MBZ GLC 4D COUPE 2024\\V-08  AUTO SAFE\\ARTES\\1708 008 030 A_PLANO.dwg"
DWG_CAJETIN     = r"c:\\Users\\abotero\\OneDrive - AGP GROUP\\Documentos\\macro_natalia\\LAYERS Y CAJETINES 1.dwg"
RUTA_SALIDA     = r"c:\\Users\\abotero\\OneDrive - AGP GROUP\\Documentos\\macro_natalia\\1708 008 030 A_ARTE.3dm"
LAYER_PLANES    = "PLANES"
LAYER_K2        = "k2"
LAYER_K         = "k"
OFFSET_PERIM    = 0.5
OFFSET_BN_DEG   = 2.5
DIVISOR_DEG     = 3
LOGO_M1         = 3.0
LOGO_M2         = 3.8
BLOQUE_25       = "25"

LOG_FILE        = r"c:\\Users\\abotero\\OneDrive - AGP GROUP\\Documentos\\macro_natalia\\arte_log.txt"

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


def offset_hacia_adentro(curva_id, distancia, layer_destino):
    """
    Crea un offset de la curva SIEMPRE hacia adentro usando RhinoCommon.
    Prueba distancia negativa y positiva, queda con la que tenga area menor.
    Retorna el ID del objeto creado, o None si falla.
    """
    import Rhino.Geometry as rg  # type: ignore
    crv = rs.coercecurve(curva_id)
    if not crv:
        return None
    plane  = rg.Plane.WorldXY
    tol    = sc.doc.ModelAbsoluteTolerance

    area_orig = rs.CurveArea(curva_id)
    area_orig = area_orig[0] if area_orig else 1e18

    mejor_id   = None
    mejor_area = 1e18

    for signo in (-1, 1):
        resultados = crv.Offset(plane, signo * distancia, tol,
                                rg.CurveOffsetCornerStyle.Sharp)
        if not resultados:
            continue
        # Unir segmentos si el offset devolvio multiples curvas
        if len(resultados) > 1:
            joined = rg.Curve.JoinCurves(resultados, tol)
            crvs_a_usar = list(joined) if joined else list(resultados)
        else:
            crvs_a_usar = list(resultados)

        for c in crvs_a_usar:
            if not c.IsClosed:
                continue
            area_mp = c.GetArea() if hasattr(c, 'GetArea') else None
            if area_mp is None:
                mp = rg.AreaMassProperties.Compute(c)
                area_mp = mp.Area if mp else 1e18
            if area_mp < area_orig and area_mp < mejor_area:
                mejor_area = area_mp
                # Agregar al doc
                if mejor_id:
                    rs.DeleteObject(mejor_id)
                attr = Rhino.DocObjects.ObjectAttributes()
                attr.LayerIndex = sc.doc.Layers.FindByFullPath(layer_destino, True)
                if attr.LayerIndex < 0:
                    attr.LayerIndex = 0
                mejor_id = sc.doc.Objects.AddCurve(c, attr)

    return mejor_id


def distribuir_bloque_freeform(bloque_id, curva_id, n_items):
    """
    Copia n_items veces el bloque a lo largo de la curva con orientacion libre.
    Los puntos grandes siempre quedan hacia afuera (hacia el hatch).
    Retorna lista de nuevos IDs.
    """
    if n_items < 1:
        return []
    longitud = rs.CurveLength(curva_id)
    if not longitud or longitud == 0:
        return []

    espaciado = longitud / n_items

    # Orientacion de la curva: CCW=1, CW=-1
    # CCW: el exterior esta a la derecha de la tangente → sumar 180
    # CW:  el exterior esta a la izquierda → no sumar nada
    orientacion = rs.ClosedCurveOrientation(curva_id)
    angulo_extra = 0 if (orientacion is not None and orientacion > 0) else 180

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
        punto = rs.CurveArcLengthPoint(curva_id, distancia)
        if punto is None:
            continue
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
            angulo = math.degrees(math.atan2(tangente[1], tangente[0])) + angulo_extra
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

    # Recolectar TODAS las curvas cerradas de BN/PHANTOM ordenadas por area
    # La de mayor area = exterior (mas cerca del perimetro) = bn_id
    _bn_lista = []
    for _pat in ["BN", "PHANTOM", "BANDA NEGRA", "BANDANEGRA", "BANDA"]:
        for _obj in objetos_en_patron(_pat):
            if rs.IsCurve(_obj) and rs.IsCurveClosed(_obj) and _obj not in _bn_lista:
                _bn_lista.append(_obj)

    def _area_curva(cid):
        a = rs.CurveArea(cid)
        return a[0] if a else 0

    _bn_lista.sort(key=_area_curva, reverse=True)  # mayor area primero = exterior

    CON_DEGRADE = len(_bn_lista) >= 2
    bn_id = _bn_lista[0] if _bn_lista else None  # BN exterior (primera linea)
    _log("  --> {} curvas cerradas BN: {}".format(
        len(_bn_lista), "CON degrade" if CON_DEGRADE else "SIN degrade"))

    # Capturar objetos del layer exactamente "LOGO" del plano ANTES de importar
    _logo_plano_objs = [o for o in (rs.AllObjects() or [])
                        if rs.IsObject(o)
                        and rs.ObjectLayer(o).upper().endswith("LOGO")]


    # 1. Importar cajetines DESPUES de detectar (para no contaminar la deteccion)
    _log("  [1/7] Importando cajetines...")
    _ids_antes = set(rs.AllObjects() or [])
    rs.Command('! _-Import "{}" _Enter _Enter'.format(DWG_CAJETIN), False)
    _ids_cajetin = set(rs.AllObjects() or []) - _ids_antes  # objetos nuevos del import

    # 2. Crear layers de arte (los objetos originales NO se mueven) ──────────
    # Sin color forzado: si el layer ya existe (vino del cajetin importado)
    # conserva su color original. Solo se crea si no existe.
    _log("  [2/7] Preparando layers de arte...")
    for lyr in [LAYER_PLANES, LAYER_K2, LAYER_K]:
        asegurar_layer(lyr)

    # 4. Offset del perímetro (0.5 mm) hacia adentro ─────────────────────────
    _log("  [4/8] Offset perimetro {:.1f} mm...".format(OFFSET_PERIM))
    offset_perim_id = offset_hacia_adentro(perim_id, OFFSET_PERIM, LAYER_PLANES)
    if not offset_perim_id:
        _log("  ERROR: No se pudo crear offset del perímetro")
        rs.EnableRedraw(True)
        return

    # 5. Hatch entre perímetro y offset → k2 ─────────────────────────────────
    _log("  [5/8] Hatch k2 (borde perímetro)...")
    hatch_solido([perim_id, offset_perim_id], LAYER_K2)

    # 6. Hatch entre BN y offset → k ─────────────────────────────────────────
    _log("  [6/8] Hatch k (banda negra)...")
    if bn_id:
        hatch_solido([bn_id, offset_perim_id], LAYER_K)
    else:
        _log("  WARN: No se encontró curva de banda negra")

    # 7. Reemplazar logo del plano con LOGO1 del cajetin ────────────────────────
    _log("  [7] Reemplazando logo con LOGO1 del cajetin...")
    _logo1_objs = [o for o in _ids_cajetin
                   if rs.IsObject(o)
                   and rs.ObjectLayer(o).upper().endswith("LOGO1")]
    _log("  LOGO1 encontrados: {}  logo plano: {}".format(len(_logo1_objs), len(_logo_plano_objs)))
    if _logo1_objs and _logo_plano_objs:
        _bbox_plano = rs.BoundingBox(_logo_plano_objs)
        _bbox_logo1 = rs.BoundingBox(_logo1_objs)
        if _bbox_plano and _bbox_logo1:
            cx_pl = (_bbox_plano[0][0] + _bbox_plano[6][0]) / 2.0
            cy_pl = (_bbox_plano[0][1] + _bbox_plano[6][1]) / 2.0
            cx_l1 = (_bbox_logo1[0][0] + _bbox_logo1[6][0]) / 2.0
            cy_l1 = (_bbox_logo1[0][1] + _bbox_logo1[6][1]) / 2.0
            rs.MoveObjects(_logo1_objs, [cx_pl - cx_l1, cy_pl - cy_l1, 0.0])
            for _lo in _logo_plano_objs:
                try:
                    rs.DeleteObject(_lo)
                except Exception:
                    pass
            _log("  Logo reemplazado con LOGO1.")
    elif not _logo1_objs:
        _log("  WARN: no se encontro LOGO1 en el cajetin.")
    elif not _logo_plano_objs:
        _log("  WARN: no hay logo en el plano, LOGO1 se deja donde esta.")

    # 7b. Centrar cajetin 1 en el centro de la pieza ────────────────────────────
    _log("  [7b] Centrando cajetin 1 en la pieza...")
    cajetin_objs = objetos_en_patron("CAJETIN 1")
    bbox_perim = rs.BoundingBox([perim_id])
    if cajetin_objs and bbox_perim:
        bbox_caj = rs.BoundingBox(cajetin_objs)
        if bbox_caj:
            cx_pieza = (bbox_perim[0][0] + bbox_perim[6][0]) / 2.0
            cy_pieza = (bbox_perim[0][1] + bbox_perim[6][1]) / 2.0
            cx_caj   = (bbox_caj[0][0]   + bbox_caj[6][0])   / 2.0
            cy_caj   = (bbox_caj[0][1]   + bbox_caj[6][1])   / 2.0
            rs.MoveObjects(cajetin_objs, [cx_pieza - cx_caj, cy_pieza - cy_caj, 0.0])
            _log("  Cajetin 1 centrado.")
    else:
        _log("  WARN: no se encontro CAJETIN 1 o perimetro.")

    # 8. Degradé ─────────────────────────────────────────────────────────────
    if CON_DEGRADE:
        _log("  [8/8] Degradé: offset BN interior...")
        if bn_id:
            # 8a. Offset de BN hacia adentro
            off_bn_id = offset_hacia_adentro(bn_id, OFFSET_BN_DEG, LAYER_PLANES)
            if off_bn_id:

                # 8b. Longitud del offset 2.5 → calcula pepas
                longitud = rs.CurveLength(off_bn_id) or 0
                n_pepas  = int(round(longitud / DIVISOR_DEG)) if longitud > 0 else 0
                _log("     longitud offset 2.5: {:.2f} mm  pepas: {}".format(longitud, n_pepas))

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
                    _log("     Distribuyendo bloque '{}' x{} sobre offset 2.5...".format(BLOQUE_25, n_pepas))
                    nuevos = distribuir_bloque_freeform(bloque_id, off_bn_id, n_pepas)
                    for nid in nuevos:
                        rs.ObjectLayer(nid, "K3")
                else:
                    _log("  WARN: bloque '{}' no encontrado o n_pepas=0".format(BLOQUE_25))
            else:
                _log("  WARN: No se pudo crear offset interior de BN")

    else:
        _log("  [8/8] Sin degradé — omitido")

    # 9. Borrar basura del import de cajetines (conservar solo CAJETIN 1) ───────
    _log("  [9] Limpiando objetos sobrantes del import...")
    _conservar = set(objetos_en_patron("CAJETIN 1")) | set(_logo1_objs)
    for _oid in _ids_cajetin:
        if _oid not in _conservar:
            try:
                rs.DeleteObject(_oid)
            except Exception:
                pass

    # 10. Mover PERIMETRO y BN/PHANTOM al layer PLANES ───────────────────────
    _log("  [10] Moviendo geometria original a PLANES...")
    for _patron in ["PERIMETRO", "BN", "PHANTOM", "BANDA NEGRA", "BANDANEGRA", "BANDA"]:
        for _obj in objetos_en_patron(_patron):
            try:
                rs.ObjectLayer(_obj, LAYER_PLANES)
            except Exception:
                pass

    # Guardar ────────────────────────────────────────────────────────────────
    if RUTA_SALIDA and RUTA_SALIDA != "":
        _log("  Guardando en: {}".format(RUTA_SALIDA))
        rs.Command('_-Save "{}" _Enter'.format(RUTA_SALIDA), False)

    rs.EnableRedraw(True)
    doc = sc.doc
    if doc:
        doc.Views.Redraw()

    _log("=== Arte Maker: COMPLETADO ===")

    # 11. Ventana para rellenar campos del cajetin ────────────────────────────
    _log("  [11] Abriendo ventana de cajetin...")
    import tkinter as tk
    from tkinter import ttk
    import datetime

    _hoy = datetime.date.today().strftime("%d.%m.%Y")

    # Defaults predeterminados (el usuario puede editarlos antes de aceptar)
    _DEFAULTS = {
        "MEDIDAS":  "Milimetros",
        "VISTA":    "Interna",
        "REVISADO": "Santiago P.",
        "FECHA":    _hoy,
        "ESCALA":   "1:1",
    }

    # Solo los campos que el dibujante debe llenar manualmente.
    # NAGS, VERSION y PIEZA se derivan del Codigo plano.
    # REVISADO, FECHA, VISTA, MEDIDAS y ESCALA se rellenan por defecto.
    _CAMPOS = [
        ("DIBUJO",    "Dibujo"),
        ("VEHICULO",  "Vehiculo"),
        ("MODELO",    "Modelo"),
        ("COD PLANO", "Codigo plano"),
        ("VITRO",     "Vitro"),
        ("MALLA",     "Malla"),
    ]

    def _parsear_cod_plano(cod):
        """
        '1795 003 001-002' -> nags='1795', version='V-003', pieza='001-002'
        Separa por espacios: [0]=NAGS, [1]=VERSION, resto=PIEZA
        """
        partes = cod.strip().split()
        if not partes:
            return "", "", ""
        nags    = partes[0]
        version = ("V-" + partes[1]) if len(partes) > 1 else ""
        pieza   = " ".join(partes[2:]) if len(partes) > 2 else ""
        return nags, version, pieza

    _valores = {}
    _cancelado = [False]

    _ventana = tk.Tk()
    _ventana.title("Rellenar Cajetin 1")
    _ventana.resizable(False, False)
    _ventana.attributes("-topmost", True)

    _frame = ttk.Frame(_ventana, padding=16)
    _frame.grid(row=0, column=0, sticky="nsew")

    _entries = {}
    for _i, (_campo, _etiqueta) in enumerate(_CAMPOS):
        ttk.Label(_frame, text=_etiqueta + ":", anchor="e", width=18).grid(
            row=_i, column=0, sticky="e", pady=4, padx=(0, 8)
        )
        _ent = ttk.Entry(_frame, width=36)
        _ent.grid(row=_i, column=1, sticky="w", pady=4)
        if _campo in _DEFAULTS:
            _ent.insert(0, _DEFAULTS[_campo])
        _entries[_campo] = _ent

    # Al salir del campo "Codigo plano" auto-rellena NAGS, VERSION y PIEZA
    def _on_cod_plano(*_):
        _n, _v, _p = _parsear_cod_plano(_entries["COD PLANO"].get())
        for _k, _val in [("NAGS", _n), ("VERSION", _v), ("PIEZA", _p)]:
            _entries[_k].delete(0, tk.END)
            _entries[_k].insert(0, _val)

    _entries["COD PLANO"].bind("<FocusOut>", _on_cod_plano)
    _entries["COD PLANO"].bind("<Tab>",      _on_cod_plano)

    list(_entries.values())[0].focus_set()

    def _aceptar(*_):
        for _c, _e in _entries.items():
            _valores[_c] = _e.get().strip()
        _ventana.destroy()

    def _cancelar(*_):
        _cancelado[0] = True
        _ventana.destroy()

    _ventana.bind("<Escape>", _cancelar)

    _btn = ttk.Frame(_frame)
    _btn.grid(row=len(_CAMPOS), column=0, columnspan=2, pady=(12, 0))
    ttk.Button(_btn, text="Aceptar",  command=_aceptar).pack(side="left", padx=8)
    ttk.Button(_btn, text="Cancelar", command=_cancelar).pack(side="left", padx=8)

    _ventana.mainloop()

    if not _cancelado[0] and _valores:
        _all_layers = rs.LayerNames() or []
        for _campo, _texto in _valores.items():
            if not _texto:
                continue
            _layer_buscar = "CAJETIN 1${} 1".format(_campo)
            _layer_found = None
            for _ln in _all_layers:
                if _ln.upper().endswith(_layer_buscar.upper()):
                    _layer_found = _ln
                    break
            if _layer_found is None:
                _log("  WARN: layer '{}' no encontrado".format(_layer_buscar))
                continue
            _n = 0
            for _oid in (rs.ObjectsByLayer(_layer_found) or []):
                try:
                    if rs.IsText(_oid):
                        rs.TextObjectText(_oid, _texto)
                        _n += 1
                except Exception:
                    pass
            _log("  {} -> '{}' ({} texto(s))".format(_campo, _texto, _n))
    else:
        _log("  Cajetin: sin cambios.")

    _log("=== Arte Maker: FIN ===")


main()
