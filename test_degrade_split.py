# -*- coding: utf-8 -*-
"""
TEST: Degradé con bloques mixtos (25 y E25) por segmento.
Ejecutar desde terminal con AutoCAD abierto y el plano activo.
"""
import os, sys, time, math, datetime

try:
    import win32com.client, pythoncom
except ImportError:
    print("ERROR: pip install pywin32"); sys.exit(1)

# ── Config ────────────────────────────────────────────────────────────────────
_DIR         = os.path.dirname(os.path.abspath(__file__))
CAJETIN_DWG  = os.path.join(_DIR, "LAYERS Y CAJETINES 1.dwg")
BLOQUE_25    = "25"
BLOQUE_E25   = "E25"
OFFSET_BN    = 2.5      # mm — off_bn = 2.5mm dentro de bn0
RADIO_25       = 6.5    # mm — gap threshold: si gap > RADIO_25 → usar E25
PEQUENO_EDGE   = 1.76   # mm — cy - r del círculo PEQUEÑO (2.01 - 0.25, medido del bloque)
DIVISOR_DEG  = 3
MIN_SEG_MM   = 100.0
LAYER_K3     = "k3"
LAYER_PLANES = "PLANES"

PAT_BN = ["BANDA NEGRA", "BANDANEGRA", "BN", "PHANTOM", "BANDA"]

# ── Helpers ───────────────────────────────────────────────────────────────────
def ts():
    return datetime.datetime.now().strftime("%H:%M:%S.%f")[:-3]

def log(msg):
    print(f"[{ts()}] {msg}")

def pt(x, y, z=0.0):
    return win32com.client.VARIANT(
        pythoncom.VT_ARRAY | pythoncom.VT_R8, [float(x), float(y), float(z)])

def area_bbox(ent):
    try:
        lo, hi = ent.GetBoundingBox()
        return abs(hi[0]-lo[0]) * abs(hi[1]-lo[1])
    except:
        return 0.0

def _layer_match(nombre, patron):
    n = nombre.upper().strip()
    if "OEM" in n: return False
    p = patron.upper().strip()
    if p.startswith("="): return n == p[1:]
    return p in n

def ents_por_patron(msp, patrones):
    res = []
    for e in msp:
        try:
            if any(_layer_match(e.Layer, p) for p in patrones):
                res.append(e)
        except: pass
    return res

def es_poly_cerrada(e):
    try: return "Polyline" in e.ObjectName and e.Closed
    except: return False

def asegurar_layer(doc, nombre):
    try:
        doc.Layers.Item(nombre)
    except Exception:
        doc.Layers.Add(nombre)

def offset_outward(ent, dist):
    """Offset hacia afuera — elige el resultado con área MAYOR al original."""
    area_orig = area_bbox(ent)
    todos = []
    for d in [dist, -dist]:
        try: todos.extend(list(ent.Offset(d)))
        except: pass
    mayores = [r for r in todos if area_bbox(r) > area_orig]
    mejor = min(mayores, key=area_bbox) if mayores else (max(todos, key=area_bbox) if todos else None)
    for r in todos:
        if r != mejor:
            try: r.Delete()
            except: pass
    return mejor

def offset_inward(ent, dist):
    area_orig = area_bbox(ent)
    todos = []
    for d in [dist, -dist]:
        try: todos.extend(list(ent.Offset(d)))
        except: pass
    menores = [r for r in todos if area_bbox(r) < area_orig]
    mejor = min(menores, key=area_bbox) if menores else (min(todos, key=area_bbox) if todos else None)
    for r in todos:
        if r != mejor:
            try: r.Delete()
            except: pass
    return mejor


# ── Vértices y bulges ─────────────────────────────────────────────────────────

def _get_vertices(ent):
    coords = list(ent.Coordinates)
    es_3d  = "3d" in ent.ObjectName.lower()
    paso   = 3 if es_3d else 2
    n      = len(coords) // paso
    return [(coords[i*paso], coords[i*paso+1]) for i in range(n)]


def _get_bulges(ent, n_verts):
    bulges = []
    for i in range(n_verts):
        try:
            bulges.append(ent.GetBulge(i))
        except Exception:
            bulges.append(0.0)
    return bulges


# ── Medir gaps ────────────────────────────────────────────────────────────────

def medir_gaps(off_bn, bn1):
    resultados = []
    try:
        verts_off = _get_vertices(off_bn)
        verts_bn1 = _get_vertices(bn1)
        puntos_bn1 = list(verts_bn1)
        for i in range(len(verts_bn1)):
            j = (i + 1) % len(verts_bn1)
            puntos_bn1.append(((verts_bn1[i][0]+verts_bn1[j][0])/2,
                                (verts_bn1[i][1]+verts_bn1[j][1])/2))
        log(f"  off_bn: {len(verts_off)} vértices  bn1: {len(verts_bn1)} vértices")
        for idx, (x, y) in enumerate(verts_off):
            gap = min(math.hypot(bx-x, by-y) for bx, by in puntos_bn1)
            resultados.append((idx, x, y, gap))
    except Exception as e:
        log(f"  ERROR medir_gaps: {e}")
    return resultados


# ── Segmentar ─────────────────────────────────────────────────────────────────

def segmentar(gaps_data):
    if not gaps_data:
        return []

    etiquetas = [BLOQUE_E25 if g > RADIO_25 else BLOQUE_25 for _, _, _, g in gaps_data]

    segmentos = []
    bloque_actual  = etiquetas[0]
    indices_actual = [gaps_data[0][0]]

    for i in range(1, len(etiquetas)):
        idx = gaps_data[i][0]
        if etiquetas[i] == bloque_actual:
            indices_actual.append(idx)
        else:
            segmentos.append({'bloque': bloque_actual, 'indices': indices_actual})
            bloque_actual  = etiquetas[i]
            indices_actual = [gaps_data[i-1][0], idx]
    segmentos.append({'bloque': bloque_actual, 'indices': indices_actual})

    coord_map = {d[0]: (d[1], d[2]) for d in gaps_data}

    def longitud_seg(indices):
        total = 0.0
        for j in range(1, len(indices)):
            x0, y0 = coord_map[indices[j-1]]
            x1, y1 = coord_map[indices[j]]
            total += math.hypot(x1-x0, y1-y0)
        return total

    fusionados = True
    while fusionados and len(segmentos) > 1:
        fusionados = False
        nuevos = []
        i = 0
        while i < len(segmentos):
            seg = segmentos[i]
            if longitud_seg(seg['indices']) < MIN_SEG_MM:
                if nuevos:
                    nuevos[-1]['indices'] += seg['indices'][1:]
                    if seg['bloque'] == BLOQUE_E25:
                        nuevos[-1]['bloque'] = BLOQUE_E25
                    fusionados = True
                elif i + 1 < len(segmentos):
                    segmentos[i+1]['indices'] = seg['indices'] + segmentos[i+1]['indices'][1:]
                    if seg['bloque'] == BLOQUE_E25:
                        segmentos[i+1]['bloque'] = BLOQUE_E25
                    fusionados = True
                else:
                    nuevos.append(seg)
            else:
                nuevos.append(seg)
            i += 1
        segmentos = nuevos

    return segmentos


# ── Crear polilínea con bulges ─────────────────────────────────────────────────

def crear_polilinea(msp, off_bn, indices):
    verts_orig  = _get_vertices(off_bn)
    bulges_orig = _get_bulges(off_bn, len(verts_orig))
    coords_flat = []
    for idx in indices:
        x, y = verts_orig[idx]
        coords_flat.extend([float(x), float(y)])
    coords_var = win32com.client.VARIANT(
        pythoncom.VT_ARRAY | pythoncom.VT_R8, coords_flat)
    pl = msp.AddLightWeightPolyline(coords_var)
    pl.Closed = False
    for i, orig_idx in enumerate(indices):
        try:
            pl.SetBulge(i, bulges_orig[orig_idx])
        except Exception:
            pass
    try:
        pl.Update()
    except Exception:
        pass
    return pl


# ── DIVIDE en un segmento ─────────────────────────────────────────────────────

def divide_segmento(doc, handle, bloque, longitud):
    n_pepas = max(1, int(round(longitud / DIVISOR_DEG)))
    log(f"    DIVIDE handle={handle} bloque={bloque} pepas={n_pepas} long={longitud:.1f}mm")
    doc.SendCommand(f'CLAYER\n{LAYER_K3}\n')
    time.sleep(0.5)
    doc.SendCommand(
        f'DIVIDE\n(handent "{handle}")\nB\n{bloque}\nY\n{n_pepas}\n'
    )
    espera = max(5.0, n_pepas * 0.05)   # más margen para segmentos largos
    log(f"    Esperando {espera:.1f}s...")
    time.sleep(espera)
    # Retry CLAYER con backoff por si AutoCAD sigue ocupado
    for intento in range(3):
        try:
            doc.SendCommand("CLAYER\n0\n")
            time.sleep(0.5)
            break
        except Exception:
            log(f"    AutoCAD ocupado, reintento {intento+1}...")
            time.sleep(3.0)


def asegurar_bloque(doc, bloque):
    try:
        doc.Blocks.Item(bloque)
        log(f"  Bloque '{bloque}' OK.")
    except Exception:
        log(f"  Importando bloque '{bloque}'...")
        abs_caj = os.path.abspath(CAJETIN_DWG)
        doc.SendCommand(f'-INSERT "{abs_caj}"\n0,0,0\n1\n1\n0\n')
        time.sleep(4)
        doc.SendCommand("ERASE\nL\n\n")
        time.sleep(1.5)
        log(f"  Bloque '{bloque}' registrado.")


# ── MAIN ──────────────────────────────────────────────────────────────────────

def main():
    pythoncom.CoInitialize()
    try:
        log("=== TEST degradé split ===")
        acad = win32com.client.GetActiveObject("AutoCAD.Application")
        doc  = acad.ActiveDocument
        msp  = doc.ModelSpace
        log(f"Documento activo: {doc.Name}")

        # 1. BN
        bn_todos = sorted(
            [e for e in ents_por_patron(msp, PAT_BN) if es_poly_cerrada(e)],
            key=area_bbox, reverse=True)
        log(f"\n[1] BN cerrados: {len(bn_todos)}")
        for i, b in enumerate(bn_todos):
            log(f"   [{i}] area={area_bbox(b):.0f}  layer={b.Layer}")

        area_max = area_bbox(bn_todos[0])
        bn_grandes = [b for b in bn_todos if area_bbox(b) > area_max * 0.1]
        if len(bn_grandes) < 2:
            log("ERROR: necesito ≥2 BN grandes.")
            return

        bn0 = bn_grandes[0]
        bn1 = bn_grandes[-1]
        log(f"   bn0 (outer) area={area_bbox(bn0):.0f}  bn1 (inner) area={area_bbox(bn1):.0f}")

        # 2. off_bn (bloque 25) y off_bn_e25 (bloque E25, path en bn0)
        log(f"\n[2] Calculando paths...")
        asegurar_layer(doc, LAYER_PLANES)

        off_bn = offset_inward(bn0, OFFSET_BN)       # 2.5mm — path bloque 25
        if not off_bn:
            log("ERROR: no se pudo crear off_bn.")
            return
        off_bn.Layer = LAYER_PLANES
        log(f"   off_bn (25): {off_bn.Length:.1f}mm")

        # E25 path = offset_outward(bn1, PEQUENO_EDGE)
        # → edge exterior de PEQUEÑO queda exactamente en bn_last (segunda phantom)
        # → se adapta automáticamente al ancho de banda en cada sección
        off_bn_e25 = offset_outward(bn1, PEQUENO_EDGE)
        if off_bn_e25:
            off_bn_e25.Layer = LAYER_PLANES
            verts_e25_g = _get_vertices(off_bn_e25)
            log(f"   off_bn_e25 (E25, {PEQUENO_EDGE}mm outward desde bn1): {off_bn_e25.Length:.1f}mm  verts={len(verts_e25_g)}")
        else:
            verts_e25_g = None
            log("   WARN: off_bn_e25 falló → E25 usará off_bn")

        # 3. Medir gaps
        log(f"\n[3] Gaps off_bn → bn1...")
        gaps_data = medir_gaps(off_bn, bn1)
        if not gaps_data:
            log("ERROR: sin gaps.")
            return
        gaps_vals = [g for _, _, _, g in gaps_data]
        log(f"   mín={min(gaps_vals):.2f}  máx={max(gaps_vals):.2f}  avg={sum(gaps_vals)/len(gaps_vals):.2f}mm")
        log(f"   → 25: {sum(1 for g in gaps_vals if g<=RADIO_25)}  E25: {sum(1 for g in gaps_vals if g>RADIO_25)} vértices")

        # 4. Segmentar — UNA SOLA VEZ con threshold fijo RADIO_25
        log(f"\n[4] Segmentando (threshold={RADIO_25}mm, mín={MIN_SEG_MM}mm)...")
        segmentos = segmentar(gaps_data)

        # Fix curva de cierre del off_bn cerrado
        if segmentos and gaps_data[0][0] not in segmentos[-1]['indices']:
            segmentos[-1]['indices'].append(gaps_data[0][0])
            log("   [fix] curva de cierre añadida al último segmento")

        verts_off = _get_vertices(off_bn)
        log(f"   {len(segmentos)} segmentos:")
        for i, seg in enumerate(segmentos):
            idxs = seg['indices']
            long_seg = sum(
                math.hypot(verts_off[idxs[j]][0]-verts_off[idxs[j-1]][0],
                           verts_off[idxs[j]][1]-verts_off[idxs[j-1]][1])
                for j in range(1, len(idxs)))
            log(f"   Seg {i+1}: {seg['bloque']}  {len(idxs)} verts  ≈{long_seg:.0f}mm")

        # 5. Confirmar
        print("\n" + "="*60)
        print(f"Se crearán {len(segmentos)} segmentos. ¿Continuar? (s/n): ", end="")
        if input().strip().lower() != 's':
            try: off_bn.Delete()
            except: pass
            return

        # 6. Bloques
        log("\n[5] Verificando bloques...")
        for blk in list({s['bloque'] for s in segmentos}):
            asegurar_bloque(doc, blk)

        # 7. DIVIDE
        # Bloque 25: path = off_bn (confirmado correcto)
        # Bloque E25: path = off_bn_e25 (~bn0) → PEQUEÑO queda tocando bn0
        # Para E25 se mapean los índices de off_bn a off_bn_e25 por proximidad.
        log("\n[6] DIVIDE por segmentos...")
        import ctypes
        verts_off = _get_vertices(off_bn)

        def indices_para_e25(idxs_off_bn):
            """Mapea índices de off_bn a los vértices más cercanos en off_bn_e25."""
            if not verts_e25_g:
                return idxs_off_bn
            resultado = []
            for idx in idxs_off_bn:
                vx, vy = verts_off[idx]
                nearest = min(range(len(verts_e25_g)),
                              key=lambda i: math.hypot(verts_e25_g[i][0]-vx, verts_e25_g[i][1]-vy))
                if not resultado or resultado[-1] != nearest:
                    resultado.append(nearest)
            return resultado if len(resultado) >= 2 else idxs_off_bn

        pls_temp = []
        for i, seg in enumerate(segmentos):
            if len(seg['indices']) < 2:
                continue
            if seg['bloque'] == BLOQUE_E25 and verts_e25_g:
                idxs_e25 = indices_para_e25(seg['indices'])
                pl = crear_polilinea(msp, off_bn_e25, idxs_e25)
                log(f"   Seg {i+1}: E25 path=~bn0  verts={len(idxs_e25)}")
            else:
                pl = crear_polilinea(msp, off_bn, seg['indices'])
            pl.Layer = LAYER_PLANES
            pls_temp.append(pl)
            log(f"   Seg {i+1}/{len(segmentos)}: {seg['bloque']} long={pl.Length:.0f}mm")
            divide_segmento(doc, pl.Handle, seg['bloque'], float(pl.Length))

        # Orientación — pregunta al final para ambos bloques
        resp_ori = ctypes.windll.user32.MessageBoxW(
            0,
            "¿El degradé quedó correcto?\n"
            "(PEQUEÑO debe tocar la línea exterior = bn0)\n\n"
            "Sí = OK\nNo = invertir 180°",
            "TEST degradé — orientación", 0x24)
        if resp_ori == 7:
            _msp2 = doc.ModelSpace
            rotados = 0
            for _e in list(_msp2):
                try:
                    if (_e.Layer.upper() == LAYER_K3.upper()
                            and _e.ObjectName == "AcDbBlockReference"):
                        lo, hi = _e.GetBoundingBox()
                        _e.Rotate(pt((lo[0]+hi[0])/2, (lo[1]+hi[1])/2, 0.0), math.pi)
                        rotados += 1
                except Exception:
                    pass
            log(f"  {rotados} bloques invertidos ✔")

        # 8. Cleanup
        log(f"\n[7] Limpiando {len(pls_temp)} polilíneas + off_bn...")
        for pl in pls_temp:
            try: pl.Delete()
            except: pass
        try: off_bn.Delete()
        except: pass
        if off_bn_e25:
            try: off_bn_e25.Delete()
            except: pass

        msp2 = doc.ModelSpace
        k3 = sum(1 for e in msp2
                 if e.Layer.upper() == LAYER_K3.upper()
                 and e.ObjectName == "AcDbBlockReference")
        log(f"\n[8] Bloques k3: {k3}")
        log("=== TEST completado ===")

    except Exception as e:
        log(f"ERROR FATAL: {e}")
        import traceback; traceback.print_exc()
    finally:
        pythoncom.CoUninitialize()


if __name__ == "__main__":
    main()
    input("\nPresiona Enter para cerrar...")
