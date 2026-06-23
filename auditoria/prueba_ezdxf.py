#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
PRUEBA ezdxf vs AutoCAD COM
===========================
Pasa la ruta de un DWG y compara resultados de ezdxf vs AutoCAD COM.
Uso:
  python auditoria/prueba_ezdxf.py "ruta\al\archivo.dwg"
"""
import sys, re, time, os

# ─── ezdxf ────────────────────────────────────────────────────
try:
    import ezdxf
    _EZDXF_OK = True
except ImportError:
    _EZDXF_OK = False
    print("[!] ezdxf no instalado.  pip install ezdxf")

# ─── AutoCAD COM ──────────────────────────────────────────────
try:
    import win32com.client, pythoncom
    _COM_OK = True
except ImportError:
    _COM_OK = False
    print("[!] pywin32 no instalado.")

COLOR_K  = 5
COLOR_K2 = 3
COLOR_K3 = 1
_RE_MTEXT = re.compile(r'\{[^}]*\}|\\[A-Za-z][^;]*;|%%.')

def _limpiar(s):
    s = _RE_MTEXT.sub(" ", s or "")
    return re.sub(r'\s+', ' ', s).strip()

def _clasificar(textos):
    texto = " ".join(textos)
    vitros  = list(dict.fromkeys(re.findall(r'T-\d{4,6}', texto, re.I)))
    grandes = list(dict.fromkeys(re.findall(r'A-\d{4,6}', texto, re.I)))
    nums_raw = re.findall(r'(?<![TA]-)\b\d{4,6}\b', texto)
    excluir  = {re.sub(r'[TA]-','',v,flags=re.I) for v in vitros+grandes}
    nums = list(dict.fromkeys(n for n in nums_raw if n not in excluir))
    return vitros, grandes, nums

# ═══════════════════════════════════════════════════════════════
#  MÉTODO 1 — ezdxf
# ═══════════════════════════════════════════════════════════════
def analizar_ezdxf(ruta):
    t0 = time.time()
    doc = ezdxf.readfile(ruta)
    msp = doc.modelspace()

    layers = {}
    for lyr in doc.layers:
        aci = lyr.color  # ACI positivo = color fijo; negativo = layer apagado
        layers[lyr.dxf.name.upper().strip()] = abs(aci)

    textos          = []
    layers_con_ents = set()
    hatch_puntos    = False
    trazo_puntos    = False
    _bloques_vistos = set()

    def _texto_ent(e):
        dxftype = e.dxftype()
        lyr = e.dxf.layer.upper().strip() if e.dxf.hasattr("layer") else ""
        if lyr:
            layers_con_ents.add(lyr)

        if dxftype in ("TEXT", "MTEXT", "ATTRIB", "ATTDEF"):
            try:
                t = _limpiar(e.dxf.text if dxftype != "MTEXT" else e.plain_mtext())
                if t: textos.append(t)
            except Exception: pass

        if dxftype == "HATCH" and "PUNTOS" in lyr:
            nonlocal hatch_puntos; hatch_puntos = True
        if dxftype in ("LWPOLYLINE","POLYLINE","LINE","SPLINE") and "PUNTOS" in lyr:
            nonlocal trazo_puntos; trazo_puntos = True

    def _leer_bloque(nombre):
        if nombre in _bloques_vistos: return
        _bloques_vistos.add(nombre)
        try:
            blk = doc.blocks[nombre]
            for e in blk:
                _texto_ent(e)
                if e.dxftype() == "INSERT":
                    try:
                        for attr in e.attribs:
                            _texto_ent(attr)
                    except Exception: pass
                    try: _leer_bloque(e.dxf.name)
                    except Exception: pass
        except Exception: pass

    for e in msp:
        _texto_ent(e)
        if e.dxftype() == "INSERT":
            try:
                for attr in e.attribs:
                    _texto_ent(attr)
            except Exception: pass
            try: _leer_bloque(e.dxf.name)
            except Exception: pass

    vitros, grandes, nums = _clasificar(textos)

    k_color  = layers.get("K")
    k2_color = layers.get("K2")
    k3_color = layers.get("K3")

    logo_ok  = "LOGO" in layers_con_ents or "LOGO1" in layers_con_ents
    logo2_ok = "LOGO2" in layers_con_ents

    puntos_ents = any("PUNTOS" in l for l in layers_con_ents)
    if puntos_ents:
        if hatch_puntos and not trazo_puntos:   puntos = "OK relleno"
        elif trazo_puntos and not hatch_puntos: puntos = "DESACTUALIZADO"
        elif hatch_puntos and trazo_puntos:     puntos = "MIXTO"
        else:                                   puntos = "vacío"
    else:
        puntos = "—"

    elapsed = time.time() - t0
    return {
        "metodo":   "ezdxf",
        "tiempo":   f"{elapsed:.2f}s",
        "K":        f"{'✔' if 'K' in layers_con_ents and k_color==COLOR_K else '✘'}({k_color})",
        "K2":       f"{'✔' if 'K2' in layers_con_ents and k2_color==COLOR_K2 else '✘'}({k2_color})",
        "K3":       f"{'✔' if 'K3' in layers_con_ents and k3_color==COLOR_K3 else '✘'}({k3_color})",
        "Logo":     "✔" if logo_ok else "✘",
        "Logo2":    "✔" if logo2_ok else "—",
        "Puntos":   puntos,
        "vitro":    " / ".join(vitros) or "—",
        "malla_g":  " / ".join(grandes) or "—",
        "malla_p":  " / ".join(nums) or "—",
        "textos_n": len(textos),
    }

# ═══════════════════════════════════════════════════════════════
#  MÉTODO 2 — AutoCAD COM
# ═══════════════════════════════════════════════════════════════
def analizar_com(ruta):
    t0 = time.time()
    pythoncom.CoInitialize()
    acad = win32com.client.GetActiveObject("AutoCAD.Application")
    doc  = acad.Documents.Open(os.path.abspath(ruta), False, True)

    for _ in range(60):
        try:
            if doc.FullName: break
        except Exception: pass
        time.sleep(0.1)

    msp = None
    for _ in range(20):
        try:
            msp = doc.ModelSpace
            _ = msp.Count; break
        except Exception:
            time.sleep(0.3)

    layers = {}
    for i in range(doc.Layers.Count):
        try:
            lyr = doc.Layers.Item(i)
            layers[lyr.Name.upper().strip()] = lyr.Color
        except Exception: pass

    textos = []
    layers_con_ents = set()
    hatch_puntos = trazo_puntos = False
    _RE_MTEXT2 = re.compile(r'\{[^}]*\}|\\[A-Za-z][^;]*;|%%.')
    _bloques_vistos = set()

    def _t(s):
        s = _RE_MTEXT2.sub(" ", s or ""); return re.sub(r'\s+',' ',s).strip()

    def _ent(e):
        n = e.ObjectName
        lyr = e.Layer.upper().strip()
        layers_con_ents.add(lyr)
        if "Text" in n or "Attrib" in n:
            try:
                t = _t(e.TextString)
                if t: textos.append(t)
            except Exception: pass
        if "PUNTOS" in lyr:
            if n == "AcDbHatch": nonlocal hatch_puntos; hatch_puntos = True
            elif n in ("AcDbPolyline","AcDb2dPolyline","AcDbLine","AcDbSpline"):
                nonlocal trazo_puntos; trazo_puntos = True

    def _blk(nombre):
        if nombre in _bloques_vistos: return
        _bloques_vistos.add(nombre)
        try:
            for be in doc.Blocks.Item(nombre):
                try:
                    _ent(be)
                    if be.ObjectName == "AcDbBlockReference":
                        try:
                            for a in be.GetAttributes(): _ent(a)
                        except Exception: pass
                        try: _blk(be.Name)
                        except Exception: pass
                except Exception: pass
        except Exception: pass

    for e in msp:
        try:
            _ent(e)
            if e.ObjectName == "AcDbBlockReference":
                try:
                    for a in e.GetAttributes(): _ent(a)
                except Exception: pass
                try: _blk(e.Name)
                except Exception: pass
        except Exception: pass

    doc.Close(False)
    pythoncom.CoUninitialize()

    vitros, grandes, nums = _clasificar(textos)
    k_color  = layers.get("K")
    k2_color = layers.get("K2")
    k3_color = layers.get("K3")
    logo_ok  = "LOGO" in layers_con_ents or "LOGO1" in layers_con_ents
    logo2_ok = "LOGO2" in layers_con_ents
    puntos_ents = any("PUNTOS" in l for l in layers_con_ents)
    if puntos_ents:
        if hatch_puntos and not trazo_puntos:   puntos = "OK relleno"
        elif trazo_puntos and not hatch_puntos: puntos = "DESACTUALIZADO"
        elif hatch_puntos and trazo_puntos:     puntos = "MIXTO"
        else:                                   puntos = "vacío"
    else:
        puntos = "—"

    elapsed = time.time() - t0
    return {
        "metodo":   "AutoCAD COM",
        "tiempo":   f"{elapsed:.2f}s",
        "K":        f"{'✔' if 'K' in layers_con_ents and k_color==COLOR_K else '✘'}({k_color})",
        "K2":       f"{'✔' if 'K2' in layers_con_ents and k2_color==COLOR_K2 else '✘'}({k2_color})",
        "K3":       f"{'✔' if 'K3' in layers_con_ents and k3_color==COLOR_K3 else '✘'}({k3_color})",
        "Logo":     "✔" if logo_ok else "✘",
        "Logo2":    "✔" if logo2_ok else "—",
        "Puntos":   puntos,
        "vitro":    " / ".join(vitros) or "—",
        "malla_g":  " / ".join(grandes) or "—",
        "malla_p":  " / ".join(nums) or "—",
        "textos_n": len(textos),
    }

# ═══════════════════════════════════════════════════════════════
#  MAIN
# ═══════════════════════════════════════════════════════════════
if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Uso: python prueba_ezdxf.py \"ruta\\archivo.dwg\"")
        sys.exit(1)

    ruta = sys.argv[1]
    if not os.path.exists(ruta):
        print(f"Archivo no encontrado: {ruta}"); sys.exit(1)

    print(f"\nArchivo: {os.path.basename(ruta)}")
    print("=" * 60)

    resultados = []

    if _EZDXF_OK:
        print("\n[1/2] Analizando con ezdxf...")
        try:
            r = analizar_ezdxf(ruta)
            resultados.append(r)
        except Exception as e:
            print(f"  ERROR ezdxf: {e}")
    else:
        print("ezdxf no disponible, instala con:  pip install ezdxf")

    if _COM_OK:
        print("\n[2/2] Analizando con AutoCAD COM...")
        try:
            r = analizar_com(ruta)
            resultados.append(r)
        except Exception as e:
            print(f"  ERROR COM: {e}")

    print("\n" + "=" * 60)
    print(f"{'Campo':<12}", end="")
    for r in resultados:
        print(f"  {r['metodo']:<18}", end="")
    print()
    print("-" * 60)
    for campo in ["tiempo","K","K2","K3","Logo","Logo2","Puntos","vitro","malla_g","malla_p","textos_n"]:
        print(f"{campo:<12}", end="")
        for r in resultados:
            print(f"  {str(r.get(campo,'')):<18}", end="")
        print()
    print("=" * 60)
