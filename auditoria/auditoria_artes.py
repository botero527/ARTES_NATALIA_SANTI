#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
AUDITORÍA ARTES AGP — validación + BD update
=============================================
Recorre  : \\\\192.168.2.37\\ingenieria\\PRODUCCION\\AGP PLANOS TECNICOS\\{marca}\\...\\ARTES\\BN\\*.dwg
Valida   :
  1. Layers K(azul=5), K2(verde=3), K3(rojo=1)
  2. Layer Logo1 (y Logo2 si tiene 2 logos)
  3. Layer PUNTOS (parabrisas con degradé especial)
  4. Parabrisas detectado por pieza terminada en 000 / 00
  5. Lee todos los textos → detecta vitro (T-XXXXX) y malla (A-XXXXX / numérico)
  6. Valida vitro/malla contra Azure SQL
  7. Actualiza ruta en BD si se encontró
Genera   : Excel con hoja por marca + resumen global
Checkpoint: JSON guardado cada CHECKPOINT_CADA archivos (retoma si AutoCAD se cae)

Uso:
  python auditoria_artes.py                   # audita todas las marcas
  python auditoria_artes.py --marca CHEVROLET # solo una marca
  python auditoria_artes.py --reiniciar       # borra checkpoint y empieza de cero
  python auditoria_artes.py --solo-excel      # regenera Excel desde checkpoint sin abrir AutoCAD
"""

import os, sys, re, json, time, argparse, traceback, threading
from datetime import datetime

# ─── rutas del proyecto (para importar pymssql / config) ───────────────────
_DIR_SCRIPT = os.path.dirname(os.path.abspath(__file__))
_DIR_RAIZ   = os.path.dirname(_DIR_SCRIPT)
sys.path.insert(0, _DIR_RAIZ)

# ═══════════════════════════════════════════════════════════════
#  CONFIGURACIÓN
# ═══════════════════════════════════════════════════════════════
RUTA_BASE          = r"\\192.168.2.37\ingenieria\PRODUCCION\AGP PLANOS TECNICOS"
CHECKPOINT_FILE    = os.path.join(_DIR_SCRIPT, "auditoria_checkpoint.json")
EXCEL_SALIDA       = os.path.join(_DIR_SCRIPT, f"Auditoria_Artes_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx")
CHECKPOINT_CADA    = 10      # guardar checkpoint cada N archivos
TIMEOUT_ARCHIVO    = 30      # segundos máximo por DWG
EXCLUIR_CARPETAS   = {"OBSOLETO", "OBSOLETOS", "ANTIGUO", "OLD"}

# Colores de layer esperados (ACI)
COLOR_K   = 5   # azul
COLOR_K2  = 3   # verde
COLOR_K3  = 1   # rojo

# BD Azure
BD_SERVER   = "agpcolombia.database.windows.net"
BD_PORT     = 1433
BD_USER     = "DevIngenieria"
BD_PASSWORD = "HiJE068i0LQVrwA"
BD_DATABASE = "AGP_Ingenieria"

# ═══════════════════════════════════════════════════════════════
#  DEPENDENCIAS
# ═══════════════════════════════════════════════════════════════
try:
    import win32com.client, pythoncom
except ImportError:
    print("Falta pywin32.  pip install pywin32"); sys.exit(1)

try:
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
except ImportError:
    print("Falta openpyxl.  pip install openpyxl"); sys.exit(1)

try:
    import pymssql as _pymssql
except ImportError:
    print("Falta pymssql.  pip install pymssql"); sys.exit(1)

try:
    import ezdxf as _ezdxf
    _EZDXF_OK = True
except ImportError:
    _ezdxf = None
    _EZDXF_OK = False


# ═══════════════════════════════════════════════════════════════
#  LOGGER
# ═══════════════════════════════════════════════════════════════
def _ts(): return time.strftime("%H:%M:%S")
def log_info(m):  print(f"{_ts()}  {m}")
def log_warn(m):  print(f"{_ts()}  [!] {m}")
def log_err(m):   print(f"{_ts()}  [X] {m}")
def log_ok(m):    print(f"{_ts()}  [✓] {m}")


# ═══════════════════════════════════════════════════════════════
#  BD — conexión pymssql con wrapper ?→%s
# ═══════════════════════════════════════════════════════════════
class _CursorWrap:
    def __init__(self, c): self._c = c
    def execute(self, sql, p=()):  return self._c.execute(sql.replace("?","%s"), p or ())
    def fetchone(self):            return self._c.fetchone()
    def fetchall(self):            return self._c.fetchall()
    def __getattr__(self, n):      return getattr(self._c, n)

class _ConnWrap:
    def __init__(self, c): self._c = c
    def cursor(self):   return _CursorWrap(self._c.cursor())
    def commit(self):   self._c.commit()
    def rollback(self): self._c.rollback()
    def close(self):    self._c.close()

_BD_CONN = None

def bd_conectar():
    global _BD_CONN
    try:
        if _BD_CONN:
            _BD_CONN.cursor().execute("SELECT 1")
            return _BD_CONN
    except Exception:
        _BD_CONN = None
    _BD_CONN = _ConnWrap(_pymssql.connect(
        server=BD_SERVER, port=BD_PORT, user=BD_USER,
        password=BD_PASSWORD, database=BD_DATABASE,
        timeout=20, login_timeout=20, charset="UTF-8", tds_version="7.3",
    ))
    return _BD_CONN

def bd_validar_y_actualizar(textos, ruta_dwg):
    """
    Busca vitro/malla en los textos extraídos del DWG,
    valida contra BD y actualiza la ruta.
    Retorna (vitro, malla, bd_actualizado, notas_bd, duplicados)
    duplicados: lista de dicts {codigo, tipo, ruta_anterior, ruta_nueva}
    """
    texto_unido = " ".join(textos)
    vitros  = list(dict.fromkeys(re.findall(r'T-\d{4,6}', texto_unido, re.I)))
    grandes = list(dict.fromkeys(re.findall(r'A-\d{4,6}', texto_unido, re.I)))
    # Excluir números que son parte de T-XXXXX o A-XXXXX
    nums_raw = re.findall(r'(?<![TA]-)\b\d{4,6}\b', texto_unido)
    # También excluir los que coincidan con los dígitos de vitros/grandes ya encontrados
    vitro_nums = {re.sub(r'[TA]-', '', v, flags=re.I) for v in vitros + grandes}
    nums = list(dict.fromkeys(n for n in nums_raw if n not in vitro_nums))

    vitro_ok, malla_ok, actualizados, notas, duplicados = [], [], 0, [], []

    try:
        cn  = bd_conectar()
        cur = cn.cursor()

        # Vitros
        for v in vitros:
            cur.execute("SELECT ruta FROM mallas.vitrojet WHERE vitro=?", (v.upper(),))
            row = cur.fetchone()
            if row is not None:
                vitro_ok.append(v.upper())
                ruta_anterior = row[0] if row[0] else ""
                if ruta_anterior and ruta_anterior.lower() != ruta_dwg.lower():
                    duplicados.append({
                        "codigo": v.upper(), "tipo": "vitro",
                        "ruta_anterior": ruta_anterior, "ruta_nueva": ruta_dwg,
                    })
                cur.execute("UPDATE mallas.vitrojet SET ruta=? WHERE vitro=?",
                            (ruta_dwg, v.upper()))
                actualizados += 1

        # Mallas grandes
        for m in grandes:
            cur.execute("SELECT ruta_dwg FROM mallas.grandes WHERE codigo=?", (m.upper(),))
            row = cur.fetchone()
            if row is not None:
                malla_ok.append(m.upper())
                ruta_anterior = row[0] if row[0] else ""
                if ruta_anterior and ruta_anterior.lower() != ruta_dwg.lower():
                    duplicados.append({
                        "codigo": m.upper(), "tipo": "malla_grande",
                        "ruta_anterior": ruta_anterior, "ruta_nueva": ruta_dwg,
                    })
                cur.execute("UPDATE mallas.grandes SET ruta_dwg=? WHERE codigo=?",
                            (ruta_dwg, m.upper()))
                actualizados += 1

        # Mallas pequeñas (números)
        for n in nums:
            cur.execute("SELECT ruta_dwg FROM mallas.pequenas WHERE CAST(codigo AS NVARCHAR)=?", (n,))
            row = cur.fetchone()
            if row is not None:
                malla_ok.append(n)
                ruta_anterior = row[0] if row[0] else ""
                if ruta_anterior and ruta_anterior.lower() != ruta_dwg.lower():
                    duplicados.append({
                        "codigo": n, "tipo": "malla_pequena",
                        "ruta_anterior": ruta_anterior, "ruta_nueva": ruta_dwg,
                    })
                cur.execute("UPDATE mallas.pequenas SET ruta_dwg=? WHERE CAST(codigo AS NVARCHAR)=?",
                            (ruta_dwg, n))
                actualizados += 1

        if actualizados:
            cn.commit()

    except Exception as e:
        notas.append(f"BD error: {e}")

    return (
        " / ".join(vitro_ok) or "—",
        " / ".join(malla_ok) or "—",
        actualizados > 0,
        "; ".join(notas),
        duplicados,
    )


# ═══════════════════════════════════════════════════════════════
#  DETECCIÓN PARABRISAS
# ═══════════════════════════════════════════════════════════════
def es_parabrisas(nombre_archivo):
    """
    Detecta si el arte es de parabrisas o posterior.
    La pieza termina en 000 o 00 (con posibles letras al final).
    Ej: P A-9959 000.dwg → True
        P 1507 000 010 B.dwg → el último grupo es 010, NO parabrisas
        P 1507 000.dwg → True
    """
    stem = os.path.splitext(nombre_archivo)[0]
    # Último grupo de dígitos consecutivos en el nombre
    grupos = re.findall(r'\d+', stem)
    if not grupos:
        return False
    ultimo = grupos[-1]
    # Es parabrisas si termina en 00 o 000 (sin dígitos adicionales después)
    return bool(re.fullmatch(r'0{2,3}', ultimo))


# ═══════════════════════════════════════════════════════════════
#  ANÁLISIS CON EZDXF (fallback sin AutoCAD)
# ═══════════════════════════════════════════════════════════════
_RE_MTEXT_EZ = re.compile(r'\{[^}]*\}|\\[A-Za-z][^;]*;|%%.')

def _limpiar_ez(s):
    s = _RE_MTEXT_EZ.sub(" ", s or "")
    return re.sub(r'\s+', ' ', s).strip()

def analizar_ezdxf(ruta_dwg):
    """Lee el DWG con ezdxf, sin necesidad de AutoCAD abierto."""
    para = es_parabrisas(os.path.basename(ruta_dwg))
    resultado = {
        "ruta": ruta_dwg, "archivo": os.path.basename(ruta_dwg),
        "es_parabrisas": para,
        "k_ok": False, "k_color": None, "k2_ok": False, "k2_color": None,
        "k3_ok": False, "k3_color": None, "logo1_ok": False, "logo2_ok": False,
        "puntos_ok": False, "puntos_estado": "—",
        "vitro": "—", "malla": "—", "bd_actualizado": False, "duplicados": [],
        "estado": "OK", "notas": [], "error": None, "metodo": "ezdxf",
    }
    try:
        doc = _ezdxf.readfile(ruta_dwg)
        msp = doc.modelspace()

        # Layers y colores
        layers = {}
        for lyr in doc.layers:
            aci = lyr.color
            layers[lyr.dxf.name.upper().strip()] = abs(aci)

        textos           = []
        layers_con_ents  = set()
        hatch_puntos     = False
        trazo_puntos     = False
        _bloques_vistos  = set()

        def _texto_ez(e):
            t = e.dxftype()
            lyr = e.dxf.layer.upper().strip() if e.dxf.hasattr("layer") else ""
            if lyr:
                layers_con_ents.add(lyr)
            if t in ("TEXT", "ATTRIB", "ATTDEF"):
                try:
                    s = _limpiar_ez(e.dxf.text)
                    if s: textos.append(s)
                except Exception: pass
            elif t == "MTEXT":
                try:
                    s = _limpiar_ez(e.plain_mtext())
                    if s: textos.append(s)
                except Exception: pass
            if "PUNTOS" in lyr:
                nonlocal hatch_puntos, trazo_puntos
                if t == "HATCH": hatch_puntos = True
                elif t in ("LWPOLYLINE","POLYLINE","LINE","SPLINE"): trazo_puntos = True

        def _leer_blk(nombre):
            if nombre in _bloques_vistos: return
            _bloques_vistos.add(nombre)
            try:
                for e in doc.blocks[nombre]:
                    _texto_ez(e)
                    if e.dxftype() == "INSERT":
                        try:
                            for a in e.attribs: _texto_ez(a)
                        except Exception: pass
                        try: _leer_blk(e.dxf.name)
                        except Exception: pass
            except Exception: pass

        for e in msp:
            _texto_ez(e)
            if e.dxftype() == "INSERT":
                try:
                    for a in e.attribs: _texto_ez(a)
                except Exception: pass
                try: _leer_blk(e.dxf.name)
                except Exception: pass

        # Validar K/K2/K3
        for key, color_esp, campo_ok, campo_color in [
            ("K", COLOR_K, "k_ok", "k_color"),
            ("K2", COLOR_K2, "k2_ok", "k2_color"),
            ("K3", COLOR_K3, "k3_ok", "k3_color"),
        ]:
            color_real = layers.get(key)
            resultado[campo_color] = color_real
            if key in layers_con_ents and color_real == color_esp:
                resultado[campo_ok] = True
            elif key in layers_con_ents and color_real != color_esp:
                resultado["notas"].append(f"Layer {key} color incorrecto ({color_real}≠{color_esp})")
            elif key not in layers_con_ents:
                resultado["notas"].append(f"Falta layer {key}")

        # Logo
        tiene_logo = "LOGO" in layers_con_ents or "LOGO1" in layers_con_ents
        resultado["logo1_ok"] = tiene_logo
        resultado["logo2_ok"] = "LOGO2" in layers_con_ents
        if not tiene_logo:
            resultado["notas"].append("Falta layer LOGO / LOGO1")

        # Puntos (solo parabrisas)
        if para:
            tiene_puntos = any("PUNTOS" in l for l in layers_con_ents)
            resultado["puntos_ok"] = tiene_puntos
            if not tiene_puntos:
                resultado["puntos_estado"] = "FALTA"
                resultado["notas"].append("Parabrisas sin layer PUNTOS")
            elif hatch_puntos and not trazo_puntos:
                resultado["puntos_estado"] = "OK relleno"
            elif trazo_puntos and not hatch_puntos:
                resultado["puntos_estado"] = "DESACTUALIZADO (trazo suelto)"
                resultado["notas"].append("PUNTOS: trazo suelto sin relleno")
            else:
                resultado["puntos_estado"] = "MIXTO"

        # BD
        vitro, malla, bd_ok, notas_bd, dups = bd_validar_y_actualizar(textos, ruta_dwg)
        resultado["vitro"] = vitro
        resultado["malla"] = malla
        resultado["bd_actualizado"] = bd_ok
        resultado["duplicados"] = dups
        if notas_bd: resultado["notas"].append(notas_bd)
        if dups: resultado["notas"].append(f"DUPLICADO ruta: {', '.join(d['codigo'] for d in dups)}")

        # Si no tiene NINGÚN layer conocido → SIN DATOS (archivo viejo/diferente)
        tiene_algo = any(l in layers_con_ents for l in ("K","K2","K3","LOGO","LOGO1","LOGO2"))
        if not tiene_algo and vitro == "—" and malla == "—":
            resultado["estado"] = "SIN DATOS"
            resultado["notas"] = ["Archivo sin layers/texto estándar"]
        else:
            errores = [n for n in resultado["notas"] if any(p in n for p in ("Falta","incorrecto","DESACTUALIZADO"))]
            resultado["estado"] = "INCOMPLETO" if errores else "OK"

    except Exception as e:
        resultado["estado"] = "ERROR"
        resultado["error"]  = str(e)
        resultado["notas"]  = [f"ezdxf error: {e}"]

    resultado["notas"] = " | ".join(resultado["notas"]) if resultado["notas"] else ""
    return resultado


# ═══════════════════════════════════════════════════════════════
#  ANÁLISIS DE UN DWG
# ═══════════════════════════════════════════════════════════════
def analizar_dwg(acad, ruta_dwg):
    """
    Abre el DWG, extrae toda la info y cierra.
    Retorna dict con los resultados.
    """
    resultado = {
        "ruta":          ruta_dwg,
        "archivo":       os.path.basename(ruta_dwg),
        "es_parabrisas": es_parabrisas(os.path.basename(ruta_dwg)),
        # layers
        "k_ok":    False,  "k_color":  None,
        "k2_ok":   False,  "k2_color": None,
        "k3_ok":   False,  "k3_color": None,
        "logo1_ok": False, "logo2_ok": False,
        "puntos_ok": False, "puntos_estado": "—",
        # bd
        "vitro": "—", "malla": "—",
        "bd_actualizado": False,
        "duplicados": [],
        # general
        "estado":  "OK",
        "notas":   [],
        "error":   None,
    }

    RPC_REJECTED = -2147418111  # AutoCAD ocupado — reintentar

    def _com_call(fn, reintentos=8, espera=0.5):
        """Llama fn(), reintenta si AutoCAD está ocupado (RPC_E_CALL_REJECTED)."""
        for i in range(reintentos):
            try:
                return fn()
            except Exception as e:
                if getattr(e, 'hresult', None) == RPC_REJECTED or RPC_REJECTED == getattr(e, 'args', [None])[0]:
                    time.sleep(espera * (i + 1))
                else:
                    raise
        raise RuntimeError("AutoCAD rechazó la llamada demasiadas veces")

    doc = None
    try:
        doc = _com_call(lambda: acad.Documents.Open(os.path.abspath(ruta_dwg), False, True))
        # Esperar a que el doc esté listo (polling, máx 6s)
        for _ in range(60):
            try:
                if _com_call(lambda: doc.FullName, reintentos=3, espera=0.3):
                    break
            except Exception:
                pass
            time.sleep(0.1)
        # Esperar a que ModelSpace esté disponible (puede tardar en DWGs pesados)
        msp = None
        for _ in range(20):
            try:
                msp = _com_call(lambda: doc.ModelSpace)
                _ = msp.Count  # forzar acceso real para confirmar que está listo
                break
            except Exception:
                time.sleep(0.3)
        if msp is None:
            raise RuntimeError("ModelSpace no disponible tras espera")

        # ── Recolectar layers ──────────────────────────────────────────
        layers = {}   # name_upper → color ACI
        for i in range(doc.Layers.Count):
            try:
                lyr = doc.Layers.Item(i)
                layers[lyr.Name.upper().strip()] = lyr.Color
            except Exception:
                pass

        # ── Helpers de extracción de texto ────────────────────────────
        _RE_MTEXT = re.compile(r'\{[^}]*\}|\\[A-Za-z][^;]*;|%%.')

        def _limpiar_mtext(s):
            """Quita códigos de formato de MTEXT ({\\fArial|...}, \\P, %%c, etc.)"""
            s = _RE_MTEXT.sub(" ", s)
            return re.sub(r'\s+', ' ', s).strip()

        _bloques_visitados = set()

        def _texto_de_entidad(ent):
            """Extrae todo el texto de una entidad, cualquier tipo."""
            n = ent.ObjectName
            # TEXT / MTEXT / ATTDEF / ATTRIB
            if "Text" in n or "Attrib" in n or "Attdef" in n:
                try:
                    t = _limpiar_mtext(ent.TextString)
                    if t: textos.append(t)
                except Exception: pass
            # Tabla (AcDbTable) — leer cada celda
            if n == "AcDbTable":
                try:
                    for row in range(ent.Rows):
                        for col in range(ent.Columns):
                            try:
                                t = _limpiar_mtext(ent.GetText(row, col))
                                if t: textos.append(t)
                            except Exception: pass
                except Exception: pass
            # MLeader
            if n == "AcDbMLeader":
                try:
                    t = _limpiar_mtext(ent.TextString)
                    if t: textos.append(t)
                except Exception: pass

        def _leer_bloque_def(nombre_blk):
            """Lee recursivamente todas las entidades de una definición de bloque."""
            if nombre_blk in _bloques_visitados:
                return
            _bloques_visitados.add(nombre_blk)
            try:
                blk_def = doc.Blocks.Item(nombre_blk)
                for be in blk_def:
                    try:
                        _texto_de_entidad(be)
                        if be.ObjectName == "AcDbBlockReference":
                            # Atributos del sub-bloque
                            try:
                                for attr in be.GetAttributes():
                                    _texto_de_entidad(attr)
                            except Exception: pass
                            # Definición del sub-bloque (recursivo)
                            try:
                                _leer_bloque_def(be.Name)
                            except Exception: pass
                    except Exception: pass
            except Exception: pass

        # ── Layers con entidades reales ────────────────────────────────
        layers_con_ents = set()
        textos          = []
        hatch_en_puntos = False
        trazo_en_puntos = False

        for e in msp:
            try:
                nombre_obj = e.ObjectName
                lyr_e = e.Layer.upper().strip()
                layers_con_ents.add(lyr_e)

                # Texto directo
                _texto_de_entidad(e)

                # Bloque → atributos + definición completa (recursivo)
                if nombre_obj == "AcDbBlockReference":
                    try:
                        for attr in e.GetAttributes():
                            _texto_de_entidad(attr)
                    except Exception: pass
                    try:
                        _leer_bloque_def(e.Name)
                    except Exception: pass

                # PUNTOS: hatch vs trazo suelto
                if "PUNTOS" in lyr_e:
                    if nombre_obj == "AcDbHatch":
                        hatch_en_puntos = True
                    elif nombre_obj in ("AcDbPolyline","AcDb2dPolyline","AcDbLine","AcDbSpline"):
                        trazo_en_puntos = True

            except Exception:
                pass

        # ── Validar K / K2 / K3 ───────────────────────────────────────
        for key, color_esperado, campo_ok, campo_color in [
            ("K",  COLOR_K,  "k_ok",  "k_color"),
            ("K2", COLOR_K2, "k2_ok", "k2_color"),
            ("K3", COLOR_K3, "k3_ok", "k3_color"),
        ]:
            color_real = layers.get(key)
            resultado[campo_color] = color_real
            if key in layers_con_ents and color_real == color_esperado:
                resultado[campo_ok] = True
            elif key in layers and key not in layers_con_ents:
                resultado["notas"].append(f"Layer {key} vacío")
            elif key in layers and color_real != color_esperado:
                resultado["notas"].append(f"Layer {key} color incorrecto ({color_real}≠{color_esperado})")
            else:
                resultado["notas"].append(f"Falta layer {key}")

        # ── Logos (aplica a TODOS los artes) ──────────────────────────
        # Un solo logo → layer "LOGO"; dos logos → layers "LOGO1" y "LOGO2"
        tiene_logo_simple = "LOGO" in layers_con_ents
        tiene_logo1       = "LOGO1" in layers_con_ents
        tiene_logo2       = "LOGO2" in layers_con_ents
        resultado["logo1_ok"] = tiene_logo_simple or tiene_logo1
        resultado["logo2_ok"] = tiene_logo2
        if not resultado["logo1_ok"]:
            resultado["notas"].append("Falta layer LOGO / LOGO1")

        # ── PUNTOS (solo parabrisas) ───────────────────────────────────
        if resultado["es_parabrisas"]:
            tiene_puntos = any("PUNTOS" in l for l in layers_con_ents)
            resultado["puntos_ok"] = tiene_puntos
            if not tiene_puntos:
                resultado["puntos_estado"] = "FALTA"
                resultado["notas"].append("Parabrisas sin layer PUNTOS")
            elif hatch_en_puntos and not trazo_en_puntos:
                resultado["puntos_estado"] = "OK relleno"
            elif trazo_en_puntos and not hatch_en_puntos:
                resultado["puntos_estado"] = "DESACTUALIZADO (trazo suelto)"
                resultado["notas"].append("PUNTOS: trazo suelto sin relleno")
            elif hatch_en_puntos and trazo_en_puntos:
                resultado["puntos_estado"] = "MIXTO"
                resultado["notas"].append("PUNTOS: mezcla hatch+trazo")
            else:
                resultado["puntos_estado"] = "vacío"
        else:
            resultado["puntos_estado"] = "—"  # no aplica

        # ── BD: buscar vitro/malla en textos ──────────────────────────
        vitro, malla, bd_ok, notas_bd, dups = bd_validar_y_actualizar(textos, ruta_dwg)
        resultado["vitro"]          = vitro
        resultado["malla"]          = malla
        resultado["bd_actualizado"] = bd_ok
        resultado["duplicados"]     = dups
        if notas_bd:
            resultado["notas"].append(notas_bd)
        if dups:
            resultado["notas"].append(f"DUPLICADO ruta: {', '.join(d['codigo'] for d in dups)}")

        # ── Estado general ─────────────────────────────────────────────
        errores = [n for n in resultado["notas"] if "Falta" in n or "incorrecto" in n or "DESACTUALIZADO" in n]
        if errores:
            resultado["estado"] = "INCOMPLETO"
        else:
            resultado["estado"] = "OK"

    except Exception as e:
        resultado["error"]  = str(e)
        resultado["estado"] = "ERROR"
        resultado["notas"].append(f"Error: {e}")
    finally:
        if doc:
            try:
                doc.Close(False)
                time.sleep(0.3)
            except Exception:
                pass

    resultado["notas"] = " | ".join(resultado["notas"]) if resultado["notas"] else ""
    return resultado


# ═══════════════════════════════════════════════════════════════
#  RECORRER ÁRBOL DE CARPETAS
# ═══════════════════════════════════════════════════════════════
def _excluir(nombre):
    n = nombre.upper()
    return any(ex in n for ex in EXCLUIR_CARPETAS)

def _extraer_contexto(ruta_artes, ruta_base_marca):
    """
    A partir de la ruta de la carpeta ARTES extrae (vehiculo, version).
    Toma los dos niveles justo antes de ARTES como vehículo y versión.
    """
    partes = []
    p = os.path.dirname(ruta_artes)   # carpeta padre de ARTES
    for _ in range(6):
        cabeza, cola = os.path.split(p)
        if not cola or p == ruta_base_marca or p == cabeza:
            break
        partes.insert(0, cola)
        p = cabeza
    if len(partes) >= 2:
        return partes[-2], partes[-1]
    elif len(partes) == 1:
        return partes[0], ""
    return "", ""

def listar_dwgs(ruta_base, solo_marca=None):
    """
    Genera (marca, vehiculo, version, carpeta_origen, ruta_dwg).
    Busca recursivamente cualquier carpeta ARTES y toma:
      - P*.dwg directamente en ARTES/
      - P*.dwg dentro de ARTES/BN/ (o cualquier subcarpeta BN)
    """
    try:
        marcas = sorted(os.listdir(ruta_base))
    except Exception as e:
        log_err(f"No se puede leer {ruta_base}: {e}")
        return

    for marca in marcas:
        if solo_marca and marca.upper() != solo_marca.upper():
            continue
        ruta_marca = os.path.join(ruta_base, marca)
        if not os.path.isdir(ruta_marca):
            continue

        # Walk recursivo bajo la marca
        for dirpath, dirnames, filenames in os.walk(ruta_marca):
            # Excluir subcarpetas descartadas para no bajar a ellas
            dirnames[:] = sorted(d for d in dirnames if not _excluir(d))

            nombre_dir = os.path.basename(dirpath).upper()
            if nombre_dir != "ARTES":
                continue

            vehiculo, version = _extraer_contexto(dirpath, ruta_marca)

            # 1. P*.dwg directamente en ARTES/
            for archivo in sorted(filenames):
                if archivo.lower().endswith((".dwg", ".dxf")) and archivo.upper().startswith("P"):
                    if not _excluir(archivo):
                        yield marca, vehiculo, version, dirpath, os.path.join(dirpath, archivo)

            # 2. Subcarpeta BN dentro de ARTES/ → P*.dwg
            for sub in sorted(os.listdir(dirpath)):
                if sub.upper() == "BN":
                    carpeta_bn = os.path.join(dirpath, sub)
                    if not os.path.isdir(carpeta_bn):
                        continue
                    for archivo in sorted(os.listdir(carpeta_bn)):
                        if archivo.lower().endswith((".dwg", ".dxf")) and archivo.upper().startswith("P"):
                            if not _excluir(archivo):
                                yield marca, vehiculo, version, carpeta_bn, os.path.join(carpeta_bn, archivo)


# ═══════════════════════════════════════════════════════════════
#  CHECKPOINT
# ═══════════════════════════════════════════════════════════════
def cp_cargar():
    if not os.path.exists(CHECKPOINT_FILE):
        return {}
    try:
        with open(CHECKPOINT_FILE, encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}

def cp_guardar(data):
    try:
        with open(CHECKPOINT_FILE, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    except Exception as e:
        log_warn(f"No se pudo guardar checkpoint: {e}")


# ═══════════════════════════════════════════════════════════════
#  EXCEL — generar reporte
# ═══════════════════════════════════════════════════════════════
COLS = [
    ("Archivo",       28), ("Marca",       14), ("Vehículo",     36),
    ("Versión",       10), ("Parabrisas",  12), ("K",             8),
    ("K2",             8), ("K3",           8), ("Logo1",        10),
    ("Logo2",         10), ("Puntos",      22), ("Vitro",        12),
    ("Malla",         18), ("BD Update",  10), ("Estado",       14),
    ("Notas",         55), ("Ruta",        70),
]

C_HDR  = "1F4E79"
C_OK   = "C6EFCE"
C_ERR  = "FFCCCC"
C_WARN = "FFE699"
C_PARA = "D6E4F0"
C_ALT  = "F2F2F2"

def _fill(hex_color): return PatternFill("solid", fgColor=hex_color)
def _font(bold=False, color="000000", size=10):
    return Font(bold=bold, color=color, size=size)
def _center(): return Alignment(horizontal="center", vertical="center", wrap_text=True)
def _left():   return Alignment(horizontal="left",   vertical="center", wrap_text=True)

def _borde():
    s = Side(style="thin", color="CCCCCC")
    return Border(left=s, right=s, top=s, bottom=s)

def generar_excel(filas_por_marca, ruta_excel):
    wb = openpyxl.Workbook()
    wb.remove(wb.active)  # quitar hoja default

    totales_resumen = []

    for marca, filas in sorted(filas_por_marca.items()):
        ws = wb.create_sheet(title=marca[:28])
        ws.freeze_panes = "A2"
        ws.row_dimensions[1].height = 28

        # Encabezados
        for ci, (titulo, ancho) in enumerate(COLS, 1):
            cel = ws.cell(row=1, column=ci, value=titulo)
            cel.font      = _font(bold=True, color="FFFFFF", size=10)
            cel.fill      = _fill(C_HDR)
            cel.alignment = _center()
            cel.border    = _borde()
            ws.column_dimensions[get_column_letter(ci)].width = ancho

        ok = inc = err = bd_upd = 0
        for ri, f in enumerate(filas, 2):
            estado = f.get("estado", "")
            para   = f.get("es_parabrisas", False)

            fila_color = (C_ERR  if estado == "ERROR"
                          else C_WARN if estado == "INCOMPLETO"
                          else C_PARA if para
                          else (C_ALT  if ri % 2 == 0 else "FFFFFF"))

            vals = [
                f.get("archivo",""),
                f.get("marca",""),
                f.get("vehiculo",""),
                f.get("version",""),
                "✔" if para else "",
                "✔" if f.get("k_ok")  else "✘",
                "✔" if f.get("k2_ok") else "✘",
                "✔" if f.get("k3_ok") else "✘",
                "✔" if f.get("logo1_ok") else "✘",
                "✔" if f.get("logo2_ok") else "—",
                f.get("puntos_estado","—"),
                f.get("vitro","—"),
                f.get("malla","—"),
                "✔" if f.get("bd_actualizado") else "—",
                estado,
                f.get("notas",""),
                f.get("ruta",""),
            ]
            for ci, val in enumerate(vals, 1):
                cel = ws.cell(row=ri, column=ci, value=val)
                cel.fill      = _fill(fila_color)
                cel.border    = _borde()
                cel.alignment = _left()
                if ci in (6,7,8,9,10,14):  # columnas de check
                    cel.alignment = _center()
                    if str(val) == "✘":
                        cel.font = _font(bold=True, color="C00000")
                    elif str(val) == "✔":
                        cel.font = _font(bold=True, color="375623")

            if estado == "OK":      ok  += 1
            elif estado == "INCOMPLETO": inc += 1
            else:                   err += 1
            if f.get("bd_actualizado"): bd_upd += 1

        total = len(filas)
        totales_resumen.append({
            "marca": marca, "total": total,
            "ok": ok, "incompleto": inc, "error": err, "bd_upd": bd_upd,
            "pct": f"{ok/total*100:.1f}%" if total else "0%",
        })

    # Hoja RESUMEN
    ws = wb.create_sheet(title="RESUMEN", index=0)
    ws.freeze_panes = "A2"
    hdr_res = ["Marca","Total","OK","Incompleto","Error","BD Update","% OK"]
    for ci, h in enumerate(hdr_res, 1):
        cel = ws.cell(row=1, column=ci, value=h)
        cel.font = _font(bold=True, color="FFFFFF", size=10)
        cel.fill = _fill(C_HDR)
        cel.alignment = _center()
        cel.border = _borde()
    anchos_res = [22, 9, 9, 12, 9, 12, 9]
    for ci, w in enumerate(anchos_res, 1):
        ws.column_dimensions[get_column_letter(ci)].width = w

    tot_total = tot_ok = tot_inc = tot_err = tot_bd = 0
    for ri, t in enumerate(totales_resumen, 2):
        fila = [t["marca"], t["total"], t["ok"], t["incompleto"], t["error"], t["bd_upd"], t["pct"]]
        for ci, val in enumerate(fila, 1):
            cel = ws.cell(row=ri, column=ci, value=val)
            cel.border = _borde()
            cel.alignment = _center() if ci > 1 else _left()
        tot_total += t["total"]; tot_ok += t["ok"]
        tot_inc += t["incompleto"]; tot_err += t["error"]; tot_bd += t["bd_upd"]

    # Fila total
    ri_total = len(totales_resumen) + 2
    fila_tot = ["TOTAL", tot_total, tot_ok, tot_inc, tot_err, tot_bd,
                f"{tot_ok/tot_total*100:.1f}%" if tot_total else "0%"]
    for ci, val in enumerate(fila_tot, 1):
        cel = ws.cell(row=ri_total, column=ci, value=val)
        cel.font = _font(bold=True)
        cel.fill = _fill("D9E1F2")
        cel.border = _borde()
        cel.alignment = _center() if ci > 1 else _left()

    # Hoja DUPLICADOS — vitro/malla con ruta anterior diferente
    todos_dups = []
    for filas in filas_por_marca.values():
        for f in filas:
            for d in f.get("duplicados", []):
                todos_dups.append({
                    "codigo":        d["codigo"],
                    "tipo":          d["tipo"],
                    "marca":         f.get("marca",""),
                    "vehiculo":      f.get("vehiculo",""),
                    "version":       f.get("version",""),
                    "ruta_anterior": d["ruta_anterior"],
                    "ruta_nueva":    d["ruta_nueva"],
                    "archivo":       f.get("archivo",""),
                })

    if todos_dups:
        ws_dup = wb.create_sheet(title="DUPLICADOS")
        ws_dup.freeze_panes = "A2"
        hdr_dup = ["Código","Tipo","Marca","Vehículo","Versión","Archivo","Ruta anterior (BD)","Ruta nueva (DWG actual)"]
        anchos_dup = [14, 14, 16, 36, 10, 36, 70, 70]
        for ci, (h, w) in enumerate(zip(hdr_dup, anchos_dup), 1):
            cel = ws_dup.cell(row=1, column=ci, value=h)
            cel.font = _font(bold=True, color="FFFFFF", size=10)
            cel.fill = _fill("7B2D00")   # naranja oscuro para llamar la atención
            cel.alignment = _center()
            cel.border = _borde()
            ws_dup.column_dimensions[get_column_letter(ci)].width = w
        for ri, d in enumerate(todos_dups, 2):
            vals = [d["codigo"], d["tipo"], d["marca"], d["vehiculo"],
                    d["version"], d["archivo"], d["ruta_anterior"], d["ruta_nueva"]]
            for ci, val in enumerate(vals, 1):
                cel = ws_dup.cell(row=ri, column=ci, value=val)
                cel.fill = _fill("FFF2CC" if ri % 2 == 0 else "FFFAF0")
                cel.border = _borde()
                cel.alignment = _left()

    # Hoja ACTUALIZADAS — solo las que se insertaron/actualizaron en BD
    todas_filas = [f for filas in filas_por_marca.values() for f in filas]
    actualizadas = [f for f in todas_filas if f.get("bd_actualizado")]

    if actualizadas:
        ws_act = wb.create_sheet(title="ACTUALIZADAS")
        ws_act.freeze_panes = "A2"
        ws_act.row_dimensions[1].height = 28

        COLS_ACT = [
            ("Archivo",    28), ("Marca",    14), ("Vehículo",  36),
            ("Versión",    10), ("Vitro",    16), ("Malla G",   16),
            ("Malla P",    18), ("Estado",   14), ("K",          8),
            ("K2",          8), ("K3",        8), ("Logo",       8),
            ("Parabrisas", 12), ("Puntos",   22), ("Ruta",      80),
        ]
        for ci, (titulo, ancho) in enumerate(COLS_ACT, 1):
            cel = ws_act.cell(row=1, column=ci, value=titulo)
            cel.font = _font(bold=True, color="FFFFFF", size=10)
            cel.fill = _fill("375623")   # verde oscuro
            cel.alignment = _center()
            cel.border = _borde()
            ws_act.column_dimensions[get_column_letter(ci)].width = ancho

        for ri, f in enumerate(actualizadas, 2):
            texto_unido = " ".join([f.get("vitro",""), f.get("malla","")])
            grandes = re.findall(r'A-\d{4,6}', texto_unido, re.I)
            pequenas_raw = re.findall(r'\b\d{4,6}\b', texto_unido)
            excluir = {re.sub(r'[TA]-','',v,flags=re.I) for v in re.findall(r'[TA]-\d{4,6}', texto_unido, re.I)}
            pequenas = [n for n in pequenas_raw if n not in excluir]

            para = f.get("es_parabrisas", False)
            estado = f.get("estado","")
            fila_color = "E2EFDA" if ri % 2 == 0 else "F0FAF0"

            vals = [
                f.get("archivo",""),
                f.get("marca",""),
                f.get("vehiculo",""),
                f.get("version",""),
                f.get("vitro","—"),
                " / ".join(grandes) or "—",
                " / ".join(pequenas) or "—",
                estado,
                "✔" if f.get("k_ok")  else "✘",
                "✔" if f.get("k2_ok") else "✘",
                "✔" if f.get("k3_ok") else "✘",
                "✔" if f.get("logo1_ok") else "✘",
                "✔" if para else "",
                f.get("puntos_estado","—"),
                f.get("ruta",""),
            ]
            for ci, val in enumerate(vals, 1):
                cel = ws_act.cell(row=ri, column=ci, value=val)
                cel.fill = _fill(fila_color)
                cel.border = _borde()
                cel.alignment = _left()
                if ci in (9,10,11,12,13):
                    cel.alignment = _center()
                    if str(val) == "✘":
                        cel.font = _font(bold=True, color="C00000")
                    elif str(val) == "✔":
                        cel.font = _font(bold=True, color="375623")

        log_ok(f"Hoja ACTUALIZADAS: {len(actualizadas)} artes con BD update")

    wb.save(ruta_excel)
    log_ok(f"Excel guardado: {ruta_excel}")


# ═══════════════════════════════════════════════════════════════
#  MAIN
# ═══════════════════════════════════════════════════════════════
def main():
    parser = argparse.ArgumentParser(description="Auditoría artes AGP")
    parser.add_argument("--marca",      default=None, help="Auditar solo esta marca")
    parser.add_argument("--reiniciar",  action="store_true", help="Borra checkpoint y empieza de cero")
    parser.add_argument("--solo-excel", action="store_true", help="Solo regenera Excel desde checkpoint")
    args = parser.parse_args()

    if args.reiniciar and os.path.exists(CHECKPOINT_FILE):
        os.remove(CHECKPOINT_FILE)
        log_info("Checkpoint borrado — empezando de cero.")

    # Cargar checkpoint
    cp = cp_cargar()
    # Estructura: { "marca::vehiculo::version::carpeta_bn": [filas...] }

    if args.solo_excel:
        log_info("Modo solo-excel — regenerando desde checkpoint...")
        filas_por_marca = {}
        for key, filas in cp.items():
            marca = key.split("::")[0] if "::" in key else key
            filas_por_marca.setdefault(marca, []).extend(filas)
        generar_excel(filas_por_marca, EXCEL_SALIDA)
        return

    # Verificar que AutoCAD esté abierto (cada thread crea su propia conexión COM)
    pythoncom.CoInitialize()
    try:
        win32com.client.GetActiveObject("AutoCAD.Application")
        log_ok("AutoCAD detectado.")
    except Exception:
        log_err("AutoCAD no está abierto. Ábrelo primero.")
        pythoncom.CoUninitialize()
        sys.exit(1)


    print(f"\n{'='*60}")
    print(f"  AUDITORÍA ARTES AGP — {datetime.now().strftime('%Y-%m-%d %H:%M')}")
    print(f"  Base: {RUTA_BASE}")
    if args.marca:
        print(f"  Marca filtro: {args.marca}")
    print(f"{'='*60}\n")

    filas_por_marca = {}
    # Precarga del checkpoint
    for key, filas in cp.items():
        marca = key.split("::")[0]
        filas_por_marca.setdefault(marca, []).extend(filas)

    buffer_carpeta = []     # filas del grupo actual (misma carpeta BN)
    clave_carpeta  = None
    n_archivo      = 0
    n_desde_cp     = 0

    try:
        for marca, vehiculo, version, carpeta_bn, ruta_dwg in listar_dwgs(RUTA_BASE, args.marca):

            clave = f"{marca}::{vehiculo}::{version}::{carpeta_bn}"

            # Ya procesado en checkpoint anterior → saltar
            if clave in cp and any(f["ruta"] == ruta_dwg for f in cp[clave]):
                continue

            # Cambio de carpeta → guardar buffer anterior
            if clave != clave_carpeta:
                if buffer_carpeta and clave_carpeta:
                    cp.setdefault(clave_carpeta, []).extend(buffer_carpeta)
                    cp_guardar(cp)
                buffer_carpeta = []
                clave_carpeta  = clave
                log_info(f"→ {marca} / {vehiculo} / {version}")

            n_archivo += 1
            n_desde_cp += 1
            log_info(f"  [{n_archivo}] {os.path.basename(ruta_dwg)}")

            # .dxf → ezdxf directo, sin AutoCAD
            if ruta_dwg.lower().endswith(".dxf") and _EZDXF_OK:
                log_info(f"  [ezdxf directo]")
                resultado = analizar_ezdxf(ruta_dwg)
                resultado["marca"]    = marca
                resultado["vehiculo"] = vehiculo
                resultado["version"]  = version
                buffer_carpeta.append(resultado)
                filas_por_marca.setdefault(marca, []).append(resultado)
                est  = resultado["estado"]
                r    = resultado
                para = r["es_parabrisas"]
                def _c(ok): return "✔" if ok else "✘"
                layers_str = (f"K={_c(r['k_ok'])}  K2={_c(r['k2_ok'])}  K3={_c(r['k3_ok'])}  Logo={_c(r['logo1_ok'])}")
                bd_str = f"  vitro={r['vitro']}  malla={r['malla']}  BD={'✔' if r['bd_actualizado'] else '—'}" if r['vitro'] != "—" or r['malla'] != "—" else "  vitro=—  malla=—"
                prefijo = "  [OK]" if est == "OK" else ("  [SIN DATOS]" if est == "SIN DATOS" else ("  [!!]" if est == "INCOMPLETO" else "  [XX]"))
                log_info(f"{prefijo}  {layers_str}{bd_str}" + (f"\n      ⚠  {r['notas']}" if r["notas"] else ""))
                if n_desde_cp >= CHECKPOINT_CADA:
                    cp.setdefault(clave_carpeta, []).extend(buffer_carpeta)
                    buffer_carpeta = []
                    cp_guardar(cp)
                    n_desde_cp = 0
                continue

            # .dwg → AutoCAD COM con thread + timeout
            _res = [None]
            def _run():
                _ERRORES_RETRIABLES = {
                    -2147418111,  # RPC_E_CALL_REJECTED — AutoCAD ocupado
                    -2147221021,  # CO_E_NOTINITIALIZED — COM no listo
                    -2147418113,  # RPC_E_DISCONNECTED
                    -2147023174,  # RPC_S_SERVER_UNAVAILABLE
                }
                try:
                    pythoncom.CoInitialize()
                    ultimo_error = None
                    for intento in range(6):
                        try:
                            _acad = win32com.client.GetActiveObject("AutoCAD.Application")
                            # Verificar que Documents esté accesible
                            _ = _acad.Documents.Count
                            _res[0] = analizar_dwg(_acad, ruta_dwg)
                            return
                        except Exception as e:
                            ultimo_error = e
                            codigo = getattr(e, 'hresult', None) or (e.args[0] if e.args else None)
                            if codigo in _ERRORES_RETRIABLES or "AutoCAD.Application" in str(e):
                                time.sleep(1.5 * (intento + 1))
                            else:
                                raise
                    raise ultimo_error
                except Exception as e:
                    _res[0] = {"ruta": ruta_dwg, "archivo": os.path.basename(ruta_dwg),
                               "es_parabrisas": es_parabrisas(os.path.basename(ruta_dwg)),
                               "k_ok": False, "k_color": None, "k2_ok": False, "k2_color": None,
                               "k3_ok": False, "k3_color": None, "logo1_ok": False, "logo2_ok": False,
                               "puntos_ok": False, "puntos_estado": "—",
                               "vitro": "—", "malla": "—", "bd_actualizado": False, "duplicados": [],
                               "estado": "ERROR", "notas": f"Error hilo: {e}", "error": str(e)}
                finally:
                    try: pythoncom.CoUninitialize()
                    except Exception: pass

            t = threading.Thread(target=_run, daemon=True)
            t.start()
            t.join(timeout=TIMEOUT_ARCHIVO)
            if t.is_alive():
                log_warn(f"  TIMEOUT ({TIMEOUT_ARCHIVO}s) — saltando {os.path.basename(ruta_dwg)}")
                _res[0] = None

            # Si AutoCAD falló o timeout → intentar ezdxf SOLO si es .dxf
            es_dxf = ruta_dwg.lower().endswith(".dxf")
            if (_res[0] is None or _res[0].get("estado") == "ERROR") and _EZDXF_OK and es_dxf:
                try:
                    log_info(f"  [ezdxf fallback] {os.path.basename(ruta_dwg)}")
                    _res[0] = analizar_ezdxf(ruta_dwg)
                except Exception as ez_e:
                    log_warn(f"  ezdxf también falló: {ez_e}")

            resultado = _res[0] or {"ruta": ruta_dwg, "archivo": os.path.basename(ruta_dwg),
                                    "estado": "ERROR", "notas": "Sin resultado (COM+ezdxf)", "error": "none",
                                    "es_parabrisas": False, "k_ok": False, "k_color": None,
                                    "k2_ok": False, "k2_color": None, "k3_ok": False, "k3_color": None,
                                    "logo1_ok": False, "logo2_ok": False, "puntos_ok": False,
                                    "puntos_estado": "—", "vitro": "—", "malla": "—",
                                    "bd_actualizado": False, "duplicados": []}
            resultado["marca"]    = marca
            resultado["vehiculo"] = vehiculo
            resultado["version"]  = version

            time.sleep(0.3)

            buffer_carpeta.append(resultado)
            filas_por_marca.setdefault(marca, []).append(resultado)

            est  = resultado["estado"]
            r    = resultado
            para = r["es_parabrisas"]

            # ── línea de layers ────────────────────────────────────────
            def _c(ok): return "✔" if ok else "✘"
            layers_str = (f"K={_c(r['k_ok'])}({r['k_color']})  "
                          f"K2={_c(r['k2_ok'])}({r['k2_color']})  "
                          f"K3={_c(r['k3_ok'])}({r['k3_color']})  "
                          f"Logo={_c(r['logo1_ok'])}")
            if r["logo2_ok"]:
                layers_str += "+Logo2✔"
            if para:
                layers_str += f"  Puntos={r['puntos_estado']}"

            # ── línea de BD ────────────────────────────────────────────
            bd_str = ""
            if r["vitro"] != "—" or r["malla"] != "—":
                bd_str = (f"  vitro={r['vitro']}  malla={r['malla']}  "
                          f"BD={'✔' if r['bd_actualizado'] else '✘'}")
                if r.get("duplicados"):
                    bd_str += f"  ⚠ dup:{','.join(d['codigo'] for d in r['duplicados'])}"
            else:
                bd_str = "  vitro=—  malla=—"

            # ── notas de fallo ─────────────────────────────────────────
            notas_str = ""
            if r["notas"]:
                notas_str = f"\n      ⚠  {r['notas']}"

            prefijo = "  [OK]" if est == "OK" else ("  [!!]" if est == "INCOMPLETO" else "  [XX]")
            log_info(f"{prefijo}  {layers_str}{bd_str}{notas_str}")

            # Checkpoint cada N archivos
            if n_desde_cp >= CHECKPOINT_CADA:
                if buffer_carpeta and clave_carpeta:
                    cp.setdefault(clave_carpeta, []).extend(buffer_carpeta)
                    buffer_carpeta = []
                    cp_guardar(cp)
                    log_info(f"  [checkpoint guardado — {n_archivo} archivos procesados]")
                n_desde_cp = 0

    except KeyboardInterrupt:
        log_warn("Interrumpido por el usuario.")
    finally:
        # Guardar lo que queda en buffer
        if buffer_carpeta and clave_carpeta:
            cp.setdefault(clave_carpeta, []).extend(buffer_carpeta)
        cp_guardar(cp)
        pythoncom.CoUninitialize()

    print(f"\n{'='*60}")
    log_ok(f"Procesados: {n_archivo} archivos")
    log_info("Generando Excel...")
    generar_excel(filas_por_marca, EXCEL_SALIDA)
    print(f"{'='*60}\n")


if __name__ == "__main__":
    main()
