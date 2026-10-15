# -*- coding: utf-8 -*-
"""
Lógica de negocio de AGP Glass Suite SIN interfaz gráfica.

Todo lo que antes vivía dentro de las clases Tk de agp_app.py (consultas,
guardado con trazabilidad, inserción manual, scanner, roles, helpers de
artes) está copiado acá tal cual — misma SQL, mismas cascadas, mismos
formatos de consecutivo. La UI nueva (agp_web.py + ui_web/) solo llama
estas funciones; agp_app.py sigue intacto como respaldo.

La conexión a Azure (AGP_Ingenieria) se reutiliza de db_app.importar_excel
(la misma que usan asignaciones.py y logos_laser.py). Las bases externas del
scanner (Comercial, SAP, Calendario) vienen de db_app.conexiones_externas.
"""

import os, re, time, shutil, threading

from db_app.importar_excel import conectar as db_connect

try:
    import win32com.client, pythoncom
    _COM_OK = True
except ImportError:
    _COM_OK = False

try:
    from autocad_ops import AutoCADMotor
    _MOTOR_OK = True
except Exception:
    _MOTOR_OK = False

try:
    from crear_arte_acad import pipeline as _pipeline_acad, ErrorGuardadoArte
    _PIPELINE_OK = True
except Exception:
    _PIPELINE_OK = False
    class ErrorGuardadoArte(RuntimeError): pass

try:
    from db_app.asignaciones import anular_asignacion as _anular_asignacion
    _ASIGN_OK = True
except Exception:
    _ASIGN_OK = False

_SENTINEL = object()


def db_query(sql, params=()):
    conn = db_connect()
    try:
        cur = conn.cursor()
        cur.execute(sql, params)
        cols = [c[0] for c in cur.description]
        return [dict(zip(cols, r)) for r in cur.fetchall()]
    finally:
        conn.close()


def _txt(v):
    """Valor de BD → texto para la UI (sin saltos de línea, None → "")."""
    if v is None:
        return ""
    if hasattr(v, "strftime"):
        return v.strftime("%Y-%m-%d %H:%M")
    return str(v).replace("\r\n", " ").replace("\n", " ").replace("\r", " ")


# ══════════════════════════════════════════════════════════════════════════════
#  ARTES — helpers (idénticos a agp_app.py)
# ══════════════════════════════════════════════════════════════════════════════
def ruta_planos(ruta_dwg):
    dest = os.path.join(os.path.dirname(os.path.abspath(ruta_dwg)), "PLANOS")
    os.makedirs(dest, exist_ok=True)
    return dest


def ruta_arte_salida(ruta_dwg, malla="", pieza="", nombre_archivo=""):
    # Siempre guarda en ARTES/BN/ — crea las carpetas si no existen
    artes = os.path.join(os.path.dirname(os.path.abspath(ruta_dwg)), "ARTES")
    dest  = os.path.join(artes, "BN")
    os.makedirs(dest, exist_ok=True)

    def _limpiar_nombre(s):
        for c in r'/\:*?"<>|':
            s = s.replace(c, "-")
        while "  " in s: s = s.replace("  ", " ")
        while "--" in s: s = s.replace("--", "-")
        return s.strip("- ")

    if nombre_archivo:
        base = nombre_archivo if nombre_archivo.lower().endswith(".dwg") else nombre_archivo + ".dwg"
        nombre = _limpiar_nombre(os.path.splitext(base)[0]) + ".dwg"
    else:
        mallas_str = " - ".join(
            m.strip() for m in malla.replace(",", "/").split("/") if m.strip()
        ) if malla.strip() else ""
        if mallas_str and pieza.strip():
            nombre = f"P {mallas_str} {pieza.strip()}.dwg"
        elif mallas_str:
            nombre = f"P {mallas_str}.dwg"
        elif pieza.strip():
            nombre = f"P {pieza.strip()}.dwg"
        else:
            nombre = "P " + os.path.splitext(os.path.basename(ruta_dwg))[0] + ".dwg"
        nombre = _limpiar_nombre(os.path.splitext(nombre)[0]) + ".dwg"

    return os.path.join(dest, nombre)


def extraer_codigos(ruta):
    base   = os.path.splitext(os.path.basename(ruta))[0]
    grupos = re.findall(r'\d+', base)
    if not grupos: return []
    codigos, total = [], 0
    for g in reversed(grupos):
        if total + len(g) > 6: break
        codigos.insert(0, g); total += len(g)
    return codigos


def buscar_artes(ruta, codigos):
    res = []
    for raiz, dirs, archivos in os.walk(ruta):
        dirs[:] = [d for d in dirs if not d.startswith(".")]
        partes = raiz.replace("\\", "/").upper().split("/")
        if not any(p == "ARTES" or p.startswith("ARTES") for p in partes): continue
        for arch in sorted(archivos):
            ext = os.path.splitext(arch)[1].lower()
            if ext not in (".dwg", ".3dm"): continue
            if not arch.upper().startswith("P"): continue
            nums = re.findall(r'\d+', os.path.splitext(arch)[0])
            if codigos and (len(nums) < len(codigos) or nums[-len(codigos):] != codigos): continue
            res.append({"version": os.path.relpath(raiz, ruta),
                        "archivo": arch,
                        "ruta_completa": os.path.join(raiz, arch)})
    return sorted(res, key=lambda x: (x["version"], x["archivo"]))


def crear_arte_autocad(ruta_dwg, log_fn=None, valores_cajetin=None,
                       ruta_salida=None, perim_index=0, _com_ya_init=False,
                       compensar=False):
    """_com_ya_init=True → el llamador ya hizo CoInitialize; no llamar de nuevo."""
    if log_fn is None: log_fn = print
    if not _PIPELINE_OK: raise RuntimeError("Pipeline no disponible")
    if not _COM_OK: raise RuntimeError("pywin32 no disponible")
    if not _com_ya_init:
        pythoncom.CoInitialize()
    try:
        try:
            acad = win32com.client.GetActiveObject("AutoCAD.Application")
        except Exception:
            raise RuntimeError("AutoCAD no está abierto")
        log_fn(f"  Abriendo: {os.path.basename(ruta_dwg)}")
        doc = acad.Documents.Open(os.path.abspath(ruta_dwg), False, False)
        time.sleep(2.5)
        try: doc.Activate(); time.sleep(0.5)
        except Exception: pass
        # Esperar a que AutoCAD termine de cargar el documento
        for _ in range(30):
            try:
                nombre = doc.FullName
                if nombre and os.path.exists(nombre):
                    _ = doc.ModelSpace.Count  # fuerza carga completa
                    break
            except Exception:
                pass
            time.sleep(0.5)
        n = _pipeline_acad(doc, log_fn=log_fn, valores_cajetin=valores_cajetin,
                           ruta_salida=ruta_salida, perim_index=perim_index,
                           compensar=compensar)
        return n or 1
    finally:
        if not _com_ya_init:
            pythoncom.CoUninitialize()


def anular_bd(valores, log_fn):
    """Anula vitro y malla en BD cuando el arte falla — evita números bloqueados."""
    if not _ASIGN_OK:
        return
    vitro = valores.get("VITRO", "").strip()
    malla = valores.get("MALLA", "").strip()
    anulados = []
    try:
        if vitro:
            _anular_asignacion("vitrojet", vitro)
            anulados.append(vitro)
        for m in malla.split("/"):
            m = m.strip()
            if not m:
                continue
            tab = "grandes" if m.upper().startswith("A-") else "pequenas"
            _anular_asignacion(tab, m)
            anulados.append(m)
        if anulados:
            log_fn(f"⚠ Asignación anulada en BD: {', '.join(anulados)}", "warn")
    except Exception as e:
        log_fn(f"WARN: no se pudo anular en BD: {e}", "warn")


# ══════════════════════════════════════════════════════════════════════════════
#  CONSULTAR BD
# ══════════════════════════════════════════════════════════════════════════════
CONSULTA_QUERIES = {
    "vitrojet":  ("SELECT TOP(?) v.vitro,v.codigo_malla,v.tipo_malla,v.bnerig,v.vehiculo,v.version,v.ruta,v.responsable "
                  "FROM mallas.vitrojet v {where} ORDER BY TRY_CAST(SUBSTRING(v.vitro,3,50) AS INT) DESC",
                  ["Vitro","Malla","Tipo","B/N","Vehículo","Versión","Ruta","Responsable"],
                  ["vitro","codigo_malla","tipo_malla","bnerig","vehiculo","version","ruta","responsable"]),
    "grandes":   ("SELECT TOP(?) codigo,cod_veh,descripcion,pieza,tipo,version,ruta_dwg,responsable "
                  "FROM mallas.grandes {where} ORDER BY TRY_CAST(SUBSTRING(codigo,3,50) AS INT) DESC",
                  ["Código","Cód.Veh.","Descripción","Pieza","Tipo","Versión","Ruta","Responsable"],
                  ["codigo","cod_veh","descripcion","pieza","tipo","version","ruta_dwg","responsable"]),
    "pequenas":  ("SELECT TOP(?) codigo,cod_veh,descripcion,pieza,tipo,version,ruta_dwg,responsable "
                  "FROM mallas.pequenas {where} ORDER BY TRY_CAST(codigo AS INT) DESC",
                  ["Código","Cód.Veh.","Descripción","Pieza","Tipo","Versión","Ruta","Responsable"],
                  ["codigo","cod_veh","descripcion","pieza","tipo","version","ruta_dwg","responsable"]),
    "vinilos":   ("SELECT TOP(?) herramental,vehiculo,cod_vehiculo,version,pieza,tipo "
                  "FROM mallas.vinilos {where} ORDER BY herramental DESC",
                  ["Herramental","Vehículo","Cód.Veh.","Versión","Pieza","Tipo"],
                  ["herramental","vehiculo","cod_vehiculo","version","pieza","tipo"]),
    "pasta":     ("SELECT TOP(?) consecutivo,tipo,vehiculo,cod_vehiculo,version,pieza,ruta_archivo,caso "
                  "FROM mallas.pasta_plata {where} ORDER BY consecutivo DESC",
                  ["Consecutivo","RED/ANT","Vehículo","Cód.Veh.","Versión","Pieza","Ruta Archivo","Caso"],
                  ["consecutivo","tipo","vehiculo","cod_vehiculo","version","pieza","ruta_archivo","caso"]),
    "glassjet_viejo": ("SELECT TOP(?) id,malla,glassjet,part_number,tipo,vehiculo,homologacion_vitro "
                       "FROM mallas.glassjet_viejo {where} ORDER BY id DESC",
                       ["ID","Malla","Glassjet","Part Number","Tipo","Vehículo","Homol. Vitro"],
                       ["id","malla","glassjet","part_number","tipo","vehiculo","homologacion_vitro"]),
}
CONSULTA_WHERE = {
    "vitrojet":       "WHERE v.vitro LIKE ? OR v.vehiculo LIKE ? OR v.codigo_malla LIKE ?",
    "grandes":        "WHERE descripcion LIKE ? OR codigo LIKE ? OR cod_veh LIKE ?",
    "pequenas":       "WHERE descripcion LIKE ? OR CAST(codigo AS NVARCHAR) LIKE ? OR cod_veh LIKE ?",
    "vinilos":        "WHERE vehiculo LIKE ? OR herramental LIKE ?",
    "pasta":          "WHERE vehiculo LIKE ? OR consecutivo LIKE ?",
    "glassjet_viejo": "WHERE malla LIKE ? OR glassjet LIKE ? OR vehiculo LIKE ?",
}

LIMITE = 300


def _buscar(queries, wheres, tab, q):
    sql_tpl, headers, fields = queries[tab][:3]
    q = (q or "").strip()
    if q:
        where  = wheres[tab]
        params = (LIMITE,) + (f"%{q}%",) * where.count("?")
        sql    = sql_tpl.format(where=where)
    else:
        params = (LIMITE,)
        sql    = sql_tpl.format(where="")
    rows = db_query(sql, params)
    return headers, fields, rows


def consultar(tab, q=""):
    headers, fields, rows = _buscar(CONSULTA_QUERIES, CONSULTA_WHERE, tab, q)
    return {"headers": headers, "fields": fields,
            "rows": [[_txt(r.get(f)) for f in fields] for r in rows],
            "limite": LIMITE}


STATS_TABLAS = {
    "mallas_grandes":  "mallas.grandes",
    "mallas_pequenas": "mallas.pequenas",
    "vitrojet":        "mallas.vitrojet",
    "pasta_plata":     "mallas.pasta_plata",
    "glassjet_viejo":  "mallas.glassjet_viejo",
    "vinilos":         "mallas.vinilos",
}


def estadisticas():
    conn = db_connect()
    try:
        cur = conn.cursor()
        out = {}
        for key, tabla in STATS_TABLAS.items():
            cur.execute(f"SELECT COUNT(*) FROM {tabla}")
            out[key] = cur.fetchone()[0]
        return out
    finally:
        conn.close()


# ══════════════════════════════════════════════════════════════════════════════
#  GESTIÓN BD
# ══════════════════════════════════════════════════════════════════════════════
GESTION_QUERIES = {
    "vitrojet": (
        "SELECT TOP(?) v.vitro,v.codigo_malla,v.tipo_malla,v.bnerig,v.vehiculo,v.version,"
        "COALESCE(g.pieza,p.pieza) AS pieza,v.ruta,v.responsable,v.estado,v.modificado_por,v.modificado_en "
        "FROM mallas.vitrojet v "
        "LEFT JOIN mallas.grandes g ON v.codigo_malla=g.codigo "
        "LEFT JOIN mallas.pequenas p ON v.codigo_malla=CAST(p.codigo AS NVARCHAR) "
        "{where} ORDER BY TRY_CAST(SUBSTRING(v.vitro,3,50) AS INT) DESC",
        ["Vitro","Malla","Tipo","B/N","Vehículo","Versión","Pieza","Ruta","Responsable","Estado","Modificado por","Modificado en"],
        ["vitro","codigo_malla","tipo_malla","bnerig","vehiculo","version","pieza","ruta","responsable","estado","modificado_por","modificado_en"],
        "vitro",
    ),
    "grandes": (
        "SELECT TOP(?) codigo,cod_veh,descripcion,pieza,tipo,version,ruta_dwg,responsable,estado,modificado_por,modificado_en "
        "FROM mallas.grandes {where} ORDER BY TRY_CAST(SUBSTRING(codigo,3,50) AS INT) DESC",
        ["Código","Cód.Veh.","Descripción","Pieza","Tipo","Versión","Ruta","Responsable","Estado","Modificado por","Modificado en"],
        ["codigo","cod_veh","descripcion","pieza","tipo","version","ruta_dwg","responsable","estado","modificado_por","modificado_en"],
        "codigo",
    ),
    "pequenas": (
        "SELECT TOP(?) codigo,cod_veh,descripcion,pieza,tipo,version,ruta_dwg,responsable,estado,modificado_por,modificado_en "
        "FROM mallas.pequenas {where} ORDER BY TRY_CAST(codigo AS INT) DESC",
        ["Código","Cód.Veh.","Descripción","Pieza","Tipo","Versión","Ruta","Responsable","Estado","Modificado por","Modificado en"],
        ["codigo","cod_veh","descripcion","pieza","tipo","version","ruta_dwg","responsable","estado","modificado_por","modificado_en"],
        "codigo",
    ),
    "vinilos": (
        "SELECT TOP(?) herramental,vehiculo,cod_vehiculo,version,pieza,tipo,ruta,modificado_por,modificado_en "
        "FROM mallas.vinilos {where} ORDER BY herramental DESC",
        ["Herramental","Vehículo","Cód.Veh.","Versión","Pieza","Tipo","Ruta","Modificado por","Modificado en"],
        ["herramental","vehiculo","cod_vehiculo","version","pieza","tipo","ruta","modificado_por","modificado_en"],
        "herramental",
    ),
    "pasta_plata": (
        "SELECT TOP(?) consecutivo,tipo,vehiculo,cod_vehiculo,version,pieza,ruta_archivo,caso,modificado_por,modificado_en "
        "FROM mallas.pasta_plata {where} ORDER BY consecutivo DESC",
        ["Consecutivo","RED/ANT","Vehículo","Cód.Veh.","Versión","Pieza","Ruta archivo","Caso","Modificado por","Modificado en"],
        ["consecutivo","tipo","vehiculo","cod_vehiculo","version","pieza","ruta_archivo","caso","modificado_por","modificado_en"],
        "consecutivo",
    ),
}
GESTION_WHERE = {
    "vitrojet":   "WHERE v.vitro LIKE ? OR v.vehiculo LIKE ? OR v.codigo_malla LIKE ?",
    "grandes":    "WHERE descripcion LIKE ? OR codigo LIKE ? OR cod_veh LIKE ?",
    "pequenas":   "WHERE descripcion LIKE ? OR CAST(codigo AS NVARCHAR) LIKE ? OR cod_veh LIKE ?",
    "vinilos":    "WHERE vehiculo LIKE ? OR herramental LIKE ? OR pieza LIKE ?",
    "pasta_plata":"WHERE vehiculo LIKE ? OR consecutivo LIKE ? OR pieza LIKE ?",
}
# (label, campo). "_pieza_malla" es virtual → cascade a la malla vinculada
EDITABLES = {
    "vitrojet": [("Malla","codigo_malla"),("B/N","bnerig"),("Vehículo","vehiculo"),
                 ("Versión","version"),("Pieza","_pieza_malla"),("Ruta","ruta"),
                 ("Responsable","responsable")],
    "grandes":  [("Cód.Veh.","cod_veh"),("Descripción","descripcion"),("Pieza","pieza"),
                 ("Tipo","tipo"),("Versión","version"),("Ruta","ruta_dwg"),
                 ("Responsable","responsable")],
    "pequenas": [("Cód.Veh.","cod_veh"),("Descripción","descripcion"),("Pieza","pieza"),
                 ("Tipo","tipo"),("Versión","version"),("Ruta","ruta_dwg"),
                 ("Responsable","responsable")],
    "vinilos":  [("Vehículo","vehiculo"),("Cód. Vehículo","cod_vehiculo"),("Versión","version"),
                 ("Pieza","pieza"),("Tipo","tipo"),("Ruta","ruta")],
    "pasta_plata": [("RED/ANT","tipo"),("Vehículo","vehiculo"),("Cód. Vehículo","cod_vehiculo"),
                    ("Versión","version"),("Pieza","pieza"),("Ruta archivo","ruta_archivo"),
                    ("Caso","caso")],
}
TABLA_SQL = {
    "vitrojet":   "mallas.vitrojet",
    "grandes":    "mallas.grandes",
    "pequenas":   "mallas.pequenas",
    "vinilos":    "mallas.vinilos",
    "pasta_plata":"mallas.pasta_plata",
}
PK = {
    "vitrojet":   "vitro",
    "grandes":    "codigo",
    "pequenas":   "codigo",
    "vinilos":    "herramental",
    "pasta_plata":"consecutivo",
}


def opciones_campo(tab, campo):
    """Dropdowns del editor (mismos que en agp_app.py)."""
    if campo == "tipo" and tab == "pasta_plata":
        return ["RED", "ANT"]
    return {"caso": ["Caso 1", "Caso 2"], "bnerig": ["BN", "BNI"]}.get(campo)


def gestion_buscar(tab, q=""):
    headers, fields, rows = _buscar(GESTION_QUERIES, GESTION_WHERE, tab, q)
    return {"headers": headers, "fields": fields, "pk": PK[tab],
            "rows": [{f: _txt(r.get(f)) for f in fields} for r in rows],
            "editables": [{"label": l, "campo": c, "opciones": opciones_campo(tab, c)}
                          for l, c in EDITABLES[tab]],
            "anulable": tab not in ("vinilos", "pasta_plata"),
            "limite": LIMITE}


def guardar_edicion(tab, pk_val, nuevos_raw, usuario_nombre):
    """UPDATE + trazabilidad + cascadas vitro↔malla (idéntico a agp_app.py)."""
    tabla  = TABLA_SQL[tab]
    pk_col = PK[tab]
    permitidos = {c for _, c in EDITABLES[tab]}
    nuevos_raw = {c: ((str(v).strip() or None) if v is not None else None)
                  for c, v in nuevos_raw.items() if c in permitidos}
    if not nuevos_raw:
        return
    pieza_virtual = nuevos_raw.pop("_pieza_malla", _SENTINEL)
    nuevos = nuevos_raw

    cn  = db_connect()
    try:
        cur = cn.cursor()
        valores_ant = {}
        if nuevos:
            campos_sel = ", ".join(nuevos.keys())
            cur.execute(f"SELECT {campos_sel} FROM {tabla} WHERE {pk_col}=?", (pk_val,))
            row_ant = cur.fetchone()
            if row_ant:
                for i, campo in enumerate(nuevos.keys()):
                    valores_ant[campo] = str(row_ant[i]) if row_ant[i] is not None else None
            sets   = [f"{c}=?" for c in nuevos] + ["modificado_por=?", "modificado_en=SYSDATETIME()"]
            params = list(nuevos.values()) + [usuario_nombre, pk_val]
            cur.execute(f"UPDATE {tabla} SET {', '.join(sets)} WHERE {pk_col}=?", params)

        for campo, val_nuevo in nuevos.items():
            val_ant = valores_ant.get(campo)
            val_nuevo_s = str(val_nuevo) if val_nuevo is not None else None
            if val_ant != val_nuevo_s:
                cur.execute(
                    "INSERT INTO MALLAS.TRAZABILIDAD "
                    "(tabla, pk_campo, pk_valor, campo, valor_anterior, valor_nuevo, usuario) "
                    "VALUES (?,?,?,?,?,?,?)",
                    (tabla, pk_col, str(pk_val), campo, val_ant, val_nuevo_s, usuario_nombre))

        # Cascade pieza virtual → malla vinculada
        if pieza_virtual is not _SENTINEL and "vitrojet" in tabla:
            try:
                cur.execute("SELECT codigo_malla FROM mallas.vitrojet WHERE vitro=?", (pk_val,))
                r_cm2 = cur.fetchone()
                if r_cm2 and r_cm2[0]:
                    cod_m2 = str(r_cm2[0]).strip()
                    cur.execute("UPDATE mallas.grandes SET pieza=? WHERE CAST(codigo AS NVARCHAR)=?",
                                (pieza_virtual, cod_m2))
                    if cur.rowcount == 0:
                        cur.execute("UPDATE mallas.pequenas SET pieza=? WHERE CAST(codigo AS NVARCHAR)=?",
                                    (pieza_virtual, cod_m2))
            except Exception:
                pass

        # Cascade ruta vitro ↔ malla vinculada (best-effort)
        try:
            if "vitrojet" in tabla and "ruta" in nuevos and nuevos["ruta"]:
                cur.execute("SELECT codigo_malla FROM mallas.vitrojet WHERE vitro=?", (pk_val,))
                r_cm = cur.fetchone()
                if r_cm and r_cm[0]:
                    cod_m = str(r_cm[0]).strip()
                    cur.execute("UPDATE mallas.grandes SET ruta_dwg=? WHERE CAST(codigo AS NVARCHAR)=?",
                                (nuevos["ruta"], cod_m))
                    if cur.rowcount == 0:
                        cur.execute("UPDATE mallas.pequenas SET ruta_dwg=? WHERE CAST(codigo AS NVARCHAR)=?",
                                    (nuevos["ruta"], cod_m))
            elif ("grandes" in tabla or "pequenas" in tabla) \
                    and "ruta_dwg" in nuevos and nuevos["ruta_dwg"]:
                cur.execute("SELECT vitro FROM mallas.vitrojet WHERE codigo_malla=?", (str(pk_val),))
                r_vt = cur.fetchone()
                if r_vt and r_vt[0]:
                    cur.execute("UPDATE mallas.vitrojet SET ruta=? WHERE vitro=?",
                                (nuevos["ruta_dwg"], r_vt[0]))
        except Exception:
            pass

        cn.commit()
    finally:
        cn.close()


# Campos del formulario de insertar: (label, campo, auto, opciones)
FORM_CAMPOS = {
    "pasta_plata": [
        ("Consecutivo",     "consecutivo",  True,  None),
        ("RED/ANT",         "tipo",         False, ["RED", "ANT"]),
        ("Nombre vehículo", "vehiculo",     False, None),
        ("Cód. Vehículo",   "cod_vehiculo", False, None),
        ("Versión",         "version",      False, None),
        ("Pieza",           "pieza",        False, None),
        ("Ruta archivo",    "ruta_archivo", False, None),
        ("Caso",            "caso",         False, ["Caso 1", "Caso 2"]),
    ],
    "vinilos": [
        ("Herramental",   "herramental",  True,  None),
        ("Vehículo",      "vehiculo",     False, None),
        ("Cód. Vehículo", "cod_vehiculo", False, None),
        ("Versión",       "version",      False, None),
        ("Pieza",         "pieza",        False, None),
        ("Tipo",          "tipo",         False, None),
    ],
}
_PK_AUTO = {"pasta_plata": "consecutivo", "vinilos": "herramental"}


def siguiente_consecutivo(tabla):
    """Vista previa del próximo código (el real lo calcula el INSERT atómico)."""
    try:
        cn = db_connect()
        try:
            cur = cn.cursor()
            if tabla == "pasta_plata":
                cur.execute("SELECT ISNULL(MAX(TRY_CAST(SUBSTRING(consecutivo,3,50) AS INT)),0)+1 FROM mallas.pasta_plata")
                return f"S-{cur.fetchone()[0]:05d}"
            if tabla == "vinilos":
                cur.execute("SELECT ISNULL(MAX(TRY_CAST(SUBSTRING(herramental,4,50) AS INT)),0)+1 FROM mallas.vinilos")
                return f"VC-{cur.fetchone()[0]:04d}"
        finally:
            cn.close()
    except Exception:
        return "?"
    return ""


def formulario_insertar(tabla):
    return {"campos": [{"label": l, "campo": c, "auto": a, "opciones": o}
                       for l, c, a, o in FORM_CAMPOS[tabla]],
            "consecutivo": siguiente_consecutivo(tabla)}


def insertar_manual(tabla, valores):
    """INSERT...SELECT atómico con UPDLOCK+HOLDLOCK (idéntico a agp_app.py).
    Retorna el código insertado."""
    tabla_sql = TABLA_SQL[tabla]
    campos    = FORM_CAMPOS[tabla]
    pk_auto   = _PK_AUTO.get(tabla)
    cols_datos = [c for _, c, _, _ in campos if c != pk_auto]
    vals_datos = [(str(valores.get(c) or "").strip() or None) for c in cols_datos]
    if not any(vals_datos):
        raise ValueError("Completa al menos un campo")

    cn = db_connect()
    try:
        cur = cn.cursor()
        if tabla == "pasta_plata":
            expr_pk = ("'S-' + RIGHT('00000' + CAST("
                       "ISNULL(MAX(TRY_CAST(SUBSTRING(consecutivo,3,50) AS INT)),0)+1"
                       " AS VARCHAR(10)), 5)")
        else:
            expr_pk = ("'VC-' + RIGHT('0000' + CAST("
                       "ISNULL(MAX(TRY_CAST(SUBSTRING(herramental,4,50) AS INT)),0)+1"
                       " AS VARCHAR(10)), 4)")
        all_cols = [pk_auto] + cols_datos
        sql = (f"INSERT INTO {tabla_sql} ({','.join(all_cols)}) "
               f"SELECT {expr_pk}, {','.join(['?'] * len(cols_datos))} "
               f"FROM {tabla_sql} WITH (UPDLOCK, HOLDLOCK)")
        cur.execute(sql, vals_datos)
        cur.execute(f"SELECT TOP 1 {pk_auto} FROM {tabla_sql} ORDER BY {pk_auto} DESC")
        pk_insertado = cur.fetchone()[0]
        cn.commit()
        return str(pk_insertado)
    finally:
        cn.close()


# ══════════════════════════════════════════════════════════════════════════════
#  SCANNER DE ÓRDENES
# ══════════════════════════════════════════════════════════════════════════════
def buscar_orden(orden):
    """Orden → ZFER (Comercial) → BOM SAP (vitro/mallas) + calendario + tipo
    de pieza (SAP) + rutas (Azure). Misma lógica que TabScanner._t_buscar."""
    from db_app.conexiones_externas import (conectar_comercial, conectar_sap,
                                            conectar_calendario)
    orden = str(orden).strip()
    cn = conectar_comercial()
    try:
        cur = cn.cursor()
        cur.execute("SELECT TOP 1 ZFER FROM [dbo].[BI_TAB_FIN_TURNO_SAP] "
                    "WHERE ORDEN = ? AND ZFER IS NOT NULL", (orden,))
        row = cur.fetchone()
    finally:
        cn.close()
    if not row:
        return {"ok": False, "error": f"Orden {orden} no encontrada", "no_encontrada": True}
    zfer = str(row[0]).strip()

    cn2 = conectar_sap()
    try:
        cur2 = cn2.cursor()
        cur2.execute("SELECT DENOMINACION_COMPONENTE, TEXTO1, TEXTO2 "
                     "FROM ODATA_ZFER_BOM WHERE MATERIAL = ?", (zfer,))
        filas = cur2.fetchall()
    finally:
        cn2.close()

    _EXCLUIR = ("M", "G")
    _EXCLUIR_DESC = ("MOLDE LLENO", "PLANTILLA", "ANTENA")
    mallas, t1s, t2s = [], [], []
    for comp, t1, t2 in filas:
        if comp and str(comp).strip():
            c = str(comp).strip()
            cu = c.upper()
            if not cu.startswith(_EXCLUIR) and not any(x in cu for x in _EXCLUIR_DESC):
                mallas.append(c)
        if t1 and str(t1).strip() and str(t1).strip() not in t1s:
            t1s.append(str(t1).strip())
        if t2 and str(t2).strip() and str(t2).strip() not in t2s:
            t2s.append(str(t2).strip())

    vitro_cod   = t2s[0].lstrip(";").strip().split()[0] if t2s else None
    cal, sap_info = {}, {}
    estado = {"vitro_ruta": None, "ruta_err": None}
    malla_rutas = {m: None for m in mallas}

    def _q_calendario():
        try:
            cn_c = conectar_calendario()
            try:
                cur_c = cn_c.cursor()
                cur_c.execute(
                    "SELECT TOP 1 PedidoGenesis, "
                    "RIGHT('0000000000' + CAST(Lote AS VARCHAR), 10), "
                    "Logo, ZFER, CodVehiculo, Vehiculo, Version, Cliente "
                    "FROM [dbo].[TCAL_CALENDARIO_COLOMBIA_DIRECT] WHERE Orden = ?", (orden,))
                r = cur_c.fetchone()
            finally:
                cn_c.close()
            if r:
                cal.update({"lote": str(r[1] or "").strip(), "logo": str(r[2] or "").strip(),
                            "cod_veh": str(r[4] or ""), "vehiculo": str(r[5] or "").strip(),
                            "version": str(r[6] or "").strip(), "cliente": str(r[7] or "").strip()})
        except Exception:
            pass

    def _q_pieza():
        # ODATA_ZFER_CLASS_001 está en la BD SAP (misma que ODATA_ZFER_BOM)
        try:
            cn_s = conectar_sap()
            try:
                cur_s = cn_s.cursor()
                cur_s.execute(
                    "SELECT ATNAM, ATWRT FROM ODATA_ZFER_CLASS_001 "
                    "WHERE MATERIAL IN (?, ?) AND ATNAM IN ('Z_PIECE_TYPE','Z_TRAZABILITY')",
                    (zfer, zfer.zfill(18)))
                for atnam, atwrt in cur_s.fetchall():
                    sap_info[atnam] = str(atwrt or "").strip()
            finally:
                cn_s.close()
        except Exception:
            pass

    def _q_rutas():
        try:
            cn3 = db_connect()
            try:
                cur3 = cn3.cursor()
                if vitro_cod:
                    try:
                        cur3.execute("SELECT TOP 1 ruta FROM mallas.vitrojet WHERE vitro = ?", (vitro_cod,))
                        r = cur3.fetchone()
                        estado["vitro_ruta"] = str(r[0]).strip() if r and r[0] else None
                    except Exception:
                        pass
                for m in mallas:
                    m_cod = m.split()[0]
                    try:
                        cur3.execute("SELECT TOP 1 ruta_dwg FROM mallas.grandes WHERE codigo = ?", (m_cod,))
                        r = cur3.fetchone()
                        if not r or not r[0]:
                            cur3.execute("SELECT TOP 1 ruta_dwg FROM mallas.pequenas WHERE codigo = ?", (m_cod,))
                            r = cur3.fetchone()
                        malla_rutas[m] = str(r[0]).strip() if r and r[0] else None
                    except Exception:
                        malla_rutas[m] = None
            finally:
                cn3.close()
        except Exception as _e:
            estado["ruta_err"] = str(_e)

    hilos = [threading.Thread(target=f, daemon=True) for f in (_q_calendario, _q_pieza, _q_rutas)]
    for t in hilos: t.start()
    for t in hilos: t.join(timeout=12)

    vitro_ruta = estado["vitro_ruta"]
    _vd = os.path.dirname(vitro_ruta).lower().rstrip("/\\") if vitro_ruta else None

    def _misma_dir(r):
        return bool(_vd and r and os.path.dirname(r).lower().rstrip("/\\") == _vd)

    lista_mallas = [{"codigo": m, "ruta": malla_rutas.get(m),
                     "misma_carpeta": _misma_dir(malla_rutas.get(m))} for m in mallas]
    distintas = [m["codigo"] for m in lista_mallas if m["ruta"] and not m["misma_carpeta"]]
    alerta = ""
    if vitro_ruta and distintas:
        alerta = (f"RUTAS DISTINTAS — Vitro: {os.path.dirname(vitro_ruta)} | "
                  f"Malla(s): {', '.join(distintas)}")

    return {
        "ok": True, "orden": orden, "zfer": zfer,
        "vitro1": " / ".join(t1s) or "—",
        "vitro": " / ".join(v.lstrip(";") for v in t2s) or "—",
        "vitro_ruta": vitro_ruta, "ruta_err": estado["ruta_err"],
        "mallas": lista_mallas, "alerta": alerta,
        "vehiculo": cal.get("vehiculo") or "—", "version": cal.get("version") or "—",
        "lote": cal.get("lote") or "—", "logo": cal.get("logo", ""),
        "cliente": cal.get("cliente") or "—", "cod_veh": cal.get("cod_veh") or "—",
        "tipo_pieza": sap_info.get("Z_PIECE_TYPE") or "—",
        "trazabilidad": sap_info.get("Z_TRAZABILITY") or "—",
    }


# ══════════════════════════════════════════════════════════════════════════════
#  ROLES Y PERMISOS
# ══════════════════════════════════════════════════════════════════════════════
ROLES = ["admin", "dibujante", "planta"]


def listar_usuarios():
    rows = db_query("SELECT id, nombre, usuario, rol, estatus "
                    "FROM MALLAS.APP_USUARIOS ORDER BY nombre")
    return [{"id": r["id"], "nombre": r["nombre"] or "", "usuario": r["usuario"] or "",
             "rol": r["rol"] or "", "estatus": bool(r["estatus"])} for r in rows]


def guardar_usuario(uid, rol, estatus):
    rol_sql = rol if rol in ROLES else None
    cn = db_connect()
    try:
        cur = cn.cursor()
        cur.execute("UPDATE MALLAS.APP_USUARIOS SET rol=?, estatus=?, "
                    "actualizado_en=SYSDATETIME() WHERE id=?",
                    (rol_sql, 1 if estatus else 0, uid))
        cn.commit()
    finally:
        cn.close()
    return rol_sql
