# -*- coding: utf-8 -*-
"""
Sincronizador Excel → Azure SQL (mallas.*)
Estrategia: DELETE tabla + bulk INSERT con fast_executemany en lotes de 500.
Cada tabla abre y cierra su propia conexión — si Azure corta, reconecta solo.
El Excel es la fuente de verdad.
"""

import os, sys, time, datetime
import pyodbc
import openpyxl

EXCEL = (
    r"C:\Users\abotero\OneDrive - AGP GROUP"
    r"\GRP - INGENIERIA PROYECTOS 2022 - Colombia - HERRAMENTALES 2020"
    r"\LISTADO DE MALLAS Y GLASSJET 2025.xlsx"
)

CONN_STR = (
    "DRIVER={ODBC Driver 17 for SQL Server};"
    "SERVER=agpcolombia.database.windows.net,1433;"
    "DATABASE=AGP_Ingenieria;"
    "UID=DevIngenieria;"
    "PWD=HiJE068i0LQVrwA;"
    "Encrypt=yes;TrustServerCertificate=no;Connection Timeout=30;"
)

MAX_VACIAS  = 15       # filas consecutivas vacías = fin real de datos
MAX_FILAS   = 60_000  # tope absoluto anti-formato-fantasma (openpyxl read_only)
CHUNK       = 150   # filas por lote en VALUES batch (límite 2100 params SQL Server)

# ─────────────────────────────────────────────────────────────────────────────
_log_fn = None

def log(msg, tag=""):
    if _log_fn:
        _log_fn(msg, tag)
    else:
        print(f"[{datetime.datetime.now():%H:%M:%S}] {msg}", flush=True)

def clean(v, maxlen=None):
    if v is None:
        return None
    s = str(v).strip()
    if not s:
        return None
    return s[:maxlen] if maxlen and len(s) > maxlen else s

def conectar():
    for drv in ["ODBC Driver 18 for SQL Server", "ODBC Driver 17 for SQL Server"]:
        try:
            cs = CONN_STR.replace("ODBC Driver 17 for SQL Server", drv)
            c = pyodbc.connect(cs, timeout=30)
            c.autocommit = False
            return c
        except Exception:
            continue
    raise RuntimeError("No se pudo conectar — instala ODBC Driver 17/18.")

def _hoja(wb, *candidatos):
    nombres = {k.upper().strip(): k for k in wb.sheetnames}
    for c in candidatos:
        if c in wb.sheetnames:
            return wb[c]
        if c.upper().strip() in nombres:
            return wb[nombres[c.upper().strip()]]
    raise KeyError(f"No se encontró ninguna de: {candidatos}")

# ─────────────────────────────────────────────────────────────────────────────
#  Motor principal
# ─────────────────────────────────────────────────────────────────────────────

def _bulk_sync(nombre, ws, sql_delete, sql_insert, build_row_fn, pk_index=0):
    """
    Lee la hoja, deduplica por pk_index, limpia la tabla y hace bulk INSERT
    en lotes de CHUNK. Reconecta automáticamente si Azure corta la conexión.
    pk_index=None → sin dedup (glassjet_viejo).
    """
    log(f"Leyendo {nombre} del Excel...")
    t0 = time.time()
    visto  = {}   # pk → tupla
    sin_pk = []
    vacias = 0
    n_leidas = 0

    for row in ws.iter_rows(min_row=2, values_only=True):
        n_leidas += 1
        if n_leidas > MAX_FILAS:
            log(f"  WARN {nombre}: se alcanzó límite de {MAX_FILAS} filas — revisar Excel", "warn")
            break
        params = build_row_fn(row)
        if params is None:
            vacias += 1
            if vacias >= MAX_VACIAS:
                break
            continue
        vacias = 0
        if pk_index is None:
            sin_pk.append(params)
        else:
            visto[params[pk_index]] = params

    filas = list(visto.values()) if pk_index is not None else sin_pk
    dups  = sum(1 for _ in visto) - len(filas) if pk_index is not None else 0

    msg = f"  {nombre}: {len(filas)} filas leídas ({time.time()-t0:.1f}s)"
    if dups > 0:
        msg += f"  [{dups} duplicados en Excel ignorados]"
    log(msg)

    if not filas:
        log(f"  {nombre}: sin datos — omitido", "warn")
        return 0, 0

    # Extraer la parte VALUES (?,?,...) del sql_insert para construir multi-row
    # sql_insert tiene forma: "INSERT INTO tabla (c1,c2,...) VALUES (?,?,...,'literal')"
    # Necesitamos contar cuántos ? hay por fila
    n_params = sql_insert.count("?")

    def _multi_insert(cur, lote):
        """Un único INSERT con múltiples VALUE sets — rápido y sin bug de fast_executemany."""
        # Extraer la parte "INSERT INTO tabla (cols)" y la parte "VALUES (?,?,...)"
        val_start = sql_insert.upper().rfind("VALUES")
        prefix    = sql_insert[:val_start].rstrip()   # "INSERT INTO tabla (cols)"
        val_tpl   = sql_insert[val_start + 6:].strip()  # "(?,?,...,'literal')"

        # Para cada fila en el lote, la parte VALUES tiene que expandir los ?
        # Construimos: VALUES (row1),(row2),...
        rows_sql  = ",".join([val_tpl] * len(lote))
        sql_multi = f"{prefix} VALUES {rows_sql}"

        # Aplanar params: solo los ? (sin literales hardcodeados)
        params = []
        for f in lote:
            params.extend(f)
        cur.execute(sql_multi, params)

    n_lotes = -(-len(filas) // CHUNK)
    ok = err = 0
    cn = conectar()
    try:
        cn.cursor().execute(sql_delete)
        cn.commit()

        for i in range(0, len(filas), CHUNK):
            lote = filas[i:i + CHUNK]
            n    = i // CHUNK + 1
            cur  = cn.cursor()
            try:
                _multi_insert(cur, lote)
                cn.commit()
                ok += len(lote)
                log(f"  {nombre}: {n}/{n_lotes} — {ok:,} insertados", "dim")
            except pyodbc.OperationalError:
                log(f"  {nombre}: reconectando (lote {n})...", "warn")
                try: cn.close()
                except Exception: pass
                cn  = conectar()
                cur = cn.cursor()
                for f in lote:
                    try:
                        cur.execute(sql_insert, f)
                        ok += 1
                    except Exception as e:
                        err += 1
                        if err <= 3:
                            log(f"    error: {str(e)[:70]}", "warn")
                cn.commit()
            except Exception as e:
                try: cn.rollback()
                except Exception: pass
                log(f"  {nombre}: lote {n} fila a fila ({str(e)[:50]})...", "warn")
                cur2 = cn.cursor()
                for f in lote:
                    try:
                        cur2.execute(sql_insert, f)
                        ok += 1
                    except Exception as e2:
                        err += 1
                        if err <= 3:
                            log(f"    error: {str(e2)[:70]}", "warn")
                cn.commit()
    finally:
        try: cn.close()
        except Exception: pass

    log(f"  {nombre}: {ok:,} insertados  |  {err} errores",
        "ok" if err == 0 else "warn")
    return ok, err

# ─────────────────────────────────────────────────────────────────────────────
#  Motor MERGE (no borra asignados)
# ─────────────────────────────────────────────────────────────────────────────

def _bulk_sync_merge(nombre, ws, tabla, pk_col, columnas, build_row_fn):
    """
    UPSERT seguro: inserta filas nuevas y actualiza las que estén en el Excel,
    EXCEPTO las que ya tienen vehiculo/descripcion llenos en BD (ya asignadas).
    Nunca borra nada.
    """
    log(f"Leyendo {nombre} del Excel (modo MERGE)...")
    t0 = time.time()
    visto = {}
    vacias = 0
    n_leidas = 0

    for row in ws.iter_rows(min_row=2, values_only=True):
        n_leidas += 1
        if n_leidas > MAX_FILAS:
            log(f"  WARN {nombre}: se alcanzó límite de {MAX_FILAS} filas — revisar Excel", "warn")
            break
        params = build_row_fn(row)
        if params is None:
            vacias += 1
            if vacias >= MAX_VACIAS:
                break
            continue
        vacias = 0
        visto[params[0]] = params   # pk siempre en índice 0

    filas = list(visto.values())
    log(f"  {nombre}: {len(filas)} filas leídas ({time.time()-t0:.1f}s)")

    if not filas:
        log(f"  {nombre}: sin datos — omitido", "warn")
        return 0, 0

    # Obtener PKs ya asignados en BD (vehiculo/descripcion no NULL)
    cn = conectar()
    try:
        cur = cn.cursor()
        # Determinar columna de "asignado" según tabla
        if "vitrojet" in tabla:
            cur.execute(f"SELECT {pk_col} FROM {tabla} WHERE vehiculo IS NOT NULL")
        else:
            cur.execute(f"SELECT {pk_col} FROM {tabla} WHERE descripcion IS NOT NULL")
        asignados = {str(r[0]).strip() for r in cur.fetchall()}

        # Construir placeholders
        cols_str = ", ".join(columnas)
        ph_str   = ", ".join(["?"] * len(columnas))
        upd_cols = [c for c in columnas if c != pk_col]
        upd_str  = ", ".join(f"{c}=?" for c in upd_cols)

        ok = err = saltados = 0
        for f in filas:
            pk_val = str(f[0]).strip()
            if pk_val in asignados:
                saltados += 1
                continue
            try:
                # Intentar UPDATE primero; si no afecta filas → INSERT
                upd_vals = [f[i] for i, c in enumerate(columnas) if c != pk_col]
                cur.execute(
                    f"UPDATE {tabla} SET {upd_str} WHERE {pk_col}=?",
                    upd_vals + [f[0]]
                )
                if cur.rowcount == 0:
                    cur.execute(
                        f"INSERT INTO {tabla} ({cols_str}) VALUES ({ph_str})",
                        list(f)
                    )
                ok += 1
            except Exception as e:
                err += 1
                if err <= 3:
                    log(f"    error {pk_val}: {str(e)[:60]}", "warn")

        cn.commit()
        log(f"  {nombre}: {ok:,} actualizados | {saltados} preservados | {err} errores",
            "ok" if err == 0 else "warn")
        return ok, err
    finally:
        cn.close()


# ─────────────────────────────────────────────────────────────────────────────
#  Funciones por tabla
# ─────────────────────────────────────────────────────────────────────────────

def importar_vitrojet(ws):
    sql_i = (
        "INSERT INTO mallas.vitrojet "
        "(vitro,codigo_malla,tipo_malla,cod_completo,bnerig,vehiculo,version,ruta,cambio) "
        "VALUES (?,?,?,?,?,?,?,?,'excel')"
    )
    def build(r):
        vitro = clean(r[0], 30)
        if not vitro:
            return None
        malla = clean(r[1], 30) if len(r) > 1 else None
        tipo  = "G" if malla and str(malla).upper().startswith("A-") else "P"
        return (
            vitro, malla, tipo,
            clean(r[2], 100)  if len(r) > 2 else None,  # cod_completo
            clean(r[3], 20)   if len(r) > 3 else None,  # bnerig
            clean(r[4], 200)  if len(r) > 4 else None,  # vehiculo
            clean(r[5], 100)  if len(r) > 5 else None,  # version
            clean(r[6], 500)  if len(r) > 6 else None,  # ruta
            'excel',                                      # cambio — debe ir aquí para que columnas[8] mapee correctamente
        )
    # Preservar asignaciones ya existentes: no borrar, hacer MERGE por vitro
    return _bulk_sync_merge("VITROJET", ws, "mallas.vitrojet", "vitro",
                            ["vitro","codigo_malla","tipo_malla","cod_completo",
                             "bnerig","vehiculo","version","ruta","cambio"],
                            build)


def importar_grandes(ws):
    sql_i = (
        "INSERT INTO mallas.grandes "
        "(codigo,cod_veh,descripcion,pieza,tipo,version,concatenar,cambio) "
        "VALUES (?,?,?,?,?,?,?,'excel')"
    )
    def build(r):
        c = clean(r[0], 30)
        if not c:
            return None
        return (
            c,
            clean(r[1], 30)   if len(r) > 1 else None,
            clean(r[2], 200)  if len(r) > 2 else None,
            clean(r[3], 100)  if len(r) > 3 else None,
            clean(r[4], 20)   if len(r) > 4 else None,
            clean(r[5], 100)  if len(r) > 5 else None,
            clean(r[6], 300)  if len(r) > 6 else None,
        )
    return _bulk_sync("MALLAS GRANDES", ws,
                      "DELETE FROM mallas.grandes", sql_i, build)


def importar_pequenas(ws):
    sql_i = (
        "INSERT INTO mallas.pequenas "
        "(codigo,cod_veh,descripcion,pieza,tipo,version,concatenar,part_number,cambio) "
        "VALUES (?,?,?,?,?,?,?,?,'excel')"
    )
    def build(r):
        raw = clean(r[0])
        if not raw:
            return None
        try:
            codigo = int(float(raw))
        except Exception:
            return None
        # Requiere descripción para filtrar ghost rows con solo número en col A
        desc = clean(r[2], 200) if len(r) > 2 else None
        if not desc:
            return None
        return (
            codigo,
            clean(r[1], 30) if len(r) > 1 else None,
            desc,
            clean(r[3], 100) if len(r) > 3 else None,
            clean(r[4], 20)  if len(r) > 4 else None,
            clean(r[5], 100) if len(r) > 5 else None,
            clean(r[6], 300) if len(r) > 6 else None,
            clean(r[7], 100) if len(r) > 7 else None,
        )
    return _bulk_sync("MALLAS PEQUEÑAS", ws,
                      "DELETE FROM mallas.pequenas", sql_i, build)


def importar_pasta_plata(ws):
    sql_i = (
        "INSERT INTO mallas.pasta_plata "
        "(consecutivo,tipo,vehiculo,cod_vehiculo,version,pieza,ruta_archivo,caso,cambio) "
        "VALUES (?,?,?,?,?,?,?,?,'excel')"
    )
    def build(r):
        c = clean(r[0], 30)
        if not c:
            return None
        return (
            c,
            clean(r[1], 20),
            clean(r[2], 200),
            clean(str(r[3]), 30)  if r[3] is not None else None,
            clean(str(r[4]), 100) if r[4] is not None else None,
            clean(str(r[5]), 100) if r[5] is not None else None,
            clean(r[6], 500)      if len(r) > 6 else None,
            clean(r[7], 200)      if len(r) > 7 else None,
        )
    return _bulk_sync("PASTA DE PLATA", ws,
                      "DELETE FROM mallas.pasta_plata", sql_i, build)


def importar_vinilos(ws):
    sql_i = (
        "INSERT INTO mallas.vinilos "
        "(herramental,vehiculo,cod_vehiculo,version,pieza,tipo,cambio) "
        "VALUES (?,?,?,?,?,?,'excel')"
    )
    def build(r):
        h = clean(r[0], 30)
        if not h:
            return None
        tipo = clean(r[5], 20) if len(r) > 5 else None
        if tipo == "BN2":
            tipo = "BN"
        return (
            h,
            clean(r[1], 200),
            clean(str(r[2]), 30)  if r[2] is not None else None,
            clean(str(r[3]), 100) if r[3] is not None else None,
            clean(str(r[4]), 100) if r[4] is not None else None,
            tipo,
        )
    return _bulk_sync("VINILOS", ws,
                      "DELETE FROM mallas.vinilos", sql_i, build)


def importar_glassjet_viejo(ws):
    sql_i = (
        "INSERT INTO mallas.glassjet_viejo "
        "(malla,glassjet,part_number,tipo,vehiculo,homologacion_vitro) "
        "VALUES (?,?,?,?,?,?)"
    )
    def build(r):
        if not any(v is not None for v in r[:6]):
            return None
        return (
            clean(str(r[0]), 50) if r[0] is not None else None,
            clean(str(r[1]), 50) if r[1] is not None else None,
            clean(r[2], 100),
            clean(r[3], 20),
            clean(r[4], 200),
            clean(r[5], 50),
        )
    return _bulk_sync("GLASSJET VIEJO", ws,
                      "DELETE FROM mallas.glassjet_viejo", sql_i, build,
                      pk_index=None)

# ─────────────────────────────────────────────────────────────────────────────
#  MAIN
# ─────────────────────────────────────────────────────────────────────────────

def main(log_fn=None):
    global _log_fn
    _log_fn = log_fn

    log("=== SINCRONIZADOR AGP Excel -> Azure SQL ===")

    if not os.path.isfile(EXCEL):
        log(f"ERROR: Excel no encontrado:\n  {EXCEL}", "err")
        return False

    log("Abriendo Excel...")
    t0 = time.time()
    try:
        wb = openpyxl.load_workbook(EXCEL, read_only=True, data_only=True)
    except Exception as e:
        log(f"ERROR al abrir Excel: {e}", "err")
        return False
    log(f"Excel cargado en {time.time()-t0:.1f}s")

    log("Verificando conexión Azure SQL...")
    try:
        _t = conectar(); _t.close()
    except Exception as e:
        log(str(e), "err")
        return False
    log("Conexión OK\n")

    errores_total = 0
    try:
        _, e = importar_vitrojet(   _hoja(wb, "VITROJET"));               errores_total += e
        _, e = importar_grandes(    _hoja(wb, "GRANDES"));                 errores_total += e
        _, e = importar_pequenas(   _hoja(wb, "PEQUEÑAS", "PEQUENAS"));    errores_total += e
        _, e = importar_pasta_plata(_hoja(wb, "PASTA DE PLATA"));          errores_total += e
        _, e = importar_glassjet_viejo(
            _hoja(wb, "GLASSJET VIEJO NO ALIMENTAR INV", "GLASSJET VIEJO")); errores_total += e
        _, e = importar_vinilos(    _hoja(wb, "VINILOS"));                 errores_total += e

        log(f"\n=== Completado en {time.time()-t0:.0f}s ===",
            "ok" if errores_total == 0 else "warn")

        # Resincronizar secuencias para que el consecutivo siga desde el MAX real
        try:
            from asignaciones import sincronizar_secuencias
            sincronizar_secuencias()
            log("  Secuencias de consecutivo actualizadas", "ok")
        except Exception as _se:
            log(f"  WARN secuencias: {_se}", "warn")

        cn = conectar()
        cur = cn.cursor()
        for tabla in ["mallas.vitrojet", "mallas.grandes", "mallas.pequenas",
                      "mallas.pasta_plata", "mallas.glassjet_viejo", "mallas.vinilos"]:
            cur.execute(f"SELECT COUNT(*) FROM {tabla}")
            log(f"  {tabla}: {cur.fetchone()[0]:,} registros", "dim")
        cn.close()

    except Exception as e:
        log(f"ERROR FATAL: {e}", "err")
        import traceback; traceback.print_exc()
        return False
    finally:
        wb.close()

    return errores_total == 0


if __name__ == "__main__":
    sys.exit(0 if main() else 1)
