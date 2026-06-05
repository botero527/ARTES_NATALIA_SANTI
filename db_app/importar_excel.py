# -*- coding: utf-8 -*-
"""
Importador Excel → SQL Server (Vitros_Mallas)
Ejecutar una sola vez para migrar todos los datos históricos.
Re-ejecutable: IF NOT EXISTS evita duplicados (excepto glassjet_viejo que se limpia antes).
"""
import os, sys, time, datetime
import pyodbc
import openpyxl

# Ruta sincronizada desde SharePoint via OneDrive — siempre actualizada
EXCEL = r"C:\Users\abotero\OneDrive - AGP GROUP\GRP - INGENIERIA PROYECTOS 2022 - Colombia - HERRAMENTALES 2020\LISTADO DE MALLAS Y GLASSJET 2025.xlsx"

CONN_STR = (
    "DRIVER={ODBC Driver 17 for SQL Server};"
    "SERVER=.\\SQLEXPRESS;"
    "DATABASE=Vitros_Mallas;"
    "Trusted_Connection=yes;"
)

BATCH = 500

def ts():
    return datetime.datetime.now().strftime("%H:%M:%S")

def log(msg):
    print(f"[{ts()}] {msg}")

def clean(v):
    if v is None:
        return None
    s = str(v).strip()
    return s if s else None

def conectar():
    for driver in ["ODBC Driver 17 for SQL Server", "ODBC Driver 18 for SQL Server",
                   "SQL Server"]:
        try:
            cs = CONN_STR.replace("ODBC Driver 17 for SQL Server", driver)
            return pyodbc.connect(cs, timeout=10)
        except Exception:
            continue
    log("ERROR: no se pudo conectar. Instala ODBC Driver 17/18 for SQL Server.")
    sys.exit(1)

def ejecutar_batch(cursor, sql, filas):
    """Bulk insert con fallback fila-a-fila si hay errores de truncación."""
    if not filas:
        return 0, 0
    ok = err = 0
    cursor.fast_executemany = False
    try:
        cursor.executemany(sql, filas)
        ok = len(filas)
    except Exception:
        for fila in filas:
            try:
                cursor.execute(sql, fila)
                ok += 1
            except Exception:
                err += 1
    return ok, err

def _importar(nombre, ws, cursor, sql, build_row_fn):
    log(f"Importando {nombre}...")
    filas = []; ok = err = 0
    for row in ws.iter_rows(min_row=2, values_only=True):
        fila = build_row_fn(row)
        if fila is None:
            continue
        filas.append(fila)
        if len(filas) >= BATCH:
            b_ok, b_err = ejecutar_batch(cursor, sql, filas)
            ok += b_ok; err += b_err; filas = []
    if filas:
        b_ok, b_err = ejecutar_batch(cursor, sql, filas)
        ok += b_ok; err += b_err
    log(f"  {nombre}: {ok} insertados, {err} errores")
    return ok, err

# ── Importadores ──────────────────────────────────────────────────────────────

def importar_grandes(ws, cursor):
    sql = """IF NOT EXISTS (SELECT 1 FROM mallas_grandes WHERE codigo=?)
             INSERT INTO mallas_grandes (codigo,cod_veh,descripcion,pieza,tipo,version,concatenar,cambio)
             VALUES (?,?,?,?,?,?,?,'importado_excel')"""
    def row(r):
        c = clean(r[0])
        return None if not c else (c, c, clean(r[1]), clean(r[2]),
               clean(r[3]), clean(r[4]), clean(r[5]), clean(r[6]))
    return _importar("MALLAS GRANDES", ws, cursor, sql, row)

def importar_pequenas(ws, cursor):
    sql = """IF NOT EXISTS (SELECT 1 FROM mallas_pequenas WHERE codigo=?)
             INSERT INTO mallas_pequenas (codigo,cod_veh,descripcion,pieza,tipo,version,concatenar,part_number,cambio)
             VALUES (?,?,?,?,?,?,?,?,'importado_excel')"""
    def row(r):
        raw = clean(r[0])
        if not raw:
            return None
        try:
            codigo = int(float(raw))
        except Exception:
            return None
        pn = clean(r[7]) if len(r) > 7 else None
        return (codigo, codigo, clean(r[1]), clean(r[2]),
                clean(r[3]), clean(r[4]), clean(r[5]), clean(r[6]), pn)
    return _importar("MALLAS PEQUEÑAS", ws, cursor, sql, row)

def importar_vitrojet(ws, cursor):
    sql = """IF NOT EXISTS (SELECT 1 FROM vitrojet WHERE vitro=?)
             INSERT INTO vitrojet (vitro,codigo_malla,tipo_malla,cod_completo,bnerig,vehiculo,version,cambio)
             VALUES (?,?,?,?,?,?,?,'importado_excel')"""
    def row(r):
        vitro = clean(r[0])
        malla = clean(r[1])
        if not vitro or not malla:
            return None
        tipo = 'G' if str(malla).startswith('A-') else 'P'
        ver  = clean(r[5]) if len(r) > 5 else None
        return (vitro, vitro, str(malla), tipo, clean(r[2]), clean(r[3]), clean(r[4]), ver)
    return _importar("VITROJET", ws, cursor, sql, row)

def importar_pasta_plata(ws, cursor):
    sql = """IF NOT EXISTS (SELECT 1 FROM pasta_plata WHERE consecutivo=?)
             INSERT INTO pasta_plata (consecutivo,tipo,vehiculo,cod_vehiculo,version,pieza,ruta_archivo,caso,cambio)
             VALUES (?,?,?,?,?,?,?,?,'importado_excel')"""
    def row(r):
        c = clean(r[0])
        if not c:
            return None
        caso = clean(r[7]) if len(r) > 7 else None
        return (c, c, clean(r[1]), clean(r[2]),
                clean(str(r[3]) if r[3] else None),
                clean(r[4]), clean(str(r[5]) if r[5] else None),
                clean(r[6]), caso)
    return _importar("PASTA DE PLATA", ws, cursor, sql, row)

def importar_glassjet_viejo(ws, cursor):
    # Tabla histórica sin PK propia — limpiar antes de importar para evitar duplicados
    cursor.execute("DELETE FROM glassjet_viejo")
    log("  glassjet_viejo limpiada (re-importación limpia)")
    sql = """INSERT INTO glassjet_viejo (malla,glassjet,part_number,tipo,vehiculo,homologacion_vitro)
             VALUES (?,?,?,?,?,?)"""
    def row(r):
        if not any(v is not None for v in r[:6]):
            return None
        return (clean(str(r[0])) if r[0] else None,
                clean(str(r[1])) if r[1] else None,
                clean(r[2]), clean(r[3]), clean(r[4]), clean(r[5]))
    return _importar("GLASSJET VIEJO", ws, cursor, sql, row)

def importar_vinilos(ws, cursor):
    sql = """IF NOT EXISTS (SELECT 1 FROM vinilos WHERE herramental=?)
             INSERT INTO vinilos (herramental,vehiculo,cod_vehiculo,version,pieza,tipo,cambio)
             VALUES (?,?,?,?,?,?,'importado_excel')"""
    def row(r):
        h = clean(r[0])
        if not h:
            return None
        tipo = clean(r[5])
        if tipo == 'BN2':
            tipo = 'BN'
        return (h, h, clean(r[1]),
                clean(str(r[2]) if r[2] else None),
                clean(str(r[3]) if r[3] else None),
                clean(str(r[4]) if r[4] else None), tipo)
    return _importar("VINILOS", ws, cursor, sql, row)

# ── MAIN ──────────────────────────────────────────────────────────────────────

def main():
    log("=== IMPORTADOR AGP Excel → SQL Server ===")
    if not os.path.isfile(EXCEL):
        log(f"ERROR: no se encontró {EXCEL}"); sys.exit(1)

    log("Abriendo Excel...")
    t0 = time.time()
    wb = openpyxl.load_workbook(EXCEL, read_only=True, data_only=True)
    log(f"Excel cargado en {time.time()-t0:.1f}s")

    log("Conectando a SQL Server...")
    conn = conectar()
    conn.autocommit = False
    cur = conn.cursor()
    log("Conexión OK")

    try:
        importar_grandes(wb['GRANDES'], cur);          conn.commit()
        importar_pequenas(wb['PEQUEÑAS'], cur);        conn.commit()
        importar_vitrojet(wb['VITROJET'], cur);        conn.commit()
        importar_pasta_plata(wb['PASTA DE PLATA'], cur); conn.commit()
        importar_glassjet_viejo(wb['GLASSJET VIEJO NO ALIMENTAR INV'], cur); conn.commit()
        importar_vinilos(wb['VINILOS'], cur);          conn.commit()

        log(f"\n=== Completado en {time.time()-t0:.0f}s ===")
        for t in ['mallas_grandes','mallas_pequenas','vitrojet',
                  'pasta_plata','glassjet_viejo','vinilos']:
            cur.execute(f"SELECT COUNT(*) FROM {t}")
            log(f"  {t}: {cur.fetchone()[0]:,} registros")
    except Exception as e:
        conn.rollback()
        log(f"ERROR FATAL: {e}")
        import traceback; traceback.print_exc()
    finally:
        cur.close(); conn.close()

if __name__ == "__main__":
    main()
