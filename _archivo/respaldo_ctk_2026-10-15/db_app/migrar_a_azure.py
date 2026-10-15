# -*- coding: utf-8 -*-
"""
Migracion LOCAL (.\\SQLEXPRESS / Vitros_Mallas) -> AZURE (agpcolombia.database.windows.net / AGP_Ingenieria)
Lee de la BD local tabla a tabla y escribe en Azure con esquema mallas.*
Ejecutar UNA SOLA VEZ despues de correr azure_crear_tablas.sql
"""
import sys, time, datetime
import pyodbc

CONN_LOCAL = (
    "DRIVER={ODBC Driver 17 for SQL Server};"
    "SERVER=.\\SQLEXPRESS;"
    "DATABASE=Vitros_Mallas;"
    "Trusted_Connection=yes;"
)

CONN_AZURE = (
    "DRIVER={ODBC Driver 17 for SQL Server};"
    "SERVER=agpcolombia.database.windows.net,1433;"
    "DATABASE=AGP_Ingenieria;"
    "UID=DevIngenieria;"
    "PWD=HiJE068i0LQVrwA;"
    "Encrypt=yes;"
    "TrustServerCertificate=no;"
    "Connection Timeout=30;"
)

BATCH = 200

def ts():
    return datetime.datetime.now().strftime("%H:%M:%S")

def log(msg):
    print(f"[{ts()}] {msg}", flush=True)

def conectar(cs, nombre):
    for driver in ["ODBC Driver 18 for SQL Server", "ODBC Driver 17 for SQL Server"]:
        try:
            return pyodbc.connect(cs.replace("ODBC Driver 17 for SQL Server", driver), timeout=30)
        except Exception:
            continue
    log(f"ERROR: no se pudo conectar a {nombre}")
    sys.exit(1)

def migrar_tabla(src_cur, dst_conn, src_sql, dst_sql, nombre, limpiar_sql=None):
    log(f"Migrando {nombre}...")
    src_cur.execute(src_sql)
    rows = src_cur.fetchall()
    log(f"  Leidos {len(rows):,} registros de local")

    dst_cur = dst_conn.cursor()
    dst_cur.fast_executemany = False

    if limpiar_sql:
        dst_cur.execute(limpiar_sql)
        dst_conn.commit()
        log(f"  {nombre} limpiada en Azure")

    ok = err = skip = 0
    first_error = None
    for i, row in enumerate(rows):
        try:
            dst_cur.execute(dst_sql, tuple(row))
            ok += 1
        except pyodbc.IntegrityError:
            skip += 1  # duplicado — ya existe, ignorar
        except Exception as e:
            err += 1
            if first_error is None:
                first_error = str(e)
        if (i + 1) % BATCH == 0:
            dst_conn.commit()
    dst_conn.commit()

    resumen = f"  {nombre}: {ok:,} insertados, {skip} duplicados ignorados, {err} errores"
    if first_error:
        resumen += f"\n  Primer error: {first_error}"
    log(resumen)
    return ok, err

def main():
    log("=== MIGRACION LOCAL -> AZURE ===")
    t0 = time.time()

    log("Conectando a SQL Server local...")
    src = conectar(CONN_LOCAL, "LOCAL")
    src.autocommit = True
    src_cur = src.cursor()
    log("Conectando a Azure SQL...")
    dst = conectar(CONN_AZURE, "AZURE")
    log("Conexiones OK\n")

    # ── mallas.grandes ──────────────────────────────────────────
    migrar_tabla(
        src_cur, dst,
        "SELECT codigo,cod_veh,descripcion,pieza,tipo,version,concatenar,cambio FROM mallas_grandes",
        "INSERT INTO mallas.grandes (codigo,cod_veh,descripcion,pieza,tipo,version,concatenar,cambio) VALUES (?,?,?,?,?,?,?,?)",
        "mallas.grandes"
    )

    # ── mallas.pequenas (solo filas reales, sin vacias) ─────────
    migrar_tabla(
        src_cur, dst,
        "SELECT codigo,cod_veh,descripcion,pieza,tipo,version,concatenar,part_number,cambio FROM mallas_pequenas WHERE descripcion IS NOT NULL OR cod_veh IS NOT NULL",
        "INSERT INTO mallas.pequenas (codigo,cod_veh,descripcion,pieza,tipo,version,concatenar,part_number,cambio) VALUES (?,?,?,?,?,?,?,?,?)",
        "mallas.pequenas"
    )

    # ── mallas.vitrojet ─────────────────────────────────────────
    migrar_tabla(
        src_cur, dst,
        "SELECT vitro,codigo_malla,tipo_malla,cod_completo,bnerig,vehiculo,version,ruta,cambio FROM vitrojet",
        "INSERT INTO mallas.vitrojet (vitro,codigo_malla,tipo_malla,cod_completo,bnerig,vehiculo,version,ruta,cambio) VALUES (?,?,?,?,?,?,?,?,?)",
        "mallas.vitrojet"
    )

    # ── mallas.pasta_plata ──────────────────────────────────────
    migrar_tabla(
        src_cur, dst,
        "SELECT consecutivo,tipo,vehiculo,cod_vehiculo,version,pieza,ruta_archivo,caso,cambio FROM pasta_plata",
        "INSERT INTO mallas.pasta_plata (consecutivo,tipo,vehiculo,cod_vehiculo,version,pieza,ruta_archivo,caso,cambio) VALUES (?,?,?,?,?,?,?,?,?)",
        "mallas.pasta_plata"
    )

    # ── mallas.glassjet_viejo ───────────────────────────────────
    migrar_tabla(
        src_cur, dst,
        "SELECT malla,glassjet,part_number,tipo,vehiculo,homologacion_vitro FROM glassjet_viejo",
        "INSERT INTO mallas.glassjet_viejo (malla,glassjet,part_number,tipo,vehiculo,homologacion_vitro) VALUES (?,?,?,?,?,?)",
        "mallas.glassjet_viejo",
        limpiar_sql="DELETE FROM mallas.glassjet_viejo"
    )

    # ── mallas.vinilos ──────────────────────────────────────────
    migrar_tabla(
        src_cur, dst,
        "SELECT herramental,vehiculo,cod_vehiculo,version,pieza,tipo,ruta,cambio FROM vinilos",
        "INSERT INTO mallas.vinilos (herramental,vehiculo,cod_vehiculo,version,pieza,tipo,ruta,cambio) VALUES (?,?,?,?,?,?,?,?)",
        "mallas.vinilos"
    )

    # ── Resumen ─────────────────────────────────────────────────
    log(f"\n=== Completado en {time.time()-t0:.0f}s ===")
    dst_cur2 = dst.cursor()
    for t in ['mallas.grandes','mallas.pequenas','mallas.vitrojet',
              'mallas.pasta_plata','mallas.glassjet_viejo','mallas.vinilos']:
        dst_cur2.execute(f"SELECT COUNT(*) FROM {t}")
        log(f"  {t}: {dst_cur2.fetchone()[0]:,} registros en Azure")
    dst_cur2.close()

    src.close(); dst.close()

if __name__ == "__main__":
    main()
