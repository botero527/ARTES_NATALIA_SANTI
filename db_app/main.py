# -*- coding: utf-8 -*-
"""AGP Glass DB — Backend FastAPI"""
from fastapi import FastAPI, Query, HTTPException
from fastapi.responses import HTMLResponse, JSONResponse
from fastapi.staticfiles import StaticFiles
import pyodbc, os, pathlib

app = FastAPI(title="AGP Glass DB")

CONN_STR = (
    "DRIVER={ODBC Driver 17 for SQL Server};"
    "SERVER=.\\SQLEXPRESS;"
    "DATABASE=Vitros_Mallas;"
    "Trusted_Connection=yes;"
)

def get_conn():
    return pyodbc.connect(CONN_STR, timeout=10)

def rows_to_dicts(cursor):
    cols = [c[0] for c in cursor.description]
    return [dict(zip(cols, row)) for row in cursor.fetchall()]

# ── Stats ──────────────────────────────────────────────────────────────────────
@app.get("/api/stats")
def stats():
    conn = get_conn(); cur = conn.cursor()
    result = {}
    for t in ["mallas_grandes","mallas_pequenas","vitrojet","pasta_plata","glassjet_viejo","vinilos"]:
        cur.execute(f"SELECT COUNT(*) FROM {t}")
        result[t] = cur.fetchone()[0]
    conn.close()
    return result

# ── Mallas Grandes ─────────────────────────────────────────────────────────────
@app.get("/api/mallas-grandes")
def buscar_grandes(q: str = Query(""), limit: int = 50):
    conn = get_conn(); cur = conn.cursor()
    if q:
        like = f"%{q}%"
        cur.execute("""
            SELECT TOP(?) codigo,cod_veh,descripcion,pieza,tipo,version,concatenar
            FROM mallas_grandes
            WHERE descripcion LIKE ? OR codigo LIKE ? OR cod_veh LIKE ? OR concatenar LIKE ?
            ORDER BY codigo
        """, limit, like, like, like, like)
    else:
        cur.execute("SELECT TOP(?) codigo,cod_veh,descripcion,pieza,tipo,version,concatenar FROM mallas_grandes ORDER BY codigo", limit)
    data = rows_to_dicts(cur); conn.close()
    return data

# ── Mallas Pequeñas ────────────────────────────────────────────────────────────
@app.get("/api/mallas-pequenas")
def buscar_pequenas(q: str = Query(""), limit: int = 50):
    conn = get_conn(); cur = conn.cursor()
    if q:
        like = f"%{q}%"
        cur.execute("""
            SELECT TOP(?) codigo,cod_veh,descripcion,pieza,tipo,version,concatenar
            FROM mallas_pequenas
            WHERE descripcion LIKE ? OR CAST(codigo AS NVARCHAR) LIKE ? OR cod_veh LIKE ?
            ORDER BY codigo
        """, limit, like, like, like)
    else:
        cur.execute("SELECT TOP(?) codigo,cod_veh,descripcion,pieza,tipo,version,concatenar FROM mallas_pequenas ORDER BY codigo", limit)
    data = rows_to_dicts(cur); conn.close()
    return data

# ── Vitrojet ───────────────────────────────────────────────────────────────────
@app.get("/api/vitrojet")
def buscar_vitrojet(q: str = Query(""), limit: int = 50):
    conn = get_conn(); cur = conn.cursor()
    if q:
        like = f"%{q}%"
        cur.execute("""
            SELECT TOP(?) v.vitro, v.codigo_malla, v.tipo_malla, v.bnerig, v.vehiculo, v.version,
                   COALESCE(g.concatenar, CAST(p.codigo AS NVARCHAR)+' - '+p.descripcion,'') AS info_malla
            FROM vitrojet v
            LEFT JOIN mallas_grandes g ON v.tipo_malla='G' AND v.codigo_malla=g.codigo
            LEFT JOIN mallas_pequenas p ON v.tipo_malla='P' AND v.codigo_malla=CAST(p.codigo AS NVARCHAR)
            WHERE v.vitro LIKE ? OR v.vehiculo LIKE ? OR v.codigo_malla LIKE ?
            ORDER BY v.vitro DESC
        """, limit, like, like, like)
    else:
        cur.execute("""
            SELECT TOP(?) v.vitro, v.codigo_malla, v.tipo_malla, v.bnerig, v.vehiculo, v.version,
                   COALESCE(g.concatenar, CAST(p.codigo AS NVARCHAR)+' - '+p.descripcion,'') AS info_malla
            FROM vitrojet v
            LEFT JOIN mallas_grandes g ON v.tipo_malla='G' AND v.codigo_malla=g.codigo
            LEFT JOIN mallas_pequenas p ON v.tipo_malla='P' AND v.codigo_malla=CAST(p.codigo AS NVARCHAR)
            ORDER BY v.vitro DESC
        """, limit)
    data = rows_to_dicts(cur); conn.close()
    return data

# ── Vinilos ────────────────────────────────────────────────────────────────────
@app.get("/api/vinilos")
def buscar_vinilos(q: str = Query(""), limit: int = 50):
    conn = get_conn(); cur = conn.cursor()
    if q:
        like = f"%{q}%"
        cur.execute("""
            SELECT TOP(?) herramental,vehiculo,cod_vehiculo,version,pieza,tipo
            FROM vinilos WHERE vehiculo LIKE ? OR herramental LIKE ? ORDER BY herramental DESC
        """, limit, like, like)
    else:
        cur.execute("SELECT TOP(?) herramental,vehiculo,cod_vehiculo,version,pieza,tipo FROM vinilos ORDER BY herramental DESC", limit)
    data = rows_to_dicts(cur); conn.close()
    return data

# ── Pasta Plata ────────────────────────────────────────────────────────────────
@app.get("/api/pasta-plata")
def buscar_pasta(q: str = Query(""), limit: int = 50):
    conn = get_conn(); cur = conn.cursor()
    if q:
        like = f"%{q}%"
        cur.execute("""
            SELECT TOP(?) consecutivo,tipo,vehiculo,cod_vehiculo,version,pieza,caso
            FROM pasta_plata WHERE vehiculo LIKE ? OR consecutivo LIKE ? ORDER BY consecutivo DESC
        """, limit, like, like)
    else:
        cur.execute("SELECT TOP(?) consecutivo,tipo,vehiculo,cod_vehiculo,version,pieza,caso FROM pasta_plata ORDER BY consecutivo DESC", limit)
    data = rows_to_dicts(cur); conn.close()
    return data

# ── HTML Principal ─────────────────────────────────────────────────────────────
@app.get("/", response_class=HTMLResponse)
def index():
    html_path = pathlib.Path(__file__).parent / "static" / "index.html"
    return HTMLResponse(html_path.read_text(encoding="utf-8"))
