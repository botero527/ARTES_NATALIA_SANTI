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

#   ── Mallas Pequeñas ────────────────────────────────────────────────────────────
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

# ── Sincronizar desde Excel SharePoint ────────────────────────────────────────
EXCEL_PATH = r"C:\Users\abotero\OneDrive - AGP GROUP\GRP - INGENIERIA PROYECTOS 2022 - Colombia - HERRAMENTALES 2020\LISTADO DE MALLAS Y GLASSJET 2025.xlsx"

import subprocess, sys, threading
_sync_state = {"running": False, "log": [], "ok": None}

def _run_sync():
    global _sync_state
    script = pathlib.Path(__file__).parent / "importar_excel.py"
    env = os.environ.copy()
    env["PYTHONUTF8"] = "1"
    _sync_state = {"running": True, "log": ["Iniciando importación..."], "ok": None}
    try:
        proc = subprocess.Popen(
            [sys.executable, "-X", "utf8", str(script)],
            stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
            stdin=subprocess.DEVNULL,
            text=True, encoding="utf-8", errors="replace", env=env
        )
        for line in proc.stdout:
            _sync_state["log"].append(line.rstrip())
        proc.wait()
        _sync_state["ok"] = proc.returncode == 0
    except Exception as e:
        _sync_state["log"].append(f"ERROR: {e}")
        _sync_state["ok"] = False
    finally:
        _sync_state["running"] = False

@app.post("/api/sync")
def sync_excel():
    if _sync_state["running"]:
        return {"started": False, "msg": "Ya hay una sincronización en curso"}
    if not pathlib.Path(EXCEL_PATH).exists():
        raise HTTPException(404, f"Excel no encontrado: {EXCEL_PATH}")
    t = threading.Thread(target=_run_sync, daemon=True)
    t.start()
    return {"started": True}

@app.get("/api/sync-status")
def sync_status():
    return {
        "running": _sync_state["running"],
        "ok": _sync_state["ok"],
        "log": _sync_state["log"][-30:]
    }

@app.get("/api/excel-status")
def excel_status():
    p = pathlib.Path(EXCEL_PATH)
    if not p.exists():
        return {"existe": False, "ruta": EXCEL_PATH}
    import datetime
    mtime = datetime.datetime.fromtimestamp(p.stat().st_mtime)
    return {"existe": True, "ruta": EXCEL_PATH,
            "modificado": mtime.strftime("%d/%m/%Y %H:%M")}

# ── HTML Principal ─────────────────────────────────────────────────────────────
@app.get("/", response_class=HTMLResponse)
def index():
    html_path = pathlib.Path(__file__).parent / "static" / "index.html"
    return HTMLResponse(html_path.read_text(encoding="utf-8"))
