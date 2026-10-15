# -*- coding: utf-8 -*-
"""
BD de la pestaña "Logos Laser" — tabla mallas.LOGOS_LASER.

Mismo patrón que ya usa `agp_app.py` para pasta_plata/vinilos (ver
`_form_insertar` en agp_app.py): el ID consecutivo "ML-0001" se genera con un
INSERT...SELECT...WITH (UPDLOCK, HOLDLOCK) — Azure SQL bloquea la lectura del
MAX hasta que el INSERT anterior confirma, así que dos personas creando una
fila al mismo tiempo desde PCs distintos nunca pueden sacar el mismo número.
No hace falta una SEQUENCE aparte para esto.

Este módulo es independiente de agp_app.py — solo depende de `importar_excel`
(la conexión compartida de db_app), igual que asignaciones.py.
"""

import os, sys

_DIR = os.path.dirname(__file__)
sys.path.insert(0, _DIR)
from importar_excel import conectar

TABLA = "mallas.LOGOS_LASER"

# Columnas editables desde la grilla (todo menos id / created_at / updated_at,
# que los controla la BD). Cualquier columna fuera de esta lista se rechaza
# antes de armar el SQL — evita que una futura tecla mal puesta en la grilla
# termine actualizando una columna que no debería.
_COLUMNAS_EDITABLES = {
    "cliente", "producto", "dxf", "sjf", "validado",
    "cargado_genesis", "azure", "app", "observaciones", "responsable",
}

_COLUMNAS_SELECT = (
    "id, cliente, producto, dxf, sjf, validado, cargado_genesis, azure, app, "
    "observaciones, responsable, created_at, updated_at"
)


def listar(filtro_texto=""):
    """
    Devuelve todas las filas (list[dict]), la más reciente primero.
    filtro_texto: si viene, filtra por cliente/producto/id (LIKE %texto%).
    Con el volumen esperado de esta tabla no hace falta paginar.
    """
    cn = conectar()
    try:
        cur = cn.cursor()
        if filtro_texto:
            like = f"%{filtro_texto}%"
            cur.execute(
                f"SELECT {_COLUMNAS_SELECT} FROM {TABLA} "
                f"WHERE cliente LIKE ? OR producto LIKE ? OR id LIKE ? "
                f"ORDER BY TRY_CAST(SUBSTRING(id,4,50) AS INT) DESC",
                (like, like, like))
        else:
            cur.execute(
                f"SELECT {_COLUMNAS_SELECT} FROM {TABLA} "
                f"ORDER BY TRY_CAST(SUBSTRING(id,4,50) AS INT) DESC")
        cols = [c[0] for c in cur.description]
        return [dict(zip(cols, r)) for r in cur.fetchall()]
    finally:
        cn.close()


def crear_fila(responsable=None):
    """
    Inserta una fila nueva con ID autogenerado 'ML-0001', 'ML-0002', ...
    (4 dígitos, sigue creciendo sin romperse pasado 9999) y la devuelve
    completa (dict) para pintarla de una en la grilla.
    Multi-PC safe: ver nota del módulo arriba.
    """
    cn = conectar()
    try:
        cur = cn.cursor()
        cur.execute(
            f"INSERT INTO {TABLA} (id, responsable) "
            f"SELECT 'ML-' + RIGHT('0000' + CAST("
            f"ISNULL(MAX(TRY_CAST(SUBSTRING(id,4,50) AS INT)),0)+1 AS VARCHAR(10)), 4), ? "
            f"FROM {TABLA} WITH (UPDLOCK, HOLDLOCK)",
            (responsable,))
        cur.execute(
            f"SELECT TOP 1 {_COLUMNAS_SELECT} FROM {TABLA} "
            f"ORDER BY TRY_CAST(SUBSTRING(id,4,50) AS INT) DESC")
        cols = [c[0] for c in cur.description]
        fila = dict(zip(cols, cur.fetchone()))
        cn.commit()
        return fila
    except Exception:
        try: cn.rollback()
        except Exception: pass
        raise
    finally:
        cn.close()


def actualizar_campo(id_fila, campo, valor):
    """
    Guarda un solo campo de una fila (lo que llama la grilla al confirmar
    una celda editada). `campo` se valida contra la lista blanca de arriba.
    """
    if campo not in _COLUMNAS_EDITABLES:
        raise ValueError(f"Columna no editable: {campo}")
    if isinstance(valor, str):
        valor = valor.strip() or None
    cn = conectar()
    try:
        cur = cn.cursor()
        cur.execute(
            f"UPDATE {TABLA} SET {campo}=?, updated_at=GETDATE() WHERE id=?",
            (valor, id_fila))
        cn.commit()
        return cur.rowcount > 0
    except Exception:
        try: cn.rollback()
        except Exception: pass
        raise
    finally:
        cn.close()
