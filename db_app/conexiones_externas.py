# -*- coding: utf-8 -*-
"""
Conexiones a las bases externas que usa el Scanner de órdenes.
Copiadas tal cual de agp_app.py (_conectar_comercial / _conectar_sap /
_conectar_calendario) para que la UI nueva no tenga que importar la app
CustomTkinter. Mismo wrapper que importar_excel: acepta placeholders '?'.

Pendiente recomendado: mover estas credenciales a un .env (como ya hacen
Tickets_Planta y Control_Inventarios) en vez de tenerlas en el código.
"""

try:
    import pymssql as _pymssql
except ImportError:
    _pymssql = None

from db_app.importar_excel import _ConnWrap


def _conectar(nombre, **kw):
    if _pymssql is None:
        raise RuntimeError("pymssql no disponible")
    try:
        return _ConnWrap(_pymssql.connect(tds_version="7.3", **kw))
    except Exception as e:
        raise RuntimeError(f"No se pudo conectar a {nombre}\n{e}")


def conectar_comercial():
    return _conectar("Comercial (192.168.2.23)",
                     server="192.168.2.23", user="Consulta", password="@GPgl4$$2021",
                     database="Comercial", timeout=10, login_timeout=10)


def conectar_sap():
    return _conectar("SAP Azure",
                     server="agpcolsap.database.windows.net", port=1433,
                     user="Viewer", password="AgpconsCol2023", database="DB_COL_SAP",
                     timeout=15, login_timeout=15, charset="UTF-8")


def conectar_calendario():
    return _conectar("Calendario",
                     server="agpcolcalendario.database.windows.net", port=1433,
                     user="Consulta", password="@GPgl4$$2021", database="CalendarioAGP",
                     timeout=15, login_timeout=15, charset="UTF-8")
