#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Autenticación AGP Glass
=======================
- Valida usuario/contraseña contra MALLAS.APP_USUARIOS en Azure SQL
- Sincroniza lista SharePoint → SQL en background al hacer login

Configuración del servicio SharePoint:
  SP_USER / SP_PASS en variables de entorno, o editar las constantes abajo.
"""

import os, re, threading, hashlib
import sys

# ──────────────────────────────────────────────────────────
#  CONFIG BD (igual que el resto de la app)
# ──────────────────────────────────────────────────────────
BD_SERVER   = "agpcolombia.database.windows.net"
BD_PORT     = 1433
BD_USER     = "DevIngenieria"
BD_PASSWORD = "HiJE068i0LQVrwA"
BD_DATABASE = "AGP_Ingenieria"

# ──────────────────────────────────────────────────────────
#  CONFIG SHAREPOINT
# ──────────────────────────────────────────────────────────
SP_SITE      = "https://agpglass.sharepoint.com/sites/ServiciosITColombia"
SP_LIST_NAME = "Apps_UsuariosIngenieria_Col"
SP_USER      = os.environ.get("AGP_SP_USER", "abotero@agpglass.com")
SP_PASS      = os.environ.get("AGP_SP_PASS", "")   # llenar o usar variable de entorno

# ──────────────────────────────────────────────────────────
#  PYMSSQL
# ──────────────────────────────────────────────────────────
try:
    import pymssql as _pymssql
    _SQL_OK = True
except ImportError:
    _SQL_OK = False


def _conn():
    return _pymssql.connect(
        server=BD_SERVER, port=BD_PORT, user=BD_USER,
        password=BD_PASSWORD, database=BD_DATABASE,
        timeout=15, login_timeout=15, charset="UTF-8", tds_version="7.3",
    )


# ──────────────────────────────────────────────────────────
#  CREAR TABLA SI NO EXISTE
# ──────────────────────────────────────────────────────────
_DDL = """
IF NOT EXISTS (
    SELECT 1 FROM sys.tables t
    JOIN sys.schemas s ON t.schema_id = s.schema_id
    WHERE s.name='MALLAS' AND t.name='APP_USUARIOS'
)
CREATE TABLE MALLAS.APP_USUARIOS (
    id             INT IDENTITY(1,1) PRIMARY KEY,
    nombre         NVARCHAR(200),
    usuario        NVARCHAR(200) NOT NULL,
    contrasenia    NVARCHAR(200) NOT NULL,
    estatus        TINYINT       NOT NULL DEFAULT 1,
    es_admin       BIT           NOT NULL DEFAULT 0,
    sp_item_id     INT,
    actualizado_en DATETIME2     DEFAULT SYSDATETIME()
)
"""

def crear_tabla():
    if not _SQL_OK:
        return False
    try:
        cn = _conn()
        cur = cn.cursor()
        cur.execute(_DDL)
        cn.commit()
        cn.close()
        return True
    except Exception as e:
        print(f"[auth] crear_tabla error: {e}")
        return False


# ──────────────────────────────────────────────────────────
#  VALIDAR LOGIN
# ──────────────────────────────────────────────────────────
def validar_login(usuario: str, contrasenia: str):
    """
    Retorna dict con info del usuario si credenciales OK, None si no.
    Campos: nombre, usuario, es_admin, estatus
    """
    if not _SQL_OK:
        return None
    if not usuario or not contrasenia:
        return None
    try:
        cn = _conn()
        cur = cn.cursor()
        cur.execute(
            "SELECT nombre, usuario, es_admin, estatus, rol "
            "FROM MALLAS.APP_USUARIOS "
            "WHERE LOWER(usuario)=LOWER(%s) AND contrasenia=%s AND estatus=1",
            (usuario.strip(), contrasenia.strip())
        )
        row = cur.fetchone()
        cn.close()
        if row:
            return {
                "nombre":   row[0] or usuario,
                "usuario":  row[1],
                "es_admin": bool(row[2]),
                "estatus":  row[3],
                "rol":      (row[4] or "").strip().lower() or None,
            }
        return None
    except Exception as e:
        print(f"[auth] validar_login error: {e}")
        return None


# ──────────────────────────────────────────────────────────
#  SINCRONIZACIÓN SHAREPOINT → SQL
# ──────────────────────────────────────────────────────────
def _sync_sharepoint():
    """Descarga la lista SharePoint y hace UPSERT en MALLAS.APP_USUARIOS."""
    if not SP_PASS:
        print("[auth] SP_PASS no configurado — sync omitido")
        return

    try:
        from office365.sharepoint.client_context import ClientContext
        from office365.runtime.auth.user_credential import UserCredential
    except ImportError:
        print("[auth] office365-rest-python-client no instalado — sync omitido")
        print("       pip install office365-rest-python-client")
        return

    try:
        ctx = ClientContext(SP_SITE).with_credentials(
            UserCredential(SP_USER, SP_PASS)
        )
        lista   = ctx.web.lists.get_by_title(SP_LIST_NAME)
        campos  = ["ID", "Title", "Usuario", "Contrasenia", "Estatus", "Adm"]
        items   = lista.items.select(campos).get().execute_query()
        usuarios = []
        for it in items:
            p = it.properties
            usuarios.append({
                "sp_id":      p.get("ID"),
                "nombre":     (p.get("Title") or "").strip(),
                "usuario":    (p.get("Usuario") or "").strip(),
                "contrasenia":(p.get("Contrasenia") or "").strip(),
                "estatus":    1 if str(p.get("Estatus","1")) == "1" else 0,
                "es_admin":   1 if p.get("Adm") else 0,
            })
    except Exception as e:
        print(f"[auth] SharePoint fetch error: {e}")
        return

    if not usuarios:
        return

    try:
        cn = _conn()
        cur = cn.cursor()
        for u in usuarios:
            if not u["usuario"]:
                continue
            cur.execute(
                "SELECT id FROM MALLAS.APP_USUARIOS WHERE LOWER(usuario)=LOWER(%s)",
                (u["usuario"],)
            )
            row = cur.fetchone()
            if row:
                cur.execute(
                    "UPDATE MALLAS.APP_USUARIOS SET "
                    "nombre=%s, contrasenia=%s, estatus=%s, es_admin=%s, "
                    "sp_item_id=%s, actualizado_en=SYSDATETIME() "
                    "WHERE LOWER(usuario)=LOWER(%s)",
                    (u["nombre"], u["contrasenia"], u["estatus"],
                     u["es_admin"], u["sp_id"], u["usuario"])
                )
            else:
                cur.execute(
                    "INSERT INTO MALLAS.APP_USUARIOS "
                    "(nombre, usuario, contrasenia, estatus, es_admin, sp_item_id) "
                    "VALUES (%s,%s,%s,%s,%s,%s)",
                    (u["nombre"], u["usuario"], u["contrasenia"],
                     u["estatus"], u["es_admin"], u["sp_id"])
                )
        cn.commit()
        cn.close()
        print(f"[auth] Sync OK — {len(usuarios)} usuarios")
    except Exception as e:
        print(f"[auth] Sync SQL error: {e}")


def sincronizar_background():
    """Lanza el sync de SharePoint en un hilo daemon (no bloquea la app)."""
    t = threading.Thread(target=_sync_sharepoint, daemon=True)
    t.start()
