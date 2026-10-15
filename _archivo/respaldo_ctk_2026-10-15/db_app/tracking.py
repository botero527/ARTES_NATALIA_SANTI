# -*- coding: utf-8 -*-
"""
Rastreo de actividad — mallas.APP_SESIONES / mallas.APP_EVENTOS.

Registra cuándo cada usuario abre la app, qué hace adentro (cambios de
pestaña + acciones clave de cada módulo) y cuándo/cómo la cierra. Pensado
para reportes gerenciales: quién entró, a qué hora, cuánto se demoró de
verdad (no solo "la ventana estaba abierta"), qué módulos usa cada quien.

Reglas de diseño — para que esto NUNCA rompa nada de lo que ya funciona:
  1. Todo INSERT/UPDATE de este módulo pasa por una cola + un hilo aparte
     (ver _encolar). Si Azure SQL está lento o caído, la app sigue andando
     normal — lo único que pasa es que ese evento no queda registrado.
  2. Todas las funciones públicas atrapan sus propios errores. Ninguna
     propaga una excepción hacia quien la llama.
  3. Los timestamps se calculan con datetime.now() en Python (hora del PC,
     que en las oficinas de AGP Colombia ya está en hora de Bogotá) — a
     propósito NO se usa GETDATE()/SYSDATETIME() de Azure SQL en ningún
     lado de este módulo, porque el servidor de Azure corre en UTC y
     mezclar las dos cosas en una misma comparación (ej. "¿ya pasaron 20
     minutos?") da resultados absurdos (Bogotá va 5 horas atrás de UTC).
  4. Una sesión por proceso: esta app es de un solo usuario por PC/proceso,
     así que el id de sesión se guarda como estado del módulo (no hay que
     pasarlo de pestaña en pestaña) — cualquier pestaña solo llama
     registrar_evento(...) y ya sabe a cuál sesión pertenece.

Uso típico (ver agp_app.py):
    from db_app.tracking import (
        iniciar_sesion, registrar_evento, marcar_actividad,
        latido, cerrar_sesion, limpiar_sesiones_huerfanas,
    )
    limpiar_sesiones_huerfanas()          # al arrancar, antes de nada más
    iniciar_sesion(usuario_info)          # al loguearse
    registrar_evento("cambio_pestana", pestana="Scanner")
    latido()                              # cada pocos minutos (heartbeat)
    marcar_actividad()                    # en cualquier clic/tecla real
    cerrar_sesion("normal")               # al cerrar la ventana
"""

import os, sys, socket, threading, queue, datetime

_DIR = os.path.dirname(__file__)
sys.path.insert(0, _DIR)
from importar_excel import conectar

# Subir manualmente en cada release — queda guardado en cada sesión, útil
# para saber "esto lo reportó alguien con la versión vieja".
VERSION_APP = "V5"

# ── estado interno del módulo ────────────────────────────────────────────
_sesion = {"id": None, "usuario": None}
_lock   = threading.Lock()
_cola   = queue.Queue()
_worker_iniciado = False
_ultima_actividad_enviada = [None]   # throttle de marcar_actividad()


def _ahora():
    """Hora 'de Bogotá' = hora local del PC (ver nota de diseño arriba)."""
    return datetime.datetime.now()


def _hostname():
    try:
        return socket.gethostname()
    except Exception:
        return None


# ── cola de fondo: acá pasa todo el trabajo que toca Azure SQL ──────────
def _asegurar_worker():
    global _worker_iniciado
    if _worker_iniciado:
        return
    _worker_iniciado = True
    threading.Thread(target=_worker_loop, daemon=True).start()


def _worker_loop():
    while True:
        fn, args = _cola.get()
        try:
            fn(*args)
        except Exception as e:
            print(f"[tracking] error de fondo ({fn.__name__}): {e}")


def _encolar(fn, *args):
    _asegurar_worker()
    _cola.put((fn, args))


def _recortar(detalle):
    if detalle in (None, ""):
        return None
    return str(detalle)[:500]


# ── sesión ────────────────────────────────────────────────────────────────
def iniciar_sesion(usuario_info):
    """
    Abre una sesión nueva. Se llama UNA vez, al arrancar AGPApp ya logueado.
    Esta sí se ejecuta de una (no encolada) porque hace falta el id real
    antes de poder registrar cualquier evento — es un solo INSERT chiquito,
    no alcanza a notarse en el arranque.
    """
    info    = usuario_info or {}
    usuario = info.get("usuario") or info.get("nombre") or "desconocido"
    nombre  = info.get("nombre")
    rol     = info.get("rol")
    ahora   = _ahora()
    try:
        cn = conectar()
        cur = cn.cursor()
        cur.execute(
            "INSERT INTO mallas.APP_SESIONES "
            "(usuario, nombre, rol, equipo, version_app, inicio, ultimo_latido) "
            "OUTPUT INSERTED.id VALUES (?, ?, ?, ?, ?, ?, ?)",
            (usuario, nombre, rol, _hostname(), VERSION_APP, ahora, ahora))
        row = cur.fetchone()
        cn.commit()
        cn.close()
        if row:
            with _lock:
                _sesion["id"]      = row[0]
                _sesion["usuario"] = usuario
    except Exception as e:
        print(f"[tracking] no se pudo abrir sesión: {e}")


def _sesion_activa():
    with _lock:
        return _sesion["id"], _sesion["usuario"]


# ── eventos ───────────────────────────────────────────────────────────────
def registrar_evento(accion, pestana=None, detalle=None, duracion_ms=None, exito=None):
    """
    Guarda un evento de la sesión activa. Si todavía no hay sesión (falló
    el login a BD, o Azure no respondió a tiempo en el arranque), no hace
    nada — nunca lanza, nunca bloquea.
    """
    sesion_id, usuario = _sesion_activa()
    if sesion_id is None:
        return
    momento = _ahora()
    _encolar(_insertar_evento, sesion_id, usuario, momento, pestana, accion,
              _recortar(detalle), duracion_ms, exito)


def registrar_evento_sin_sesion(usuario, accion, detalle=None, exito=None):
    """Para eventos que pasan ANTES de iniciar sesión (ej. login fallido)."""
    momento = _ahora()
    _encolar(_insertar_evento, None, usuario or "desconocido", momento, None,
              accion, _recortar(detalle), None, exito)


def _insertar_evento(sesion_id, usuario, momento, pestana, accion, detalle, duracion_ms, exito):
    cn = conectar()
    try:
        cur = cn.cursor()
        cur.execute(
            "INSERT INTO mallas.APP_EVENTOS "
            "(sesion_id, usuario, momento, pestana, accion, detalle, duracion_ms, exito) "
            "VALUES (?, ?, ?, ?, ?, ?, ?, ?)",
            (sesion_id, usuario, momento, pestana, accion, detalle, duracion_ms, exito))
        cn.commit()
    finally:
        cn.close()


# ── latido (sigue viva) y actividad real (clic/tecla) ───────────────────
def latido():
    """Heartbeat: llamar cada pocos minutos mientras la app siga abierta."""
    sesion_id, _ = _sesion_activa()
    if sesion_id is None:
        return
    _encolar(_actualizar_latido, sesion_id, _ahora())


def _actualizar_latido(sesion_id, momento):
    cn = conectar()
    try:
        cur = cn.cursor()
        cur.execute("UPDATE mallas.APP_SESIONES SET ultimo_latido=? WHERE id=?",
                    (momento, sesion_id))
        cn.commit()
    finally:
        cn.close()


def marcar_actividad():
    """
    Llamar en cualquier clic/tecla real del usuario (bind_all a nivel de
    ventana). Throttle de 30s — no le pega a la BD en cada movimiento,
    solo necesitamos saber "más o menos hasta qué hora estuvo trabajando".
    """
    sesion_id, _ = _sesion_activa()
    if sesion_id is None:
        return
    ahora  = _ahora()
    ultima = _ultima_actividad_enviada[0]
    if ultima is not None and (ahora - ultima).total_seconds() < 30:
        return
    _ultima_actividad_enviada[0] = ahora
    _encolar(_actualizar_actividad, sesion_id, ahora)


def _actualizar_actividad(sesion_id, momento):
    cn = conectar()
    try:
        cur = cn.cursor()
        cur.execute("UPDATE mallas.APP_SESIONES SET ultima_actividad=? WHERE id=?",
                    (momento, sesion_id))
        cn.commit()
    finally:
        cn.close()


# ── cierre ────────────────────────────────────────────────────────────────
def cerrar_sesion(tipo="normal"):
    """
    Cierra la sesión activa (fin + duración). Se llama al cerrar la
    ventana. A propósito NO se encola — la app se está terminando, y
    queremos que este UPDATE alcance a salir antes de que el proceso muera.
    """
    with _lock:
        sesion_id = _sesion["id"]
        _sesion["id"] = None
    if sesion_id is None:
        return
    fin = _ahora()
    try:
        cn = conectar()
        cur = cn.cursor()
        cur.execute(
            "UPDATE mallas.APP_SESIONES SET fin=?, "
            "duracion_seg=DATEDIFF(SECOND, inicio, ?), cierre_tipo=? WHERE id=?",
            (fin, fin, tipo, sesion_id))
        cn.commit()
        cn.close()
    except Exception as e:
        print(f"[tracking] no se pudo cerrar sesión: {e}")


def limpiar_sesiones_huerfanas(minutos_frios=20):
    """
    Al iniciar la app: cierra sesiones (de esta máquina o de cualquier
    otra) que quedaron sin `fin` porque el proceso se mató sin avisar
    (crash, Alt+F4 forzado por el SO, corte de luz, apagón del PC). Usa el
    último latido como hora de cierre — es la mejor aproximación real de
    "hasta cuándo estuvo abierta".

    Igual de conservador que limpiar_pendientes_huerfanos (asignaciones.py):
    solo toca sesiones cuyo latido ya está viejo, nunca una sesión activa
    de otro PC que sigue mandando latidos.

    El umbral se calcula en Python (no con GETDATE()) — ver nota de diseño
    al inicio del archivo, mezclar Bogotá-local con Azure-UTC rompe la
    comparación.
    """
    umbral = _ahora() - datetime.timedelta(minutes=minutos_frios)
    try:
        cn = conectar()
        cur = cn.cursor()
        cur.execute(
            "UPDATE mallas.APP_SESIONES SET "
            "fin=ultimo_latido, "
            "duracion_seg=DATEDIFF(SECOND, inicio, ultimo_latido), "
            "cierre_tipo='huerfana' "
            "WHERE fin IS NULL AND ultimo_latido < ?",
            (umbral,))
        n = cur.rowcount
        cn.commit()
        cn.close()
        return n
    except Exception as e:
        print(f"[tracking] limpiar huérfanas error: {e}")
        return 0
