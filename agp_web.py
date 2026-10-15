#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
AGP Glass Suite — versión HTML (pywebview).

Misma app de siempre (Crear Arte, Consultar BD, Gestión BD, Scanner, Roles,
Logos Laser) con interfaz HTML en ui_web/, en la misma familia visual que el
Panel AGP, PIEZAS_PLANAS y Tickets. La lógica de negocio NO cambió:
  - servicios.py            → consultas / guardado / scanner (copiado de agp_app.py)
  - crear_arte_acad.py      → pipeline AutoCAD (sin tocar)
  - db_app/asignaciones.py  → reservar / confirmar / cancelar (sin tocar)
  - db_app/tracking.py      → rastreo de actividad (sin tocar)
agp_app.py (CustomTkinter) sigue intacto como respaldo.

La ventana usa el marco NATIVO de Windows (frameless=False) → minimizar,
maximizar/restaurar, cerrar, Aero Snap y doble clic en el título funcionan
como en cualquier programa.

Contrato con la UI: cada método de Api devuelve un dict {"ok": bool, ...}.
"""

import os, sys, time, json, queue, shutil, threading, subprocess, traceback, uuid

import webview

import servicios as S
from servicios import ErrorGuardadoArte

try:
    import pythoncom, win32com.client
    _COM_OK = True
except ImportError:
    _COM_OK = False

try:
    from autocad_ops import AutoCADMotor
    _MOTOR_OK = True
except Exception:
    _MOTOR_OK = False

try:
    from db_app.asignaciones import (
        reservar as _reservar, confirmar as _confirmar, cancelar as _cancelar,
        actualizar_ruta_arte as _actualizar_ruta_arte,
        anular_asignacion as _anular_asignacion,
        limpiar_pendientes_huerfanos as _limpiar_pendientes_huerfanos,
    )
    _ASIGN_OK = True
except Exception as _e_as:
    _ASIGN_OK = False
    print(f"[asignaciones] no disponible: {_e_as}")

try:
    from db_app.logos_laser import listar as _ll_listar, crear_fila as _ll_crear, \
        actualizar_campo as _ll_actualizar
    _LL_OK = True
except Exception as _e_ll:
    _LL_OK = False
    print(f"[logos_laser] no disponible: {_e_ll}")

try:
    from db_app.tracking import (
        iniciar_sesion as _iniciar_sesion, registrar_evento as _registrar_evento,
        marcar_actividad as _marcar_actividad, latido as _latido,
        cerrar_sesion as _cerrar_sesion,
        limpiar_sesiones_huerfanas as _limpiar_sesiones_huerfanas,
    )
except Exception as _e_tr:
    print(f"[tracking] no disponible: {_e_tr}")
    def _iniciar_sesion(*_a, **_k): pass
    def _registrar_evento(*_a, **_k): pass
    def _marcar_actividad(*_a, **_k): pass
    def _latido(*_a, **_k): pass
    def _cerrar_sesion(*_a, **_k): pass
    def _limpiar_sesiones_huerfanas(*_a, **_k): return 0


if getattr(sys, "frozen", False):
    RES_DIR = getattr(sys, "_MEIPASS", os.path.dirname(sys.executable))
else:
    RES_DIR = os.path.dirname(os.path.abspath(__file__))
INDEX_PATH = os.path.join(RES_DIR, "ui_web", "index.html")

# Preferencias por PC (tema, idioma, menú fijo, últimas rutas). Van a disco porque
# pywebview sirve la página en un puerto distinto cada vez y el localStorage no dura.
PREFS_PATH = os.path.join(os.environ.get("LOCALAPPDATA") or os.path.expanduser("~"),
                          "AGP_Artes", "prefs.json")


def _leer_prefs():
    try:
        with open(PREFS_PATH, encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}

VERSION = "3.0"

# (id, nombre — el nombre es el mismo que usa el tracking y los permisos)
PAGINAS = [
    ("arte",    "Crear Arte"),
    ("bd",      "Consultar BD"),
    ("gestion", "Gestión BD"),
    ("scanner", "Scanner"),
    ("roles",   "Roles y Permisos"),
    ("logos",   "Logos Laser"),
]
PERMISOS = {
    "Crear Arte":       {"admin", "dibujante"},
    "Consultar BD":     {"admin", "dibujante"},
    "Gestión BD":       {"admin", "dibujante"},
    "Scanner":          {"admin", "dibujante", "planta"},
    "Roles y Permisos": {"admin"},
    "Logos Laser":      {"admin"},
}


def _ok(**kw):
    kw["ok"] = True
    return kw


def _err(e):
    return {"ok": False, "error": str(e)}


class Api:
    """Expuesta a JS como window.pywebview.api.

    OJO: nunca guardar la ventana como atributo (self.window = ...). pywebview
    recorre los atributos de este objeto y con la ventana adentro entra en una
    recursión infinita al arrancar (bug visto en PIEZAS_PLANAS). Siempre se
    pide webview.windows[0] en el momento.
    """

    def __init__(self):
        self._usuario = None
        self._logs    = queue.Queue()
        self._eventos = queue.Queue()
        self._busy    = {"arte": False, "sync": False}
        self._reservas = {}          # rid → reserva PENDIENTE abierta en la UI
        self._lock    = threading.Lock()
        self._sesion_abierta = False

    # ── utilidades internas ──────────────────────────────────────────────
    def _log(self, canal, msg, tag=""):
        self._logs.put({"canal": canal, "msg": str(msg), "tag": tag,
                        "ts": time.strftime("%H:%M:%S")})

    def _evento(self, tipo, **kw):
        kw["tipo"] = tipo
        self._eventos.put(kw)

    def _alerta(self, nivel, titulo, mensaje):
        self._evento("alerta", nivel=nivel, titulo=titulo, mensaje=str(mensaje))

    def _nombre_usuario(self):
        u = self._usuario or {}
        return u.get("nombre") or u.get("usuario") or "desconocido"

    def _rol(self):
        return ((self._usuario or {}).get("rol") or "").lower()

    def _puede(self, pagina):
        return self._rol() in PERMISOS.get(pagina, set())

    # ── sesión / login ───────────────────────────────────────────────────
    def inicio(self):
        """Primera llamada de la UI. Si el Panel AGP abrió la app con un
        token SSO válido, entra directo sin pedir login."""
        threading.Thread(target=self._crear_tabla_usuarios, daemon=True).start()
        if self._usuario is None:
            try:
                import agp_sso
                from db_app import auth as _auth
                token = agp_sso.token_de_argv(sys.argv)
                if token:
                    payload = agp_sso.verificar_token(token)
                    if payload:
                        self._usuario = _auth.cargar_por_usuario(payload["usuario"])
            except Exception as e:
                print(f"[sso] error: {e}")
        if self._usuario:
            self._post_login()
        return _ok(usuario=self._usuario_publico(), version=VERSION, prefs=_leer_prefs())

    def guardar_pref(self, clave, valor):
        try:
            prefs = _leer_prefs()
            prefs[str(clave)] = valor
            os.makedirs(os.path.dirname(PREFS_PATH), exist_ok=True)
            with open(PREFS_PATH, "w", encoding="utf-8") as f:
                json.dump(prefs, f, ensure_ascii=False, indent=1)
            return _ok()
        except Exception as e:
            return _err(e)

    def _crear_tabla_usuarios(self):
        try:
            from db_app.auth import crear_tabla
            crear_tabla()
        except Exception as e:
            print(f"[login] init BD error: {e}")

    def login(self, usuario, contrasenia):
        usuario = (usuario or "").strip()
        contrasenia = (contrasenia or "").strip()
        if not usuario or not contrasenia:
            return {"ok": False, "error": "campos"}
        try:
            from db_app.auth import validar_login, sincronizar_background
            info = validar_login(usuario, contrasenia)
        except Exception as e:
            return {"ok": False, "error": "conexion", "detalle": str(e)}
        if not info:
            return {"ok": False, "error": "credenciales"}
        sincronizar_background()
        self._usuario = info
        self._post_login()
        return _ok(usuario=self._usuario_publico())

    def _usuario_publico(self):
        if not self._usuario:
            return None
        u = self._usuario
        paginas = [{"id": pid, "nombre": nom} for pid, nom in PAGINAS
                   if self._puede(nom) and (pid != "logos" or _LL_OK)]
        return {"nombre": u.get("nombre") or u.get("usuario"), "usuario": u.get("usuario"),
                "rol": self._rol() or None, "paginas": paginas}

    def _post_login(self):
        """Arranque de sesión (una sola vez): limpiar huérfanos, tracking, latido."""
        if self._sesion_abierta or not self._rol():
            return
        self._sesion_abierta = True

        def _arranque():
            try:
                if _ASIGN_OK:
                    nv, ng, np_ = _limpiar_pendientes_huerfanos()
                    if nv + ng + np_ > 0:
                        print(f"[startup] Pendientes huerfanos cancelados: {nv}/{ng}/{np_}")
            except Exception as e:
                print(f"[startup] WARN limpiar pendientes: {e}")
            _limpiar_sesiones_huerfanas()
            _iniciar_sesion(self._usuario)
            _registrar_evento("login", detalle=f"rol={self._rol() or 'sin rol'}")

        def _latidos():
            while True:
                time.sleep(240)
                _latido()

        threading.Thread(target=_arranque, daemon=True).start()
        threading.Thread(target=_latidos, daemon=True).start()

    def test_conexion(self):
        try:
            S.db_connect().close()
            return _ok()
        except Exception as e:
            return _err(e)

    def actividad(self):
        _marcar_actividad()
        return _ok()

    def cambio_pestana(self, nombre):
        _registrar_evento("cambio_pestana", pestana=nombre)
        return _ok()

    def al_cerrar(self):
        """window.events.closing — X nativa, Alt+F4, apagado."""
        for rid in list(self._reservas):
            try:
                _cancelar(self._reservas.pop(rid))
            except Exception:
                pass
        if self._sesion_abierta:
            _registrar_evento("logout")
            _cerrar_sesion("normal")

    def cerrar_app(self):
        webview.windows[0].destroy()   # dispara events.closing → al_cerrar
        return _ok()

    # ── canal de logs / eventos (la UI hace polling cada ~300 ms) ───────
    def poll(self):
        logs, eventos = [], []
        try:
            while len(logs) < 400:
                logs.append(self._logs.get_nowait())
        except queue.Empty:
            pass
        try:
            while True:
                eventos.append(self._eventos.get_nowait())
        except queue.Empty:
            pass
        return _ok(logs=logs, eventos=eventos, busy=dict(self._busy))

    # ── ventana / sistema ────────────────────────────────────────────────
    def set_tema(self, oscuro):
        """Barra de título nativa oscura/clara (Windows 10 1809+ / 11)."""
        try:
            import ctypes
            hwnd = webview.windows[0].native.Handle.ToInt64()
            val = ctypes.c_int(1 if oscuro else 0)
            for attr in (20, 19):  # DWMWA_USE_IMMERSIVE_DARK_MODE (nuevo / viejo)
                if ctypes.windll.dwmapi.DwmSetWindowAttribute(
                        ctypes.c_void_p(hwnd), attr, ctypes.byref(val), 4) == 0:
                    break
        except Exception:
            pass
        return _ok()

    def copiar(self, texto):
        try:
            import win32clipboard
            win32clipboard.OpenClipboard()
            try:
                win32clipboard.EmptyClipboard()
                win32clipboard.SetClipboardText(str(texto), win32clipboard.CF_UNICODETEXT)
            finally:
                win32clipboard.CloseClipboard()
            return _ok()
        except Exception as e:
            return _err(e)

    def abrir_ruta(self, ruta):
        """Abre el Explorador con el archivo seleccionado (o la carpeta)."""
        try:
            ruta = (ruta or "").strip().strip('"')
            if os.path.isfile(ruta):
                subprocess.Popen(["explorer", "/select,", ruta])
            elif os.path.isdir(ruta):
                subprocess.Popen(["explorer", ruta])
            elif os.path.isdir(os.path.dirname(ruta)):
                subprocess.Popen(["explorer", os.path.dirname(ruta)])
            else:
                return {"ok": False, "error": "La ruta no existe o no hay acceso a la red"}
            return _ok()
        except Exception as e:
            return _err(e)

    def elegir_carpeta(self, inicial=""):
        r = webview.windows[0].create_file_dialog(
            webview.FOLDER_DIALOG, directory=inicial if os.path.isdir(inicial or "") else "")
        return _ok(ruta=(r[0] if r else "").replace("/", "\\"))

    def elegir_archivo(self, tipo="dwg", inicial=""):
        filtros = {
            "dwg":  ("AutoCAD DWG (*.dwg)", "Todos (*.*)"),
            "ruta": ("DWG / PDF (*.dwg;*.pdf)", "Todos (*.*)"),
        }.get(tipo, ("Todos (*.*)",))
        r = webview.windows[0].create_file_dialog(
            webview.OPEN_DIALOG, directory=inicial if os.path.isdir(inicial or "") else "",
            file_types=filtros)
        return _ok(ruta=(r[0] if r else "").replace("/", "\\"))

    # ══════════════════════════════════════════════════════════════════════
    #  CREAR ARTE
    # ══════════════════════════════════════════════════════════════════════
    def arte_estado(self):
        return _ok(com=_COM_OK, motor=_MOTOR_OK, pipeline=S._PIPELINE_OK, asign=_ASIGN_OK)

    def arte_validar(self, base, dwg):
        base = (base or "").strip().strip('"')
        dwg  = (dwg or "").strip().strip('"')
        return _ok(base_ok=os.path.isdir(base), dwg_ok=os.path.isfile(dwg))

    def _arte_ocupar(self):
        with self._lock:
            if self._busy["arte"]:
                return False
            self._busy["arte"] = True
            return True

    def _arte_liberar(self):
        self._busy["arte"] = False

    def arte_extraer(self, dwg):
        dwg = (dwg or "").strip().strip('"')
        if not self._arte_ocupar():
            return {"ok": False, "error": "Ya hay un proceso de AutoCAD corriendo"}
        threading.Thread(target=self._t_extraer, args=(dwg,), daemon=True).start()
        return _ok()

    def _t_extraer(self, dwg):
        L = lambda m, t="": self._log("arte", m, t)
        L("=" * 50)
        L("EXTRAER PLANO...", "ok")
        nombre = os.path.splitext(os.path.basename(dwg))[0]
        try:
            dest = os.path.join(S.ruta_planos(dwg), f"{nombre}_PLANO.dwg")
            if not _MOTOR_OK: raise RuntimeError("AutoCADMotor no disponible")
            if not _COM_OK:   raise RuntimeError("pywin32 no disponible")
            motor = AutoCADMotor()
            try:
                motor.extraer_layers(dwg, dest, log_fn=lambda m: L(m, "dim"))
            finally:
                motor.quit()  # quit() ya llama CoUninitialize
            L(f"Guardado → {dest}", "ok")
            subprocess.Popen(["explorer", "/select,", dest])
            _registrar_evento("extraer_plano", pestana="Crear Arte", detalle=dest, exito=True)
        except Exception as e:
            L(str(e), "err")
            _registrar_evento("extraer_plano", pestana="Crear Arte", detalle=str(e)[:200], exito=False)
        finally:
            self._arte_liberar()

    def arte_buscar(self, base, dwg=""):
        base = (base or "").strip()
        dwg  = (dwg or "").strip().strip('"')
        if not self._arte_ocupar():
            return {"ok": False, "error": "Ya hay un proceso corriendo"}
        try:
            cods = S.extraer_codigos(dwg) if dwg else []
            self._log("arte", "BUSCAR ARTE...", "ok")
            res = S.buscar_artes(base, cods)
            self._log("arte", f"{len(res)} arte(s) encontrados.", "ok" if res else "warn")
            _registrar_evento("buscar_arte", pestana="Crear Arte", detalle=f"{len(res)} resultado(s)")
            return _ok(resultados=res)
        except Exception as e:
            self._log("arte", str(e), "err")
            return _err(e)
        finally:
            self._arte_liberar()

    def arte_iniciar(self, modo, dwg, valores, rid=None, compensar=False):
        """modo = 'crear' | 'todo'. Llega después de que el usuario aceptó el
        cajetín en la UI. Primero confirma la asignación (igual que hacía el
        aceptar() del cajetín Tk, con ruta "" porque el arte aún no existe),
        y luego lanza el pipeline en un hilo."""
        dwg = (dwg or "").strip().strip('"')
        valores = {k: str(v or "").strip() for k, v in (valores or {}).items()}
        if not os.path.isfile(dwg):
            return {"ok": False, "error": "Selecciona el DWG del plano."}
        if not self._arte_ocupar():
            return {"ok": False, "error": "Ya hay un proceso de AutoCAD corriendo"}

        reserva = self._reservas.pop(rid, None) if rid else None
        if reserva is not None:
            nombre_plano = os.path.splitext(os.path.basename(dwg))[0]
            try:
                _confirmar(reserva,
                           vehiculo     = valores.get("VEHICULO", ""),
                           version      = valores.get("MODELO", ""),
                           pieza        = valores.get("PIEZA", ""),
                           cod_vehiculo = valores.get("NAGS", ""),
                           cod_completo = valores.get("COD PLANO", "") or nombre_plano,
                           ruta_archivo = "",
                           responsable  = valores.get("DIBUJO", "") or None)
            except Exception as e:
                self._log("arte", f"WARN confirmar asignación: {e}", "warn")

        try:
            arte0 = S.ruta_arte_salida(dwg, valores.get("MALLA", ""), valores.get("PIEZA", ""),
                                       valores.get("NOMBRE ARTE", ""))
        except Exception as e:
            self._arte_liberar()
            return _err(e)
        destino = self._t_crear if modo == "crear" else self._t_todo
        threading.Thread(target=destino, args=(dwg, valores, arte0, bool(compensar)),
                         daemon=True).start()
        return _ok(arte=arte0)

    def _t_crear(self, dwg, valores, arte0, compensar):
        L = lambda m, t="": self._log("arte", m, t)
        L("=" * 50)
        L("CREAR ARTE...", "ok")
        try:
            S.crear_arte_autocad(dwg, log_fn=lambda m: L(m, "dim"), valores_cajetin=valores,
                                 ruta_salida=arte0, compensar=compensar)
            L(f"Arte completado ✔  {os.path.basename(arte0)}", "ok")
            self._guardar_ruta_bd(valores, arte0)
            self._evento("arte_done")
            _registrar_evento("crear_arte", pestana="Crear Arte",
                              detalle=os.path.basename(arte0), exito=True)
        except ErrorGuardadoArte as e:
            L(f"Arte creado, pero falló al guardar: {e}", "warn")
            self._alerta("warning", "Arte creado — Error al guardar", e)
            _registrar_evento("crear_arte", pestana="Crear Arte", detalle=str(e)[:200], exito=False)
        except Exception as e:
            L(f"ERROR: {e}", "err")
            L(traceback.format_exc(), "err")
            S.anular_bd(valores, L)
            _registrar_evento("crear_arte", pestana="Crear Arte", detalle=str(e)[:200], exito=False)
        finally:
            self._arte_liberar()

    def _guardar_ruta_bd(self, valores, arte0):
        vitro_bd = valores.get("VITRO", "").strip()
        malla_bd = valores.get("MALLA", "").strip()
        if (vitro_bd or malla_bd) and _ASIGN_OK:
            try:
                _actualizar_ruta_arte(vitro_bd, malla_bd, arte0)
                self._log("arte", "Ruta guardada en BD ✔", "dim")
            except Exception as _re:
                self._log("arte", f"WARN ruta BD: {_re}", "warn")

    def _t_todo(self, dwg, valores, arte0, compensar):
        L = lambda m, t="": self._log("arte", m, t)
        L("=" * 50)
        L("TODO EN UNO...", "ok")
        if not _MOTOR_OK or not _COM_OK or not S._PIPELINE_OK:
            L("Faltan dependencias: AutoCADMotor / pywin32 / pipeline", "err")
            self._arte_liberar()
            return
        nombre = os.path.splitext(os.path.basename(dwg))[0]
        plano  = os.path.join(S.ruta_planos(dwg), f"{nombre}_PLANO.dwg")
        pythoncom.CoInitialize()
        todo_error = None
        try:
            # 1. Extraer plano
            motor = AutoCADMotor()
            try:
                motor.extraer_layers(dwg, plano, log_fn=lambda m: L(m, "dim"))
            finally:
                motor.quit()
            L("Plano extraído ✔", "ok")

            # Esperar a que AutoCAD termine de cerrar el documento
            time.sleep(2.5)
            for _ in range(8):
                try:
                    _t = win32com.client.GetActiveObject("AutoCAD.Application")
                    _t.Documents
                    break
                except Exception:
                    time.sleep(0.8)

            # 2. Crear arte de cada pieza
            n = S.crear_arte_autocad(plano, log_fn=lambda m: L(m, "dim"), valores_cajetin=valores,
                                     ruta_salida=arte0, perim_index=0, _com_ya_init=True,
                                     compensar=compensar)
            L(f"Arte guardado ✔  {os.path.basename(arte0)}", "ok")
            self._guardar_ruta_bd(valores, arte0)

            malla_bd = valores.get("MALLA", "").strip()
            pieza_bd = valores.get("PIEZA", "").strip()
            for i in range(1, n or 1):
                copia  = plano.replace("_PLANO.dwg", f"_PLANO_p{i+1}.dwg")
                arte_i = S.ruta_arte_salida(dwg, malla_bd, f"{pieza_bd} {i+1}".strip())
                try:
                    shutil.copy2(plano, copia)
                    S.crear_arte_autocad(copia, log_fn=lambda m: L(m, "dim"), valores_cajetin=valores,
                                         ruta_salida=arte_i, perim_index=i, _com_ya_init=True,
                                         compensar=compensar)
                    L(f"Arte {i+1} ✔  {os.path.basename(arte_i)}", "ok")
                except Exception as e:
                    L(f"Pieza {i+1}: {e}", "warn")
        except ErrorGuardadoArte as e:
            L(f"Arte creado, pero falló al guardar: {e}", "warn")
            self._alerta("warning", "Arte creado — Error al guardar", e)
            todo_error = str(e)
        except Exception as e:
            L(f"ERROR: {e}", "err")
            L(traceback.format_exc(), "err")
            S.anular_bd(valores, L)
            self._alerta("error", "Error — Todo en Uno", e)
            todo_error = str(e)
        finally:
            pythoncom.CoUninitialize()

        L("Todo en Uno completado.", "ok")
        self._evento("arte_done")
        try:
            subprocess.Popen(["explorer", os.path.dirname(arte0)])
        except Exception:
            pass
        _registrar_evento("todo_en_uno", pestana="Crear Arte",
                          detalle=todo_error or os.path.basename(arte0),
                          exito=todo_error is None)
        self._arte_liberar()

    def arte_superponer(self, ruta_arte, dwg):
        dwg = (dwg or "").strip().strip('"')
        if not dwg or not os.path.isfile(dwg):
            return {"ok": False, "error": "Indica el plano DWG para superponer."}
        if not _COM_OK:
            return {"ok": False, "error": "pywin32 no disponible"}
        if not self._arte_ocupar():
            return {"ok": False, "error": "Ya hay un proceso de AutoCAD corriendo"}
        self._log("arte", f"Superponiendo: {os.path.basename(ruta_arte)}", "ok")

        def _t():
            pythoncom.CoInitialize()
            try:
                from arte_maker import _overlay_autocad
                _overlay_autocad(ruta_arte, dwg, log_fn=lambda m: self._log("arte", m, "dim"))
                self._log("arte", "Superposición lista en AutoCAD.", "ok")
            except Exception as e:
                self._log("arte", str(e), "err")
            finally:
                pythoncom.CoUninitialize()
                self._arte_liberar()

        threading.Thread(target=_t, daemon=True).start()
        return _ok()

    # ── reservas vitro / malla (cajetín y Separar) ──────────────────────
    def reservar(self, n_vitros, n_grandes, n_pequenas, rid_anterior=None):
        if not _ASIGN_OK:
            return {"ok": False, "error": "Módulo de asignaciones no disponible"}
        nv, ng, np_ = int(n_vitros or 0), int(n_grandes or 0), int(n_pequenas or 0)
        if nv == 0 and ng == 0 and np_ == 0:
            return {"ok": False, "error": "Indica al menos 1 vitro o malla."}
        aviso = None
        if rid_anterior and rid_anterior in self._reservas:
            try:
                _cancelar(self._reservas.pop(rid_anterior))
            except Exception as e:
                aviso = (f"No se pudo liberar la reserva anterior: {e}. Los números anteriores "
                         "quedarán PENDIENTE hasta que el limpiador automático los libere (30 min).")
        try:
            res = _reservar(nv, ng, np_)
        except Exception as e:
            return _err(e)
        rid = uuid.uuid4().hex
        self._reservas[rid] = res
        return _ok(rid=rid, vitros=[str(x) for x in res["vitros"]],
                   grandes=[str(x) for x in res["grandes"]],
                   pequenas=[str(x) for x in res["pequenas"]], aviso=aviso)

    def cancelar_reserva(self, rid):
        res = self._reservas.pop(rid, None) if rid else None
        if res is not None:
            try:
                _cancelar(res)
            except Exception as e:
                return _err(e)
        return _ok()

    def confirmar_separar(self, rid, datos):
        res = self._reservas.get(rid)
        if res is None:
            return {"ok": False, "error": "Primero reserva los números."}
        d = {k: (str(v).strip() or None) if v is not None else None for k, v in (datos or {}).items()}
        try:
            _confirmar(res,
                       vehiculo=d.get("vehiculo"), version=d.get("version"),
                       pieza=d.get("pieza"), cod_vehiculo=d.get("cod_vehiculo"),
                       cod_completo=d.get("cod_completo"),
                       bnerig=d.get("bnerig") or "BN", tipo=d.get("bnerig") or "BN",
                       ruta_archivo=d.get("ruta"), responsable=d.get("responsable"))
            self._reservas.pop(rid, None)
            _registrar_evento("separar_vitro_malla", pestana="Gestión BD", exito=True)
            return _ok()
        except Exception as e:
            _registrar_evento("separar_vitro_malla", pestana="Gestión BD",
                              detalle=str(e)[:200], exito=False)
            return _err(e)

    # ══════════════════════════════════════════════════════════════════════
    #  CONSULTAR BD
    # ══════════════════════════════════════════════════════════════════════
    def bd_stats(self):
        try:
            return _ok(stats=S.estadisticas())
        except Exception as e:
            return _err(e)

    def bd_consultar(self, tab, q=""):
        try:
            return _ok(**S.consultar(tab, q))
        except Exception as e:
            return _err(e)

    def sync_excel(self):
        with self._lock:
            if self._busy["sync"]:
                return {"ok": False, "error": "Ya hay una sincronización corriendo"}
            self._busy["sync"] = True

        def _t():
            ok = False
            try:
                import importlib
                import db_app.importar_excel as imp
                importlib.reload(imp)
                ok = imp.main(log_fn=lambda m, tag="": self._log("sync", m, tag))
            except Exception as e:
                self._log("sync", f"Error al importar módulo: {e}", "err")
            finally:
                self._busy["sync"] = False
                self._evento("sync_done", exito=bool(ok))
                _registrar_evento("sync_excel", pestana="Consultar BD", exito=bool(ok))

        threading.Thread(target=_t, daemon=True).start()
        return _ok()

    # ══════════════════════════════════════════════════════════════════════
    #  GESTIÓN BD
    # ══════════════════════════════════════════════════════════════════════
    def gestion_buscar(self, tab, q=""):
        try:
            return _ok(**S.gestion_buscar(tab, q))
        except Exception as e:
            return _err(e)

    def gestion_guardar(self, tab, pk_val, valores):
        tabla = S.TABLA_SQL.get(tab, tab)
        try:
            S.guardar_edicion(tab, pk_val, valores or {}, self._nombre_usuario())
            _registrar_evento("editar_registro", pestana="Gestión BD",
                              detalle=f"{tabla} {pk_val}", exito=True)
            return _ok()
        except Exception as e:
            _registrar_evento("editar_registro", pestana="Gestión BD",
                              detalle=f"{tabla} {pk_val}: {e}"[:200], exito=False)
            return _err(e)

    def gestion_anular(self, tab, pk_val):
        if tab not in ("vitrojet", "grandes", "pequenas"):
            return {"ok": False, "error": "Esta tabla no se anula"}
        try:
            res = _anular_asignacion(tab, pk_val)
            _registrar_evento("anular_asignacion", pestana="Gestión BD",
                              detalle=f"{tab} {pk_val}", exito=True)
            return _ok(resultado=res)
        except Exception as e:
            _registrar_evento("anular_asignacion", pestana="Gestión BD",
                              detalle=f"{tab} {pk_val}: {e}"[:200], exito=False)
            return _err(e)

    def gestion_formulario(self, tabla):
        try:
            return _ok(**S.formulario_insertar(tabla))
        except Exception as e:
            return _err(e)

    def gestion_insertar(self, tabla, valores):
        try:
            pk = S.insertar_manual(tabla, valores or {})
            _registrar_evento("insertar_manual", pestana="Gestión BD",
                              detalle=f"{tabla} {pk}", exito=True)
            return _ok(pk=pk)
        except Exception as e:
            _registrar_evento("insertar_manual", pestana="Gestión BD",
                              detalle=f"{tabla}: {e}"[:200], exito=False)
            return _err(e)

    # ══════════════════════════════════════════════════════════════════════
    #  SCANNER
    # ══════════════════════════════════════════════════════════════════════
    def scanner_buscar(self, orden):
        orden = (orden or "").strip()
        if not orden:
            return {"ok": False, "error": "vacio"}
        try:
            r = S.buscar_orden(orden)
            if r.get("ok"):
                _registrar_evento("escaneo_orden", pestana="Scanner",
                                  detalle=f"orden={orden} zfer={r['zfer']}", exito=True)
            else:
                _registrar_evento("escaneo_orden", pestana="Scanner",
                                  detalle=f"orden={orden} no encontrada", exito=False)
            return r
        except Exception as e:
            _registrar_evento("escaneo_orden", pestana="Scanner",
                              detalle=f"orden={orden}: {e}"[:200], exito=False)
            return _err(str(e)[:200])

    # ══════════════════════════════════════════════════════════════════════
    #  ROLES
    # ══════════════════════════════════════════════════════════════════════
    def roles_listar(self):
        if not self._puede("Roles y Permisos"):
            return {"ok": False, "error": "Sin permiso"}
        try:
            return _ok(usuarios=S.listar_usuarios(), roles=S.ROLES)
        except Exception as e:
            return _err(e)

    def roles_guardar(self, uid, rol, estatus, usuario_label=""):
        if not self._puede("Roles y Permisos"):
            return {"ok": False, "error": "Sin permiso"}
        try:
            rol_sql = S.guardar_usuario(uid, rol, estatus)
            _registrar_evento("editar_rol_usuario", pestana="Roles y Permisos",
                              detalle=f"{usuario_label} → rol={rol_sql}, estatus={1 if estatus else 0}",
                              exito=True)
            return _ok()
        except Exception as e:
            _registrar_evento("editar_rol_usuario", pestana="Roles y Permisos",
                              detalle=str(e)[:200], exito=False)
            return _err(e)

    # ══════════════════════════════════════════════════════════════════════
    #  LOGOS LASER
    # ══════════════════════════════════════════════════════════════════════
    def ll_listar(self):
        try:
            filas = _ll_listar()
            for f in filas:
                for k, v in list(f.items()):
                    if hasattr(v, "strftime"):
                        f[k] = v.strftime("%Y-%m-%d %H:%M")
                    elif isinstance(v, bool):
                        f[k] = 1 if v else 0
            return _ok(filas=filas)
        except Exception as e:
            return _err(e)

    def ll_crear(self):
        try:
            fila = _ll_crear(responsable=self._nombre_usuario())
            for k, v in list(fila.items()):
                if hasattr(v, "strftime"):
                    fila[k] = v.strftime("%Y-%m-%d %H:%M")
                elif isinstance(v, bool):
                    fila[k] = 1 if v else 0
            _registrar_evento("nueva_fila_logos_laser", pestana="Logos Laser",
                              detalle=fila.get("id"), exito=True)
            return _ok(fila=fila)
        except Exception as e:
            _registrar_evento("nueva_fila_logos_laser", pestana="Logos Laser",
                              detalle=str(e)[:200], exito=False)
            return _err(e)

    def ll_actualizar(self, id_fila, campo, valor):
        try:
            _ll_actualizar(id_fila, campo, valor)
            _registrar_evento("editar_logos_laser", pestana="Logos Laser",
                              detalle=f"{id_fila}.{campo}", exito=True)
            return _ok()
        except Exception as e:
            _registrar_evento("editar_logos_laser", pestana="Logos Laser",
                              detalle=f"{id_fila}.{campo}: {e}"[:200], exito=False)
            return _err(e)


def main():
    api = Api()
    window = webview.create_window(
        "AGP Glass — Artes",
        INDEX_PATH,
        js_api=api,
        width=1380,
        height=880,
        min_size=(1000, 660),
        resizable=True,
        frameless=False,        # marco nativo: minimizar / maximizar / cerrar de Windows
        text_select=True,
        zoomable=True,
        background_color="#07111F",
    )
    window.events.closing += api.al_cerrar
    webview.start(debug="--debug" in sys.argv)


if __name__ == "__main__":
    main()
