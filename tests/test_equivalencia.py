# -*- coding: utf-8 -*-
"""
Pruebas de equivalencia: la app nueva (agp_web.py / servicios.py / ui_web) tiene
que hacer EXACTAMENTE lo mismo que la vieja (agp_app.py y sus diálogos Tk).

Cada prueba ejecuta el código ORIGINAL (extraído literal con ast cuando vive
dentro de un diálogo Tk) y el nuevo con las mismas entradas, y compara:
  - el SQL + parámetros que mandan a la BD (BD grabadora, sin tocar Azure)
  - los argumentos con que llaman a asignaciones / AutoCAD / tracking
  - los resultados de las consultas de lectura contra Azure real (solo SELECT)

Correr:  py -3.14 -m unittest tests.test_equivalencia -v
         (las de BD real se saltan con AGP_SIN_BD=1)
"""

import os, sys, re, shutil, tempfile, threading, datetime, unittest, importlib, warnings
from unittest import mock

warnings.simplefilter("ignore", ResourceWarning)

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import _soporte as SP
from _soporte import (PROY, BDGrabadora, Grabadora, Campo, Dummy, extraer,
                      extraer_asignacion, compilar, js)

VIEJA = SP.importar_app_vieja()            # agp_app.py original
import servicios as S
import agp_web as W
import crear_arte_acad as CAA
import db_app.asignaciones as ASG
import tkinter

APP_PY  = os.path.join(PROY, "agp_app.py")
CAJ_PY  = os.path.join(PROY, "crear_arte_acad.py")
ASG_PY  = os.path.join(PROY, "db_app", "asignaciones.py")
TLL_PY  = os.path.join(PROY, "tab_logos_laser.py")
SIN_BD  = os.environ.get("AGP_SIN_BD") == "1"

# Ninguna prueba debe escribir eventos de tracking reales
EVENTOS_V, EVENTOS_N = Grabadora(), Grabadora()
VIEJA._registrar_evento = EVENTOS_V
W._registrar_evento = EVENTOS_N
for _n in ("_iniciar_sesion", "_marcar_actividad", "_latido", "_cerrar_sesion"):
    setattr(W, _n, lambda *a, **k: None)
    setattr(VIEJA, _n, lambda *a, **k: None)


def _ev(g):
    """Eventos de tracking registrados → lista comparable."""
    return [(a[0] if a else None, k.get("pestana"), k.get("detalle"), k.get("exito")) for a, k in g.llamadas]


def _fake_self(**attrs):
    o = type("FakeSelf", (), {})()
    o.after = lambda ms, fn=None, *a: fn(*a) if fn else None
    for k, v in attrs.items():
        setattr(o, k, v)
    return o


class _Arbol(Dummy):
    def get_children(self): return []


def _api(usuario=None):
    a = W.Api()
    a._usuario = usuario or {"nombre": "Natalia León", "usuario": "nleon@agpglass.com", "rol": "admin"}
    a._sesion_abierta = True
    return a


# ══════════════════════════════════════════════════════════════════════════
#  1. Constantes de negocio idénticas
# ══════════════════════════════════════════════════════════════════════════
class T1_Constantes(unittest.TestCase):
    def test_consultar_queries_y_where(self):
        self.assertEqual(VIEJA.TabBD.QUERIES, S.CONSULTA_QUERIES)
        self.assertEqual(VIEJA.TabBD.WHERE, S.CONSULTA_WHERE)

    def test_gestion_queries_where_editables_tablas_pk(self):
        self.assertEqual(VIEJA.TabGestion.QUERIES, S.GESTION_QUERIES)
        self.assertEqual(VIEJA.TabGestion.WHERE, S.GESTION_WHERE)
        self.assertEqual(VIEJA.TabGestion.EDITABLES, S.EDITABLES)
        self.assertEqual(VIEJA.TabGestion._TABLA, S.TABLA_SQL)
        self.assertEqual(VIEJA.TabGestion._PK, S.PK)

    def test_dropdowns_del_editor(self):
        # agp_app: _DROPDOWNS caso/bnerig + tipo de pasta_plata = RED/ANT
        for tab, campos in S.EDITABLES.items():
            for _, campo in campos:
                vieja = {"caso": ["Caso 1", "Caso 2"], "bnerig": ["BN", "BNI"]}.get(campo)
                if campo == "tipo" and tab == "pasta_plata":
                    vieja = ["RED", "ANT"]
                self.assertEqual(S.opciones_campo(tab, campo), vieja, f"{tab}.{campo}")

    def test_formulario_insertar(self):
        for tabla, campos in VIEJA.TabGestion._FORM_CAMPOS.items():
            nuevos = S.FORM_CAMPOS[tabla]
            self.assertEqual([(l, c) for l, c, _, _ in campos], [(l, c) for l, c, _, _ in nuevos])
            self.assertEqual([ro for *_, ro in campos], [a for _, _, a, _ in nuevos])  # readonly == auto
            for (_, c, _, _), (_, _, _, ops) in zip(campos, nuevos):
                vieja = {"caso": ["Caso 1", "Caso 2"],
                         "tipo": ["RED", "ANT"] if tabla == "pasta_plata" else None}.get(c)
                self.assertEqual(ops, vieja, f"{tabla}.{c}")

    def test_permisos_y_paginas(self):
        self.assertEqual(VIEJA._PERMISOS, W.PERMISOS)
        self.assertEqual([n for n, _, _ in VIEJA.AGPApp.PAGES], [n for _, n in W.PAGINAS])

    def test_tablas_de_estadisticas(self):
        fuente = extraer(APP_PY, "_load_stats", "_worker", clase="TabBD")
        mapa = eval(re.search(r"tabla_map = (\{.*?\})", fuente, re.S).group(1))
        self.assertEqual(mapa, S.STATS_TABLAS)

    def test_conexiones_externas_mismos_datos(self):
        src = open(APP_PY, encoding="utf-8").read()
        nuevo = open(os.path.join(PROY, "db_app", "conexiones_externas.py"), encoding="utf-8").read()
        for clave in ("192.168.2.23", "Comercial", "agpcolsap.database.windows.net", "DB_COL_SAP",
                      "agpcolcalendario.database.windows.net", "CalendarioAGP"):
            self.assertIn(clave, src); self.assertIn(clave, nuevo)


# ══════════════════════════════════════════════════════════════════════════
#  2. Funciones puras de artes (rutas / códigos / búsqueda)
# ══════════════════════════════════════════════════════════════════════════
class T2_RutasArtes(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.mkdtemp(prefix="agp_eq_")

    def tearDown(self):
        shutil.rmtree(self.tmp, ignore_errors=True)

    def test_ruta_arte_salida(self):
        dwg = os.path.join(self.tmp, "CHEVROLET", "1708 008 030.dwg")
        casos = [("", "", ""), ("A-13771", "030", ""), ("A-1/A-2", "030", ""), ("A-1, 35543", "", ""),
                 ("", "030", ""), ("A-1 / 35543 /", "001-002", ""), ("", "", "P A-1 030"),
                 ("", "", "P A-1 030.dwg"), ("", "", 'P: A*1?"x"<y>|z'), ("", "", "P  A--1"),
                 ("A/B", "1:2", ""), ("  ", "  ", ""), ("A-1", "030", "Mi Nombre.DWG")]
        for m, p, n in casos:
            self.assertEqual(VIEJA._ruta_arte_salida(dwg, m, p, n), S.ruta_arte_salida(dwg, m, p, n), (m, p, n))

    def test_extraer_codigos(self):
        for r in ["1708 008 030.dwg", "PLANO.dwg", "AC 1009 900 001-002.dwg", "x123456789.dwg",
                  r"C:\a\12 34 56 78.dwg", "1 2 3 4 5 6 7.dwg", "A-13771_PLANO.dwg", ""]:
            self.assertEqual(VIEJA._extraer_codigos(r), S.extraer_codigos(r), r)

    def test_ruta_planos(self):
        dwg = os.path.join(self.tmp, "X", "p.dwg")
        self.assertEqual(VIEJA._ruta_planos(dwg), S.ruta_planos(dwg))

    def test_buscar_artes(self):
        for rel in [r"V1\ARTES\BN\P A-1 030.dwg", r"V1\ARTES\BN\P 008 030.dwg", r"V2\ARTES\P 1708 008 030.3dm",
                    r"V2\ARTESx\P 008 030.dwg", r"V3\PLANOS\P 008 030.dwg", r"V1\ARTES\BN\X 008 030.dwg",
                    r"V1\ARTES\BN\P 008 030.pdf", r".oculto\ARTES\P 008 030.dwg", r"V4\artes\bn\p 9 008 030.dwg"]:
            ruta = os.path.join(self.tmp, rel)
            os.makedirs(os.path.dirname(ruta), exist_ok=True)
            open(ruta, "w").close()
        for cods in ([], ["008", "030"], ["030"], ["1708", "008", "030"], ["999"]):
            self.assertEqual(VIEJA._buscar_artes(self.tmp, cods), S.buscar_artes(self.tmp, cods), cods)


# ══════════════════════════════════════════════════════════════════════════
#  3. Cajetín: reglas de auto-relleno (Python original vs JS nuevo)
# ══════════════════════════════════════════════════════════════════════════
class T3_Cajetin(unittest.TestCase):
    import re as _re

    def test_parsear_cod_y_anio(self):
        ent = {"_r": re}
        parsear = compilar(extraer(CAJ_PY, "dialogo_cajetin", "parsear_cod"), "parsear_cod", ent)
        anio = compilar(extraer(CAJ_PY, "dialogo_cajetin", "extraer_anio"), "extraer_anio", ent)
        cods = ["1708 008 030", "1708-008-030", " 1708  008 ", "AC 1009 900 001-002", "-12- -3-", "", "abc",
                "P A-13771 030", "1708_PLANO", "12 34 56 78", "--", "1-2-3 4", "M1352 000 000"]
        self.assertEqual([list(parsear(c)) for c in cods],
                         js(["arte.js"], "(c) => ReglasCajetin.parsearCod(c)", [[c] for c in cods]))
        vehs = ["CHEVROLET TRACKER 2021", "2019 MAZDA", "HILUX 20201", "RANGER 1899 2005", "", "TX 2099/2100"]
        self.assertEqual([anio(v) for v in vehs],
                         js(["arte.js"], "(v) => ReglasCajetin.extraerAnio(v)", [[v] for v in vehs]))

    def test_nombre_arte_sugerido(self):
        casos = [("A-13771", "030", "1708 008 030"), ("", "", "1708 008 030"), ("A-1 / 35543", "001-002", "x"),
                 ("A-1,A-2", "", "x"), ("  A-1  ", "  030 ", "x"), ('A:1*"?', "0<3>0|", "x"),
                 ("A-1//A-2", "030", "x"), ("", "030", "PL"), ("A--1", "0  3", "x"), ("/", "", "x/y")]
        fuente = extraer(CAJ_PY, "dialogo_cajetin", "_sugerir_nombre_arte")
        viejos = []
        for malla, pieza, plano in casos:
            entries = {"MALLA": Campo(malla), "PIEZA": Campo(pieza), "NOMBRE ARTE": Campo("")}
            f = compilar(fuente, "_sugerir_nombre_arte",
                         {"entries": entries, "nombre_plano": plano, "tk": tkinter, "_C": CAA._C})
            f()
            viejos.append(entries["NOMBRE ARTE"].valor)
        nuevos = js(["arte.js"], "(m, p, n) => ReglasCajetin.nombreSugerido(m, p, n)", [list(c) for c in casos])
        self.assertEqual(viejos, nuevos)

    def test_campos_y_valores_por_defecto(self):
        hoy = datetime.date.today().strftime("%d.%m.%Y")
        secciones = {}
        exec(extraer_asignacion(CAJ_PY, "dialogo_cajetin", "SECCIONES"), {"hoy": hoy}, secciones)
        viejos = [(c, d) for _, filas in secciones["SECCIONES"] for c, _, d in filas]
        dia = datetime.date.today()
        nuevos = js(["arte.js"], """() => { const d = new Date(%d, %d, %d);
            const hoy = `${String(d.getDate()).padStart(2,"0")}.${String(d.getMonth()+1).padStart(2,"0")}.${d.getFullYear()}`;
            return hoy; }""" % (dia.year, dia.month - 1, dia.day), [[]])
        self.assertEqual(nuevos[0], hoy)
        js_src = open(os.path.join(PROY, "ui_web", "js", "arte.js"), encoding="utf-8").read()
        for campo, defecto in viejos:
            self.assertIn(f'["{campo}"', js_src, campo)
            if defecto and defecto not in ("auto", hoy):
                self.assertIn(f'"{defecto}"', js_src, campo)

    def test_aplicar_cajetin_confirma_igual(self):
        """aceptar() del cajetín Tk vs Api.arte_iniciar: mismos argumentos a confirmar()
        y misma ruta de salida del arte."""
        tmp = tempfile.mkdtemp(prefix="agp_eq_")
        try:
            dwg = os.path.join(tmp, "1708 008 030.dwg"); open(dwg, "w").close()
            reserva = {"vitros": ["T-31149"], "grandes": ["A-13772"], "pequenas": [], "n_vitros": 1,
                       "n_grandes": 1, "n_pequenas": 0}
            for valores in (
                {"DIBUJO": "Natalia", "VEHICULO": "CHEVROLET TRACKER 2021", "MODELO": "2021",
                 "COD PLANO": "1708 008 030", "NAGS": "1708", "VERSION": "008", "PIEZA": "030",
                 "VITRO": "T-31149", "MALLA": "A-13772", "NOMBRE ARTE": "P A-13772 030.dwg",
                 "FECHA": "15.10.2026", "REVISADO": "Santiago P.", "MEDIDAS": "Milimetros",
                 "VISTA": "Interna", "ESCALA": "1:1"},
                {"DIBUJO": "", "VEHICULO": "", "MODELO": "", "COD PLANO": "", "NAGS": "", "VERSION": "",
                 "PIEZA": "", "VITRO": "", "MALLA": "", "NOMBRE ARTE": "", "FECHA": "x",
                 "REVISADO": "", "MEDIDAS": "", "VISTA": "", "ESCALA": ""},
            ):
                # ── viejo: aceptar() literal + dialogo_cajetin(nombre) => ruta_salida="" ──
                conf_v = Grabadora()
                secciones = {}
                exec(extraer_asignacion(CAJ_PY, "dialogo_cajetin", "SECCIONES"), {"hoy": "x"}, secciones)
                entries = {k: Campo(v) for k, v in valores.items()}
                resultado = [None]
                aceptar = compilar(extraer(CAJ_PY, "dialogo_cajetin", "aceptar"), "aceptar", {
                    "entries": entries, "SECCIONES": secciones["SECCIONES"], "resultado": resultado,
                    "_asignacion_cache": [reserva], "nombre_plano": "1708 008 030", "ruta_salida": "",
                    "root": Dummy(), "print": print})
                with mock.patch.object(ASG, "confirmar", conf_v):
                    aceptar()
                vals_v = resultado[0]
                arte_v = VIEJA._ruta_arte_salida(dwg, vals_v.get("MALLA", "").strip(),
                                                 vals_v.get("PIEZA", "").strip(),
                                                 vals_v.get("NOMBRE ARTE", "").strip())
                # ── nuevo ──
                conf_n = Grabadora()
                api = _api()
                api._reservas["rid"] = reserva
                lanzado = Grabadora()
                api._t_crear = lambda *a: lanzado(*a)
                with mock.patch.object(W, "_confirmar", conf_n), \
                     mock.patch.object(W.threading, "Thread", lambda target, args, daemon: type("T", (), {"start": lambda s: target(*args)})()):
                    r = api.arte_iniciar("crear", dwg, valores, "rid", False)
                self.assertTrue(r["ok"], r)
                self.assertEqual(conf_v.llamadas, conf_n.llamadas)
                self.assertEqual(r["arte"], arte_v)
                self.assertEqual(lanzado.llamadas[0][0][1], vals_v)   # mismos valores al pipeline
                self.assertEqual(lanzado.llamadas[0][0][2], arte_v)
        finally:
            shutil.rmtree(tmp, ignore_errors=True)


# ══════════════════════════════════════════════════════════════════════════
#  4. Escrituras en BD (BD grabadora): editar, insertar, anular, roles, separar
# ══════════════════════════════════════════════════════════════════════════
def _responder_editor(viejos, filas_pieza_grandes=1, filas_ruta_grandes=1):
    def r(sql, p):
        m = re.match(r"SELECT (.+) FROM mallas\.\w+ WHERE \w+=%s$", sql)
        if m and "codigo_malla FROM mallas.vitrojet WHERE vitro" not in sql \
                and "vitro FROM mallas.vitrojet WHERE codigo_malla" not in sql:
            campos = [c.strip() for c in m.group(1).split(",")]
            return [tuple(viejos.get(c) for c in campos)], 1
        if "SELECT codigo_malla FROM mallas.vitrojet WHERE vitro" in sql:
            return [("A-555",)], 1
        if "SELECT vitro FROM mallas.vitrojet WHERE codigo_malla" in sql:
            return [("T-777",)], 1
        if sql.startswith("UPDATE mallas.grandes SET pieza"):
            return [], filas_pieza_grandes
        if sql.startswith("UPDATE mallas.grandes SET ruta_dwg"):
            return [], filas_ruta_grandes
        return [], 1
    return r


class T4_Escrituras(unittest.TestCase):
    def _editar_viejo(self, tab, fila, nuevos, responder):
        bd = BDGrabadora(responder)
        pk_col = VIEJA.TabGestion._PK[tab]
        fuente = extraer(APP_PY, "_abrir_editor", "_guardar", clase="TabGestion")
        ent = dict(vars(VIEJA))
        ent.update({"entries_edit": {c: Campo(v) for c, v in nuevos.items()},
                    "self": _fake_self(_usuario={"nombre": "Natalia León"}, _do_search=lambda: None,
                                       _on_data_changed=None),
                    "db_connect": bd, "tabla": VIEJA.TabGestion._TABLA[tab], "pk_col": pk_col,
                    "pk_val": fila[pk_col], "msg_lbl": Dummy(), "win": Dummy(),
                    "_registrar_evento": lambda *a, **k: None})
        compilar(fuente, "_guardar", ent)()
        return bd.log

    def _editar_nuevo(self, tab, fila, nuevos, responder):
        bd = BDGrabadora(responder)
        with mock.patch.object(S, "db_connect", bd):
            r = _api({"nombre": "Natalia León"}).gestion_guardar(tab, S._txt(fila[S.PK[tab]]), nuevos)
        self.assertTrue(r["ok"], r)
        return bd.log

    def test_editar_todas_las_tablas(self):
        escenarios = {
            "vitrojet": ({"vitro": "T-100"}, {"codigo_malla": "A-555", "bnerig": "BNI", "vehiculo": "HILUX 2020",
                         "version": "", "_pieza_malla": "030", "ruta": r"\\srv\a.dwg", "responsable": "  Nati  "}),
            "grandes": ({"codigo": "A-555"}, {"cod_veh": "1708", "descripcion": "TRACKER", "pieza": "030",
                        "tipo": "S", "version": "008", "ruta_dwg": r"\\srv\b.dwg", "responsable": ""}),
            "pequenas": ({"codigo": 35543}, {"cod_veh": "1708", "descripcion": "X", "pieza": "001",
                         "tipo": "", "version": "1", "ruta_dwg": r"\\srv\c.dwg", "responsable": "Ana"}),
            "vinilos": ({"herramental": "VC-0001"}, {"vehiculo": "A", "cod_vehiculo": "1", "version": "2",
                        "pieza": "3", "tipo": "4", "ruta": "r"}),
            "pasta_plata": ({"consecutivo": "S-00001"}, {"tipo": "ANT", "vehiculo": "A", "cod_vehiculo": "",
                            "version": "2", "pieza": "3", "ruta_archivo": "r", "caso": "Caso 2"}),
        }
        for tab, (fila, nuevos) in escenarios.items():
            for viejos in ({}, {k: v for k, v in nuevos.items()}, {k: "otro" for k in nuevos}):
                for fp, fr in ((1, 1), (0, 0)):    # con y sin fallback a pequenas
                    resp = _responder_editor(viejos, fp, fr)
                    with self.subTest(tab=tab, viejos=bool(viejos), fallback=fp):
                        self.assertEqual(self._editar_viejo(tab, fila, nuevos, resp),
                                         self._editar_nuevo(tab, fila, nuevos, resp))

    def test_editar_ruta_vacia_no_cascadea(self):
        resp = _responder_editor({})
        nuevos = {"codigo_malla": "", "bnerig": "BN", "vehiculo": "", "version": "",
                  "_pieza_malla": "", "ruta": "", "responsable": ""}
        self.assertEqual(self._editar_viejo("vitrojet", {"vitro": "T-1"}, nuevos, resp),
                         self._editar_nuevo("vitrojet", {"vitro": "T-1"}, nuevos, resp))

    def test_insertar_manual(self):
        for tabla, valores in (
            ("pasta_plata", {"tipo": "RED", "vehiculo": "HILUX", "cod_vehiculo": "1708", "version": "008",
                             "pieza": "030", "ruta_archivo": r"\\srv\x.pdf", "caso": "Caso 1"}),
            ("pasta_plata", {"tipo": "ANT", "vehiculo": "", "cod_vehiculo": " 1 ", "version": "",
                             "pieza": "", "ruta_archivo": "", "caso": "Caso 2"}),
            ("vinilos", {"vehiculo": "TRACKER", "cod_vehiculo": "1708", "version": "", "pieza": "030", "tipo": "X"}),
        ):
            resp = lambda sql, p: (([("S-01093",)] if sql.startswith("SELECT TOP 1") else []), 1)
            # viejo
            bd_v = BDGrabadora(resp)
            campos = VIEJA.TabGestion._FORM_CAMPOS[tabla]
            pk = {"pasta_plata": "consecutivo", "vinilos": "herramental"}[tabla]
            entries = {c: Campo(valores.get(c, "AUTO")) for _, c, _, _ in campos}
            ent = dict(vars(VIEJA))
            ent.update({"tabla": tabla, "campos": campos, "entries_ins": entries, "msg_lbl": Dummy(),
                        "db_connect": bd_v, "tabla_sql": VIEJA.TabGestion._FORM_TITULO[tabla][2],
                        "_tk": tkinter, "win": Dummy(), "self": _fake_self(_do_search=lambda: None,
                        _on_data_changed=None), "_registrar_evento": lambda *a, **k: None})
            compilar(extraer(APP_PY, "_form_insertar", "_insertar", clase="TabGestion"), "_insertar", ent)()
            # nuevo
            bd_n = BDGrabadora(resp)
            with mock.patch.object(S, "db_connect", bd_n):
                r = _api().gestion_insertar(tabla, valores)
            with self.subTest(tabla=tabla):
                self.assertTrue(r["ok"], r)
                self.assertEqual(bd_v.log, bd_n.log)
                self.assertEqual(r["pk"], "S-01093")

    def test_insertar_vacio_se_rechaza_igual(self):
        bd_v = BDGrabadora()
        campos = VIEJA.TabGestion._FORM_CAMPOS["vinilos"]
        ent = dict(vars(VIEJA))
        msg = Dummy()
        ent.update({"tabla": "vinilos", "campos": campos, "entries_ins": {c: Campo("") for _, c, _, _ in campos},
                    "msg_lbl": msg, "db_connect": bd_v, "tabla_sql": "mallas.vinilos", "_tk": tkinter,
                    "win": Dummy(), "self": _fake_self(), "_registrar_evento": lambda *a, **k: None})
        compilar(extraer(APP_PY, "_form_insertar", "_insertar", clase="TabGestion"), "_insertar", ent)()
        self.assertEqual(bd_v.log, [])
        self.assertIn("Completa al menos un campo", msg.texto)
        bd_n = BDGrabadora()
        with mock.patch.object(S, "db_connect", bd_n):
            r = _api().gestion_insertar("vinilos", {})
        self.assertFalse(r["ok"]); self.assertIn("Completa al menos un campo", r["error"])
        self.assertEqual(bd_n.log, [])

    def test_anular(self):
        import ctypes
        for tab, pk in (("vitrojet", "T-1"), ("grandes", "A-1"), ("pequenas", "35543")):
            g_v, g_n = Grabadora({"vitrojet": 1}), Grabadora({"vitrojet": 1})
            ent = dict(vars(VIEJA))
            ent.update({"pk_col": S.PK[tab], "pk_val": pk, "tab": tab, "_anular_asignacion": g_v,
                        "msg_lbl": Dummy(), "win": Dummy(), "self": _fake_self(_do_search=lambda: None,
                        _on_data_changed=None), "_registrar_evento": lambda *a, **k: None})
            with mock.patch.object(ctypes.windll.user32, "MessageBoxW", lambda *a: 6, create=True):
                compilar(extraer(APP_PY, "_abrir_editor", "_anular", clase="TabGestion"), "_anular", ent)()
            with mock.patch.object(W, "_anular_asignacion", g_n):
                _api().gestion_anular(tab, pk)
            self.assertEqual(g_v.llamadas, g_n.llamadas, tab)

    def test_roles_guardar(self):
        for rol, activo in (("admin", True), ("— sin rol —", False), ("planta", True)):
            bd_v, bd_n = BDGrabadora(), BDGrabadora()
            yo = _fake_self(_sel_id=7, _combo_rol=Campo(rol), _switch_est=Campo(1 if activo else 0),
                            _lbl_msg=Dummy(), refresh=lambda: None, _lbl_usuario=Dummy(text="u@agp"))
            with mock.patch.object(VIEJA, "db_connect", bd_v):
                VIEJA.TabRoles._guardar(yo)
            with mock.patch.object(S, "db_connect", bd_n):
                _api().roles_guardar(7, "" if rol.startswith("—") else rol, activo, "u@agp")
            self.assertEqual(bd_v.log, bd_n.log, rol)

    def test_separar_confirmar(self):
        reserva = {"vitros": ["T-1"], "grandes": [], "pequenas": ["35543"]}
        for datos in ({"vehiculo": "HILUX", "cod_vehiculo": "1708", "version": "008", "pieza": "030",
                       "cod_completo": "1708 008 030", "bnerig": "BNI", "ruta": r"\\s\a.dwg", "responsable": "Ana"},
                      {"vehiculo": " ", "cod_vehiculo": "", "version": "", "pieza": "", "cod_completo": "",
                       "bnerig": "BN", "ruta": "", "responsable": ""}):
            g_v = Grabadora(reserva)
            ent = {"reserva_actual": [reserva], "messagebox": Dummy(), "_do_confirmar": g_v,
                   "v_veh": Campo(datos["vehiculo"]), "v_ver": Campo(datos["version"]),
                   "v_piez": Campo(datos["pieza"]), "v_cod": Campo(datos["cod_vehiculo"]),
                   "v_comp": Campo(datos["cod_completo"]), "v_bn": Campo(datos["bnerig"]),
                   "ruta_var": Campo(datos["ruta"]), "v_resp": Campo(datos["responsable"]),
                   "resultado": [None], "confirmado": [False], "lbl_reserva_estado": Dummy(),
                   "C": {"green": "", "panel": "", "muted": "", "sep": ""}, "b_confirmar": Dummy(),
                   "b_reservar": Dummy(), "b_cerrar": Dummy(), "lbl_sub_hdr": Dummy()}
            compilar(extraer(ASG_PY, "dialogo_separar", "_hacer_confirmar"), "_hacer_confirmar", ent)()
            g_n = Grabadora(reserva)
            api = _api(); api._reservas["r"] = reserva
            with mock.patch.object(W, "_confirmar", g_n):
                self.assertTrue(api.confirmar_separar("r", datos)["ok"])
            self.assertEqual(g_v.llamadas, g_n.llamadas)

    def test_reservar_y_cancelar(self):
        res = {"vitros": ["T-9"], "grandes": ["A-9"], "pequenas": [], "n_vitros": 1, "n_grandes": 1, "n_pequenas": 0}
        g_r, g_c = Grabadora(res), Grabadora()
        api = _api()
        with mock.patch.object(W, "_reservar", g_r), mock.patch.object(W, "_cancelar", g_c):
            r1 = api.reservar(1, 1, 0)
            r2 = api.reservar(2, 0, 1, r1["rid"])            # re-reservar libera la anterior
            api.cancelar_reserva(r2["rid"])
            self.assertFalse(api.reservar(0, 0, 0)["ok"])    # mismo aviso que el diálogo viejo
        self.assertEqual(g_r.llamadas, [((1, 1, 0), {}), ((2, 0, 1), {})])
        self.assertEqual(g_c.llamadas, [((res,), {}), ((res,), {})])

    def test_cerrar_ventana_libera_reservas_abiertas(self):
        g_c = Grabadora()
        api = _api(); api._reservas = {"a": {"vitros": ["T-1"]}}
        with mock.patch.object(W, "_cancelar", g_c):
            api.al_cerrar()
        self.assertEqual(g_c.llamadas, [(({"vitros": ["T-1"]},), {})])


# ══════════════════════════════════════════════════════════════════════════
#  5. Pipeline AutoCAD: misma orquestación (AutoCAD / BD falsos)
# ══════════════════════════════════════════════════════════════════════════
class _Motor:
    llamadas = []
    def __init__(self): _Motor.llamadas.append(("init",))
    def extraer_layers(self, dwg, dest, log_fn=None):
        _Motor.llamadas.append(("extraer", dwg, dest)); log_fn and log_fn("capas ok")
    def quit(self): _Motor.llamadas.append(("quit",))


class T5_Pipeline(unittest.TestCase):
    VALORES = {"VITRO": "T-1 / T-2", "MALLA": "A-1 / 35543", "PIEZA": "030", "NOMBRE ARTE": "",
               "DIBUJO": "Nati"}

    def _correr(self, vieja, metodo, crear_fn, args, extra_patches=(), yo_extra=None):
        """Corre _t_crear / _t_todo / _t_extraer viejo o nuevo con todo AutoCAD falso.
        Devuelve (log, llamadas a AutoCAD/BD, alertas, eventos)."""
        logs, alertas = [], []
        crear = Grabadora(crear_fn)
        ruta, anular, popen, copia = Grabadora(), Grabadora({"vitrojet": 1}), Grabadora(), Grabadora()
        _Motor.llamadas = []
        acad = mock.MagicMock()
        comunes = [mock.patch("time.sleep", lambda *_: None),
                   mock.patch("subprocess.Popen", popen), mock.patch("shutil.copy2", copia),
                   mock.patch("pythoncom.CoInitialize", lambda: None),
                   mock.patch("pythoncom.CoUninitialize", lambda: None),
                   mock.patch("win32com.client.GetActiveObject", lambda *_: acad)] + list(extra_patches)
        if vieja:
            EVENTOS_V.llamadas.clear()
            yo = _fake_self(_log_fn=lambda m, t="": logs.append((m, t)), _compensar=Campo(True),
                            _busy=lambda on: None, _on_art_done=lambda: logs.append(("<art_done>", "")),
                            **(yo_extra or {}))
            parches = [mock.patch.object(VIEJA, "_crear_arte_autocad", crear),
                       mock.patch.object(VIEJA, "_actualizar_ruta_arte", ruta),
                       mock.patch.object(VIEJA, "_anular_asignacion", anular),
                       mock.patch.object(VIEJA, "AutoCADMotor", _Motor),
                       mock.patch.object(VIEJA, "_msgbox_topmost", lambda *a: alertas.append(a))]
            with _juntos(comunes + parches):
                getattr(VIEJA.TabArte, metodo)(yo, *args)
            ev = _ev(EVENTOS_V)
        else:
            EVENTOS_N.llamadas.clear()
            api = _api(); api._busy["arte"] = True
            parches = [mock.patch.object(S, "crear_arte_autocad", crear),
                       mock.patch.object(W, "_actualizar_ruta_arte", ruta),
                       mock.patch.object(S, "_anular_asignacion", anular),
                       mock.patch.object(W, "AutoCADMotor", _Motor)]
            with _juntos(comunes + parches):
                getattr(api, metodo)(*args)
            while not api._logs.empty():
                l = api._logs.get(); logs.append((l["msg"], l["tag"]))
            while not api._eventos.empty():
                e = api._eventos.get()
                if e["tipo"] == "arte_done": logs.append(("<art_done>", ""))
                if e["tipo"] == "alerta": alertas.append((e["nivel"], e["titulo"], e["mensaje"]))
            self.assertFalse(api._busy["arte"], "el flag busy debe liberarse")
            ev = _ev(EVENTOS_N)
        # el traceback trae rutas/números de línea distintos: se compara solo su primera línea
        logs = [(m.splitlines()[0] if m.startswith("Traceback") else m, t) for m, t in logs]
        llamadas = {"crear": [(a, {k: v for k, v in kw.items() if k != "log_fn"}) for a, kw in crear.llamadas],
                    "ruta": ruta.llamadas, "anular": anular.llamadas, "popen": popen.llamadas,
                    "copia": copia.llamadas, "motor": _Motor.llamadas}
        return logs, llamadas, alertas, ev

    def _comparar(self, metodo, crear_fn, args_v, args_n):
        v = self._correr(True, metodo, crear_fn, args_v)
        n = self._correr(False, metodo, crear_fn, args_n)
        for nombre, a, b in zip(("log", "llamadas", "alertas", "tracking"), v, n):
            self.assertEqual(a, b, f"{metodo}: difiere {nombre}")

    def setUp(self):
        self.tmp = tempfile.mkdtemp(prefix="agp_eq_")
        self.dwg = os.path.join(self.tmp, "1708 008 030.dwg"); open(self.dwg, "w").close()
        self.arte = S.ruta_arte_salida(self.dwg, "A-1 / 35543", "030")

    def tearDown(self):
        shutil.rmtree(self.tmp, ignore_errors=True)

    def test_crear_ok(self):
        self._comparar("_t_crear", lambda *a, **k: 1, (self.dwg, self.VALORES, self.arte),
                       (self.dwg, self.VALORES, self.arte, True))

    def test_crear_error_anula_en_bd(self):
        def falla(*a, **k): raise RuntimeError("AutoCAD no está abierto")
        self._comparar("_t_crear", falla, (self.dwg, self.VALORES, self.arte), (self.dwg, self.VALORES, self.arte, True))

    def test_crear_error_al_guardar_no_anula(self):
        def falla(*a, **k): raise CAA.ErrorGuardadoArte("archivo bloqueado")
        with mock.patch.object(VIEJA, "ErrorGuardadoArte", CAA.ErrorGuardadoArte):
            self._comparar("_t_crear", falla, (self.dwg, self.VALORES, self.arte),
                           (self.dwg, self.VALORES, self.arte, True))

    def test_todo_en_uno_varias_piezas(self):
        self._comparar("_t_todo", lambda *a, **k: 3, (self.dwg, self.VALORES, self.arte),
                       (self.dwg, self.VALORES, self.arte, True))

    def test_todo_en_uno_error(self):
        def falla(*a, **k): raise RuntimeError("se cayó COM")
        self._comparar("_t_todo", falla, (self.dwg, self.VALORES, self.arte), (self.dwg, self.VALORES, self.arte, True))

    def test_extraer(self):
        for motor_ok in (True, False):
            with mock.patch.object(VIEJA, "_MOTOR_OK", motor_ok), mock.patch.object(W, "_MOTOR_OK", motor_ok):
                v = self._correr(True, "_t_extraer", None, (), yo_extra={"_ruta_dwg": Campo(self.dwg)})
                n = self._correr(False, "_t_extraer", None, (self.dwg,))
            self.assertEqual(v, n, f"motor_ok={motor_ok}")

    def test_buscar(self):
        for rel in (r"V1\ARTES\BN\P A-1 008 030.dwg", r"V2\ARTES\P 1708 008 030.dwg"):
            p = os.path.join(self.tmp, rel); os.makedirs(os.path.dirname(p), exist_ok=True); open(p, "w").close()
        capt = {}
        yo = _fake_self(_ruta_base=Campo(self.tmp), _ruta_dwg=Campo(self.dwg),
                        _log_fn=lambda *a: None, _busy=lambda on: None,
                        _fill_table=lambda res: capt.setdefault("res", res))
        VIEJA.TabArte._t_buscar(yo)
        r = _api().arte_buscar(self.tmp, self.dwg)
        self.assertEqual(capt["res"], r["resultados"])

    def test_superponer_mismos_argumentos(self):
        import arte_maker
        g_v, g_n = Grabadora(), Grabadora()
        with mock.patch.object(arte_maker, "_overlay_autocad", g_v), \
             mock.patch("pythoncom.CoInitialize", lambda: None), mock.patch("pythoncom.CoUninitialize", lambda: None):
            VIEJA.TabArte._t_overlay(_fake_self(_log_fn=lambda *a: None, _busy=lambda on: None), "arte.dwg", self.dwg)
        with mock.patch.object(arte_maker, "_overlay_autocad", g_n), \
             mock.patch("pythoncom.CoInitialize", lambda: None), mock.patch("pythoncom.CoUninitialize", lambda: None), \
             mock.patch.object(W.threading, "Thread", lambda target, daemon: type("T", (), {"start": lambda s: target()})()):
            _api().arte_superponer("arte.dwg", self.dwg)
        self.assertEqual([a for a, _ in g_v.llamadas], [a for a, _ in g_n.llamadas])


class _juntos:
    def __init__(self, parches): self.p = parches
    def __enter__(self):
        for p in self.p: p.__enter__()
    def __exit__(self, *e):
        for p in reversed(self.p): p.__exit__(*e)


# ══════════════════════════════════════════════════════════════════════════
#  6. Logos Laser: reglas de la grilla (excel_grid.py vs logos.js)
# ══════════════════════════════════════════════════════════════════════════
class T6_LogosLaser(unittest.TestCase):
    COLS = [{"key": "id", "type": "text", "editable": False}, {"key": "cliente", "type": "text"},
            {"key": "producto", "type": "text"}, {"key": "dxf", "type": "bool"}, {"key": "sjf", "type": "bool"},
            {"key": "validado", "type": "bool"}, {"key": "cargado_genesis", "type": "bool"},
            {"key": "azure", "type": "bool"}, {"key": "app", "type": "bool"}, {"key": "observaciones", "type": "text"}]
    FILAS = [{"id": "ML-0003", "cliente": "WBA", "producto": "B33", "dxf": 1, "sjf": 0, "validado": 0,
              "cargado_genesis": 0, "azure": 1, "app": 0, "observaciones": None},
             {"id": "ML-0002", "cliente": "BMW", "producto": "X5\tG05", "dxf": 0, "sjf": 1, "validado": 1,
              "cargado_genesis": 0, "azure": 0, "app": 0, "observaciones": "nota"},
             {"id": "ML-0001", "cliente": None, "producto": "", "dxf": 0, "sjf": 0, "validado": 0,
              "cargado_genesis": 1, "azure": 0, "app": 1, "observaciones": "x"}]

    def _grid_viejo(self, activo, ancla=None):
        import excel_grid, copy
        g = object.__new__(excel_grid.ExcelGrid)
        g._columns = self.COLS
        g._rows = copy.deepcopy(self.FILAS)
        g._active = activo; g._anchor = ancla or activo
        g._editing = None
        g.ediciones = []
        g._on_edit = lambda row, key, val: g.ediciones.append((row["id"], key, val)) or True
        for n in ("_toast", "_redraw", "_ensure_visible"):
            setattr(g, n, lambda *a, **k: None)
        g._canvas_body = Dummy()
        return g

    def test_pegar(self):
        textos = ["hola\tmundo", "1\t0\tx\r\nsi\tno\tSÍ\n", "a\tb\tc\td\te\tf\tg\th\ti\tj\tk", "  ok  ",
                  "WBA\tB33", "", "\n", "TRUE\tyes\ton\t✓\tverdadero", "1\n2\n3\n4\n5", "\t\t1"]
        inicios = [(0, 1), (1, 3), (0, 0), (2, 9), (0, 1), (1, 1), (0, 1), (0, 3), (1, 1), (0, 1)]
        for texto, (r0, c0) in zip(textos, inicios):
            g = self._grid_viejo((r0, c0))
            g.clipboard_get = lambda t=texto: t
            g._on_paste(None)
            nuevo = js(["logos.js"], """(texto, filas, cols, r0, c0) => {
                const out = [];
                for (const [r, c, v] of ReglasGrid.pegar(texto, filas.length, cols, r0, c0)) {
                  const k = cols[c].key, viejo = filas[r][k];
                  if (ReglasGrid.sinCambio(viejo, v)) continue;
                  out.push([filas[r].id, k, v]); filas[r][k] = v;
                }
                return out; }""", [[texto, self.FILAS, self.COLS, r0, c0]])[0]
            with self.subTest(texto=texto):
                self.assertEqual([list(e) for e in g.ediciones], nuevo)

    def test_copiar(self):
        for (r0, c0), (r1, c1) in [((0, 0), (2, 9)), ((1, 2), (1, 2)), ((0, 3), (2, 8)), ((2, 9), (0, 1))]:
            g = self._grid_viejo((r1, c1), (r0, c0))
            capt = []
            g.clipboard_clear = lambda: None
            g.clipboard_append = capt.append
            g._on_copy(None)
            a, b = sorted((r0, r1)), sorted((c0, c1))
            nuevo = js(["logos.js"], "(f, cols, a, b, c, d) => ReglasGrid.copiar(f, cols, a, b, c, d)",
                       [[self.FILAS, self.COLS, a[0], a[1], b[0], b[1]]])[0]
            self.assertEqual(capt[0], nuevo)

    def test_tab_avanza_igual(self):
        for (r, c), paso in [((0, 0), 1), ((0, 9), 1), ((2, 9), 1), ((1, 0), -1), ((0, 0), -1), ((1, 4), -1)]:
            g = self._grid_viejo((r, c))
            g._advance(paso)
            nuevo = js(["logos.js"], "(r, c, p) => ReglasGrid.avanzar(r, c, p, 3, 10)", [[r, c, paso]])[0]
            self.assertEqual(list(g._active), nuevo, (r, c, paso))

    def test_supr_y_check(self):
        g = self._grid_viejo((2, 9), (0, 0))
        g._iter_selected_cells = lambda: [(r, c) for r in range(3) for c in range(10)]
        g._on_delete_key(None)
        esperado = js(["logos.js"], """(filas, cols) => { const out = [];
            for (let r = 0; r < 3; r++) for (let c = 0; c < 10; c++) {
              if (cols[c].editable === false) continue;
              const v = cols[c].type === "bool" ? false : "";
              if (!ReglasGrid.sinCambio(filas[r][cols[c].key], v)) out.push([filas[r].id, cols[c].key, v]); }
            return out; }""", [[self.FILAS, self.COLS]])[0]
        self.assertEqual([list(e) for e in g.ediciones], esperado)
        g = self._grid_viejo((0, 3))
        g._toggle_bool(0, 3); g._toggle_bool(0, 4)
        self.assertEqual(g.ediciones, [("ML-0003", "dxf", False), ("ML-0003", "sjf", True)])
        # logos.js: guardar(r, c, !valor) → mismo booleano
        self.assertEqual(js(["logos.js"], "(f) => [!f[0].dxf, !f[0].sjf]", [[self.FILAS]])[0], [False, True])

    def test_filtro(self):
        filas = self.FILAS + [{"id": "ML-0004", "cliente": "audi", "producto": "Q7", "observaciones": "wba"}]
        for q in ["", "wba", "BMW", "ml-000", "x5", "nota", "  q7 ", "zzz"]:
            capt = {}
            yo = _fake_self(_buscador=Campo(q), _todas=filas,
                            _grid=type("G", (), {"set_rows": lambda s, v: capt.setdefault("v", v)})())
            import tab_logos_laser
            tab_logos_laser.TabLogosLaser._aplicar_filtro(yo)
            nuevo = js(["logos.js"], "(f, q) => ReglasGrid.filtrar(f, q).map(x => x.id)", [[filas, q]])[0]
            self.assertEqual([f["id"] for f in capt["v"]], nuevo, q)

    def test_editar_y_crear_llaman_igual_a_la_bd(self):
        import tab_logos_laser
        g_v, g_n = Grabadora(True), Grabadora(True)
        yo = _fake_self(_status=lambda *a, **k: None, _usuario={"nombre": "Natalia León"})
        with mock.patch.object(tab_logos_laser, "actualizar_campo", g_v), \
             mock.patch.object(tab_logos_laser, "_registrar_evento", lambda *a, **k: None):
            for fila, k, v in (({"id": "ML-1"}, "cliente", "  WBA "), ({"id": "ML-1"}, "dxf", True),
                               ({"id": "ML-2"}, "observaciones", "")):
                tab_logos_laser.TabLogosLaser._on_edit(yo, fila, k, v)
        with mock.patch.object(W, "_ll_actualizar", g_n):
            api = _api()
            for i, k, v in (("ML-1", "cliente", "  WBA "), ("ML-1", "dxf", True), ("ML-2", "observaciones", "")):
                api.ll_actualizar(i, k, v)
        self.assertEqual(g_v.llamadas, g_n.llamadas)
        c_v, c_n = Grabadora({"id": "ML-9"}), Grabadora({"id": "ML-9"})
        yo = _fake_self(_usuario={"nombre": "Natalia León"}, _todas=[], _buscador=Dummy(),
                        _grid=Dummy(), _status=lambda *a, **k: None)
        with mock.patch.object(tab_logos_laser, "crear_fila", c_v), \
             mock.patch.object(tab_logos_laser, "_registrar_evento", lambda *a, **k: None):
            tab_logos_laser.TabLogosLaser._nueva_fila(yo)
        with mock.patch.object(W, "_ll_crear", c_n):
            _api({"nombre": "Natalia León"}).ll_crear()
        self.assertEqual(c_v.llamadas, c_n.llamadas)


# ══════════════════════════════════════════════════════════════════════════
#  7. Login / SSO / sync Excel
# ══════════════════════════════════════════════════════════════════════════
class T7_LoginSync(unittest.TestCase):
    def test_login_mismas_llamadas(self):
        import db_app.auth as AUTH
        info = {"nombre": "Ana Pérez", "usuario": "ana@agp", "rol": "dibujante", "es_admin": False, "estatus": 1}
        for u, p, devuelve in ((" ana@agp ", " 123 ", info), ("ana@agp", "mal", None)):
            v, s = Grabadora(devuelve), Grabadora()
            with mock.patch.object(AUTH, "validar_login", v), mock.patch.object(AUTH, "sincronizar_background", s):
                r = W.Api().login(u, p)
            self.assertEqual(v.llamadas, [((u.strip(), p.strip()), {})])
            self.assertEqual(len(s.llamadas), 1 if devuelve else 0)       # login_window: sync solo si entra
            self.assertEqual(r["ok"], bool(devuelve))
        self.assertFalse(W.Api().login("", "x")["ok"])                     # "Completa todos los campos"

    def test_sso_carga_usuario_fresco_de_bd(self):
        import agp_sso, db_app.auth as AUTH
        token = agp_sso.crear_token("ana@agp", "Ana", "admin")
        carga = Grabadora({"nombre": "Ana", "usuario": "ana@agp", "rol": "planta"})
        with mock.patch.object(sys, "argv", ["x", f"--sso-token={token}"]), \
             mock.patch.object(AUTH, "cargar_por_usuario", carga), \
             mock.patch.object(W.Api, "_post_login", lambda s: None), \
             mock.patch.object(W.Api, "_crear_tabla_usuarios", lambda s: None):
            r = W.Api().inicio()
        self.assertEqual(carga.llamadas, [(("ana@agp",), {})])
        self.assertEqual(r["usuario"]["rol"], "planta")                    # rol de la BD, no del token
        self.assertEqual([p["id"] for p in r["usuario"]["paginas"]], ["scanner"])

    def test_paginas_por_rol(self):
        for rol in ("admin", "dibujante", "planta", None):
            api = _api({"nombre": "x", "rol": rol})
            nuevas = [p["nombre"] for p in api._usuario_publico()["paginas"]]
            viejas = [n for n, _, _ in VIEJA.AGPApp.PAGES if (rol or "") in VIEJA._PERMISOS.get(n, set())]
            self.assertEqual(nuevas, viejas, rol)

    def test_sync_excel_mismo_llamado(self):
        import db_app.importar_excel as IMP
        g_v, g_n = Grabadora(True), Grabadora(True)
        yo = _fake_self(_sync_log=Dummy(), _sync_done=lambda ok: None)
        with mock.patch.object(IMP, "main", g_v), mock.patch("importlib.reload", lambda m: m):
            VIEJA.TabBD._t_sync(yo)
        with mock.patch.object(IMP, "main", g_n), mock.patch("importlib.reload", lambda m: m), \
             mock.patch.object(W.threading, "Thread", lambda target, daemon: type("T", (), {"start": lambda s: target()})()):
            _api().sync_excel()
        self.assertEqual([sorted(k) for _, k in g_v.llamadas], [sorted(k) for _, k in g_n.llamadas])


# ══════════════════════════════════════════════════════════════════════════
#  8. Lecturas contra Azure REAL (solo SELECT): mismas filas vieja vs nueva
# ══════════════════════════════════════════════════════════════════════════
@unittest.skipIf(SIN_BD, "AGP_SIN_BD=1")
class T8_BDReal(unittest.TestCase):
    BUSQUEDAS = ["", "T-30", "CHEVROLET", "A-1", "123", "toyota", "S-00", "VC-0", "zzzz_nada"]

    def test_consultar_bd_misma_tabla_que_antes(self):
        CLASE = VIEJA.TabBD
        for tab in S.CONSULTA_QUERIES:
            for q in self.BUSQUEDAS:
                capt = {}
                yo = _fake_self(_search=Campo(q), _tab=tab, _tree=_Arbol(), _lbl_count=Dummy(), QUERIES=CLASE.QUERIES, WHERE=CLASE.WHERE,
                                _fill=lambda rows, fields, headers: capt.update(rows=rows, fields=fields))
                VIEJA.TabBD._t_search(yo)
                viejo = [[str(r.get(f, "") or "—").replace("\r\n", " ").replace("\n", " ").replace("\r", " ")
                          for f in capt["fields"]] for r in capt["rows"]]
                # la UI nueva pinta "—" en las celdas vacías (igual que la vieja)
                nuevo = [[v if v != "" else "—" for v in fila] for fila in S.consultar(tab, q)["rows"]]
                with self.subTest(tab=tab, q=q):
                    self.assertEqual(viejo, nuevo)

    def test_gestion_misma_tabla_que_antes(self):
        CLASE = VIEJA.TabGestion
        for tab in S.GESTION_QUERIES:
            for q in self.BUSQUEDAS:
                capt = {}
                yo = _fake_self(_search=Campo(q), _tab=tab, _tree=_Arbol(), _lbl_count=Dummy(), QUERIES=CLASE.QUERIES, WHERE=CLASE.WHERE,
                                _fill=lambda rows, fields, headers: capt.update(rows=rows, fields=fields))
                VIEJA.TabGestion._t_search(yo)
                viejo = [[(v.strftime("%Y-%m-%d %H:%M") if hasattr(v, "strftime") else
                           str(v).replace("\r\n", " ").replace("\n", " ").replace("\r", " ")) if v else ""
                          for v in (r.get(f) for f in capt["fields"])] for r in capt["rows"]]
                r = S.gestion_buscar(tab, q)
                nuevo = [[fila[f] for f in r["fields"]] for fila in r["rows"]]
                with self.subTest(tab=tab, q=q):
                    self.assertEqual(viejo, nuevo)

    def test_estadisticas(self):
        capt = {}
        cards = {k: type("C", (), {"set": lambda s, v, k=k: capt.__setitem__(k, v)})() for k in S.STATS_TABLAS}
        yo = _fake_self(_cards=cards)
        with mock.patch.object(VIEJA.threading, "Thread",
                               lambda target, daemon: type("T", (), {"start": lambda s: target()})()):
            VIEJA.TabBD._load_stats(yo)
        self.assertEqual(capt, S.estadisticas())

    def test_scanner_mismo_resultado(self):
        from db_app.conexiones_externas import conectar_comercial
        cn = conectar_comercial(); cur = cn.cursor()
        cur.execute("SELECT TOP 40 ORDEN FROM [dbo].[BI_TAB_FIN_TURNO_SAP] WHERE ZFER IS NOT NULL "
                    "GROUP BY ORDEN ORDER BY ORDEN DESC")
        ordenes = [str(r[0]) for r in cur.fetchall()][::4] + ["999999999999"]
        cn.close()
        for orden in ordenes:
            capt = {}
            yo = _fake_self(_mostrar_resultado=lambda *a: capt.update(res=a),
                            _mostrar_error=lambda m: capt.update(err=m))
            VIEJA.TabScanner._t_buscar(yo, orden)
            nuevo = S.buscar_orden(orden)
            with self.subTest(orden=orden):
                if "err" in capt:
                    self.assertFalse(nuevo["ok"]); continue
                (o, zfer, vitro1, vitro2, mallas, vitro_ruta, malla_rutas, ruta_err, cal, sap) = capt["res"]
                self.assertEqual((o, zfer, vitro1, vitro2 or "—", vitro_ruta),
                                 (nuevo["orden"], nuevo["zfer"], nuevo["vitro1"], nuevo["vitro"], nuevo["vitro_ruta"]))
                self.assertEqual(mallas, [m["codigo"] for m in nuevo["mallas"]])
                self.assertEqual([malla_rutas[m] for m in mallas], [m["ruta"] for m in nuevo["mallas"]])
                for k_v, k_n in (("vehiculo", "vehiculo"), ("lote", "lote"), ("version", "version"),
                                 ("cod_veh", "cod_veh"), ("logo", "logo")):
                    self.assertEqual(cal.get(k_v) or ("" if k_n == "logo" else "—"), nuevo[k_n] , k_v)
                self.assertEqual(sap.get("Z_PIECE_TYPE") or "—", nuevo["tipo_pieza"])
                self.assertEqual(sap.get("Z_TRAZABILITY") or "—", nuevo["trazabilidad"])
                # Lógica de pantalla vieja (misma carpeta / alerta) ejecutada literal
                SP.CREADOS.clear()
                pant = _fake_self(_mallas_labels=[], **{n: Dummy() for n in (
                    "_prog", "_entry", "_c_orden", "_c_zfer", "_lbl_logo", "_lbl_vehiculo", "_lbl_version",
                    "_lbl_lote", "_lbl_version2", "_lbl_sap_codveh", "_lbl_sap_tipo", "_lbl_trazabilidad",
                    "_lbl_mallas_count", "_lbl_vitro2", "_lbl_vitro_ruta", "_mallas_box", "_alerta_frame",
                    "_alerta_lbl", "_zona")})
                pant._copiar = lambda t: None
                VIEJA.TabScanner._mostrar_resultado(pant, *capt["res"])
                self.assertEqual(pant._alerta_txt, nuevo["alerta"])
                textos = [k.get("text") for c, k in SP.CREADOS if c == "CTkLabel"]
                misma_vieja = [textos.count("✓  Misma carpeta que Vitro")]
                self.assertEqual(misma_vieja, [sum(m["misma_carpeta"] for m in nuevo["mallas"])])


if __name__ == "__main__":
    unittest.main(verbosity=2)
