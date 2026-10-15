# -*- coding: utf-8 -*-
"""
Soporte de las pruebas de equivalencia app vieja (agp_app.py, CustomTkinter)
vs. app nueva (agp_web.py + servicios.py + ui_web/).

- Instala un `customtkinter` falso para poder importar agp_app.py sin la
  librería (Python 3.14 del sistema no la tiene): cualquier widget es un
  `Dummy` que acepta todo y recuerda el último texto que se le puso.
- `extraer(ruta, externa, interna)` saca con `ast` el código fuente LITERAL de
  una función anidada (p. ej. el `_guardar` dentro de `_abrir_editor`) para
  ejecutarla igualita, con sus variables de cierre inyectadas.
- `BDGrabadora` reemplaza la conexión a Azure: registra cada SQL + parámetros
  y responde con datos programados. Ninguna prueba de escritura toca la BD real.
- `js(archivos, fn, casos)` corre funciones puras de ui_web/js en Node.
"""

import ast, os, sys, json, types, textwrap, subprocess, re

PROY = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if PROY not in sys.path:
    sys.path.insert(0, PROY)

# ── customtkinter falso ───────────────────────────────────────────────────
CREADOS = []   # (clase, kwargs) de cada widget creado, en orden


class Dummy:
    def __init__(self, *a, **k):
        self._kw = dict(k)
        self.texto = k.get("text")
        CREADOS.append((type(self).__name__, self._kw))

    def configure(self, **k):
        self._kw.update(k)
        if "text" in k:
            self.texto = k["text"]

    config = configure

    def cget(self, k):
        return self._kw.get(k)

    def get(self):
        return self._kw.get("_valor", "")

    def __getattr__(self, nombre):
        if nombre.startswith("__"):
            raise AttributeError(nombre)
        return lambda *a, **k: Dummy()


def _instalar_ctk_falso():
    if "customtkinter" in sys.modules and getattr(sys.modules["customtkinter"], "_FALSO", False):
        return
    m = types.ModuleType("customtkinter")
    m._FALSO = True

    def __getattr__(nombre):
        if nombre.startswith("__"):
            raise AttributeError(nombre)
        return type(nombre, (Dummy,), {})
    m.__getattr__ = __getattr__
    m.set_appearance_mode = lambda *a, **k: None
    m.set_default_color_theme = lambda *a, **k: None
    sys.modules["customtkinter"] = m


def importar_app_vieja():
    _instalar_ctk_falso()
    import agp_app
    return agp_app


# ── Funciones anidadas originales ─────────────────────────────────────────
def extraer(ruta, externa, interna, clase=None):
    """Fuente literal de `def interna` dentro de `def externa` (opcional: método de `clase`)."""
    src = open(ruta, encoding="utf-8").read()
    arbol = ast.parse(src)
    nodos = arbol.body
    if clase:
        nodos = next(n for n in arbol.body if isinstance(n, ast.ClassDef) and n.name == clase).body
    ext = next(n for n in nodos if isinstance(n, ast.FunctionDef) and n.name == externa)
    for n in ast.walk(ext):
        if isinstance(n, ast.FunctionDef) and n.name == interna:
            return textwrap.dedent(ast.get_source_segment(src, n))
    raise LookupError(f"{interna} no está en {externa}")


def extraer_asignacion(ruta, externa, variable):
    src = open(ruta, encoding="utf-8").read()
    arbol = ast.parse(src)
    ext = next(n for n in arbol.body if isinstance(n, ast.FunctionDef) and n.name == externa)
    for n in ast.walk(ext):
        if isinstance(n, ast.Assign) and any(getattr(t, "id", None) == variable for t in n.targets):
            return textwrap.dedent(ast.get_source_segment(src, n))
    raise LookupError(variable)


def compilar(fuente, nombre, entorno):
    """Ejecuta el `def` original en `entorno` (sus variables libres) y devuelve la función."""
    exec(fuente, entorno)
    return entorno[nombre]


# ── BD grabadora ──────────────────────────────────────────────────────────
def _norm_sql(sql):
    return re.sub(r"\s+", " ", sql.replace("?", "%s")).strip()


def _norm_param(v):
    return None if v is None else str(v)


class _Cursor:
    def __init__(self, bd):
        self.bd = bd
        self._filas = []
        self.rowcount = -1
        self.description = None

    def execute(self, sql, params=()):
        p = tuple(_norm_param(x) for x in (params or ()))
        self.bd.log.append((_norm_sql(sql), p))
        filas, self.rowcount = self.bd.responder(_norm_sql(sql), p)
        self._filas = list(filas or [])

    def fetchone(self):
        return self._filas.pop(0) if self._filas else None

    def fetchall(self):
        f, self._filas = self._filas, []
        return f


class BDGrabadora:
    def __init__(self, responder=None):
        self.log = []
        self.commits = 0
        self.responder = responder or (lambda sql, p: ([], 1))

    def __call__(self):           # se usa como reemplazo de db_connect()
        return self

    def cursor(self):
        return _Cursor(self)

    def execute(self, sql, params=()):
        c = self.cursor(); c.execute(sql, params); return c

    def commit(self):
        self.commits += 1
        self.log.append(("COMMIT", ()))

    def rollback(self):
        self.log.append(("ROLLBACK", ()))

    def close(self):
        pass


# ── Grabadora de llamadas ─────────────────────────────────────────────────
class Grabadora:
    def __init__(self, devuelve=None):
        self.llamadas = []
        self.devuelve = devuelve

    def __call__(self, *a, **k):
        self.llamadas.append((a, k))
        return self.devuelve(*a, **k) if callable(self.devuelve) else self.devuelve


class Campo:
    """Entry / StringVar falso: solo .get() (y lo que el código viejo le hace encima)."""
    def __init__(self, valor=""):
        self.valor = valor

    def get(self):
        return self.valor

    def delete(self, *a):
        self.valor = ""

    def insert(self, _i, v):
        self.valor = self.valor + str(v)

    def configure(self, **k):
        pass

    config = configure


# ── Node ──────────────────────────────────────────────────────────────────
_RUNNER = r"""
const fs = require("fs"), vm = require("vm");
const entrada = JSON.parse(fs.readFileSync(0, "utf8"));
const ctx = { console }; vm.createContext(ctx);
vm.runInContext("var $=()=>null,$$=()=>[],t=k=>k,esc=s=>s,icon=()=>'',api=async()=>({ok:true}),toast=()=>{},Bus={on(){}},I18N={es:{}};", ctx);
for (const f of entrada.archivos) vm.runInContext(fs.readFileSync(f, "utf8"), ctx, { filename: f });
const fn = vm.runInContext("(" + entrada.fn + ")", ctx);
process.stdout.write(JSON.stringify(entrada.casos.map((c) => fn(...c))));
"""


def js(archivos, fn, casos):
    rutas = [os.path.join(PROY, "ui_web", "js", a) for a in archivos]
    r = subprocess.run(["node", "-e", _RUNNER], input=json.dumps({"archivos": rutas, "fn": fn, "casos": casos}),
                       capture_output=True, text=True, encoding="utf-8")
    if r.returncode != 0:
        raise RuntimeError(r.stderr)
    return json.loads(r.stdout)
