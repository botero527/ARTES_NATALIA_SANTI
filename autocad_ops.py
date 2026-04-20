"""
Operaciones AutoCAD vía COM (pywin32).
Estrategia: copiar el DWG origen, abrirlo en AutoCAD y eliminar todo lo que
NO esté en los layers objetivo (perimetro, bn/phantom, logo).
"""
import os
import shutil
import time
import sys

try:
    import win32com.client
    import pythoncom
except ImportError:
    print("Falta pywin32.  Ejecuta:  pip install pywin32")
    sys.exit(1)

from config import PATRONES_PERIMETRO, PATRONES_BN, PATRONES_LOGO

_TODOS_PATRONES = PATRONES_PERIMETRO + PATRONES_BN + PATRONES_LOGO


# ─── helpers ──────────────────────────────────────────────────────────────────

def _es_layer_objetivo(nombre_layer: str) -> bool:
    n = nombre_layer.upper().strip()
    for p in _TODOS_PATRONES:
        if p.upper() in n:
            return True
    return False


def _tipo_layer(nombre_layer: str) -> str | None:
    n = nombre_layer.upper().strip()
    for p in PATRONES_PERIMETRO:
        if p.upper() in n:
            return "perimetro"
    for p in PATRONES_BN:
        if p.upper() in n:
            return "bn"
    for p in PATRONES_LOGO:
        if p.upper() in n:
            return "logo"
    return None


def detectar_layers_en_doc(doc) -> dict:
    """
    Devuelve qué layers objetivo encontró en el documento.
    { 'perimetro': 'PERIMETRO', 'bn': 'BN', 'logo': ['LOGO1', ...] }
    """
    resultado = {"perimetro": None, "bn": None, "logo": []}
    try:
        for i in range(doc.Layers.Count):
            nombre = doc.Layers.Item(i).Name
            tipo = _tipo_layer(nombre)
            if tipo == "perimetro" and resultado["perimetro"] is None:
                resultado["perimetro"] = nombre
            elif tipo == "bn" and resultado["bn"] is None:
                resultado["bn"] = nombre
            elif tipo == "logo":
                resultado["logo"].append(nombre)
    except Exception:
        pass
    return resultado


# ─── motor AutoCAD ────────────────────────────────────────────────────────────

class AutoCADMotor:
    def __init__(self):
        pythoncom.CoInitialize()
        try:
            self.acad = win32com.client.GetActiveObject("AutoCAD.Application")
        except Exception:
            raise RuntimeError(
                "AutoCAD no está abierto.\n"
                "Abre AutoCAD primero y vuelve a ejecutar."
            )

    # ── abrir / cerrar ──────────────────────────────────────────────────────

    def abrir(self, ruta: str, readonly: bool = False, espera: float = 2.0):
        ruta_abs = os.path.abspath(ruta)
        doc = self.acad.Documents.Open(ruta_abs, False, readonly)
        time.sleep(espera)
        return doc

    def cerrar(self, doc, guardar: bool = False):
        try:
            doc.Close(guardar)
        except Exception:
            pass
 # ── operación principal ─────────────────────────────────────────────────

    def extraer_layers(
        self,
        ruta_origen: str,
        ruta_destino: str,
        log_fn=None,
    ) -> tuple[str, dict]:
        """
        1. Copia ruta_origen → ruta_destino (sin tocar el original).
        2. Abre la copia en AutoCAD.
        3. Elimina del ModelSpace todos los objetos que NO estén en layers objetivo.
        4. Guarda y cierra.

        Devuelve (ruta_destino, dict_layers_encontrados).
        """
        if log_fn is None:
            log_fn = print

        # 1. Copia física del archivo
        log_fn(f"  Copiando plano a carpeta temporal...")
        os.makedirs(os.path.dirname(ruta_destino), exist_ok=True)
        shutil.copy2(ruta_origen, ruta_destino)

        # 2. Abrir copia
        log_fn(f"  Abriendo copia en AutoCAD...")
        doc = self.abrir(ruta_destino, readonly=False)

        # 3. Detectar layers
        layers_info = detectar_layers_en_doc(doc)
        log_fn(
            f"  Layers detectados — "
            f"perimetro: {layers_info['perimetro']}  "
            f"bn: {layers_info['bn']}  "
            f"logo: {layers_info['logo']}"
        )

        if not any([layers_info["perimetro"], layers_info["bn"], layers_info["logo"]]):
            self.cerrar(doc, guardar=False)
            raise ValueError(
                "No se encontraron layers de perimetro, banda negra ni logo.\n"
                "Verifica que el plano tenga esos layers."
            )

        # 4. Eliminar objetos fuera de layers objetivo
        ms = doc.ModelSpace
        a_borrar = []
        for i in range(ms.Count):
            try:
                ent = ms.Item(i)
                if not _es_layer_objetivo(ent.Layer):
                    a_borrar.append(ent)
            except Exception:
                pass

        log_fn(f"  Eliminando {len(a_borrar)} objetos no requeridos...")
        for ent in a_borrar:
            try:
                ent.Delete()
            except Exception:
                pass

        # 5. Guardar y cerrar
        log_fn(f"  Guardando archivo filtrado...")
        doc.Save()
        time.sleep(0.5)
        self.cerrar(doc, guardar=False)
        log_fn(f"  Extracción completada: {os.path.basename(ruta_destino)}")

        return ruta_destino, layers_info

    def quit(self):
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass
