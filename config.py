import os

_DIR = os.path.dirname(os.path.abspath(__file__))

# ─── RUTAS BASE ────────────────────────────────────────────────────────────────
# Carpeta raíz donde están las marcas  (ajustar si la red cambia)
RUTA_BASE = r"\\192.168.2.37\ingenieria\PRODUCCION\AGP PLANOS TECNICOS"

# Plantilla Rhino con layers y cajetines (mismo directorio que este script)
CAJETINES_DWG = os.path.join(_DIR, "LAYERS Y CAJETINES 1.dwg")

# Ejecutable de Rhino 8
RHINO_EXE = r"C:\Program Files\Rhino 8\System\Rhino.exe"

# ─── PATRONES DE LAYERS EN AUTOCAD ────────────────────────────────────────────
# Se hace coincidencia parcial (el nombre del layer CONTIENE el patrón)
PATRONES_PERIMETRO = ["PERIMETRO"]
PATRONES_BN        = ["BANDA NEGRA", "BANDANEGRA", "BN", "PHANTOM"]
PATRONES_LOGO      = ["LOGO", "TRAZABILIDAD"]

# ─── NOMBRES DE LAYERS EN RHINO ───────────────────────────────────────────────
LAYER_PLANES = "PLANES"
LAYER_K2     = "k2"
LAYER_K      = "k"

# ─── PARÁMETROS GEOMÉTRICOS (en mm) ──────────────────────────────────────────
OFFSET_PERIMETRO  = 0.5    # offset del perímetro
OFFSET_BN_DEGRADE = 2.5    # offset de BN hacia adentro (degradé)
DIVISOR_DEGRADE   = 3      # longitud / 3 = número de pepas
LOGO_MARGEN_1     = 3.0    # distancia bajo logo para línea 1
LOGO_MARGEN_2     = 3.8    # distancia bajo logo para línea 2

# ─── BLOQUE 25 (pepas para degradé) ──────────────────────────────────────────
# Nombre exacto del bloque en el DWG de cajetines
NOMBRE_BLOQUE_25 = "25"

# ─── REPORTE ─────────────────────────────────────────────────────────────────
REPORTE_EXCEL = os.path.join(_DIR, "Reporte_Artes.xlsx")
