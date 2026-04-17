"""
Verificacion: busca si ya existe un arte con el mismo codigo en la carpeta de artes.
Estructura: RUTA_BASE / VEHICULO / MODELO / V*** / ARTES / *.dwg|3dm
"""
import os


def extraer_sufijo(nombre_archivo, n_partes=2):
    """
    Extrae los últimos n_partes grupos del nombre antes de la extensión.
    Ejemplo: "1708 008 030 A.dwg"  →  "030 A"
    """
    base = os.path.splitext(os.path.basename(nombre_archivo))[0].strip()
    partes = base.split()
    if len(partes) >= n_partes:
        return " ".join(partes[-n_partes:])
    return base


def _localizar_artes(ruta_version: str) -> str | None:
    """
    Dado el path de una version (o accidentalmente de ARTES),
    devuelve la ruta real de la carpeta ARTES.
    """
    # Si el usuario ya seleccionó ARTES directamente
    if os.path.basename(ruta_version).upper() == "ARTES":
        return ruta_version

    # Buscar subcarpeta ARTES (case-insensitive)
    try:
        for sub in os.listdir(ruta_version):
            if sub.upper() == "ARTES" and os.path.isdir(
                os.path.join(ruta_version, sub)
            ):
                return os.path.join(ruta_version, sub)
    except Exception:
        pass
    return None


def listar_artes(ruta_version: str) -> list:
    """
    Devuelve TODOS los archivos .dwg/.3dm que haya en la carpeta ARTES
    de la version dada, sin filtrar por nombre.
    """
    ruta_artes = _localizar_artes(ruta_version)
    if not ruta_artes:
        return []
    resultado = []
    try:
        for archivo in sorted(os.listdir(ruta_artes)):
            ext = os.path.splitext(archivo)[1].lower()
            if ext in (".dwg", ".3dm"):
                resultado.append(
                    {
                        "archivo":    archivo,
                        "ruta":       os.path.join(ruta_artes, archivo),
                        "ruta_artes": ruta_artes,
                        "coincide":   False,
                    }
                )
    except Exception:
        pass
    return resultado


def buscar_en_version(ruta_version: str, sufijo: str) -> list:
    """
    Lista todos los archivos en ARTES y marca cuáles coinciden con el sufijo.
    Devuelve lista de dicts con campo 'coincide'.
    """
    todos = listar_artes(ruta_version)
    if not sufijo:
        return todos

    sufijo_cmp = sufijo.upper().strip()
    for item in todos:
        nombre = os.path.splitext(item["archivo"])[0].upper()
        item["coincide"] = sufijo_cmp in nombre

    return todos


def ruta_artes_de_version(ruta_version: str) -> str:
    """
    Devuelve la ruta a la carpeta ARTES dentro de ruta_version.
    La crea si no existe. Maneja el caso donde el usuario ya dio la ruta ARTES.
    """
    if os.path.basename(ruta_version).upper() == "ARTES":
        ruta = ruta_version
    else:
        ruta = os.path.join(ruta_version, "ARTES")
    os.makedirs(ruta, exist_ok=True)
    return ruta


def buscar_artes_existentes(ruta_base, vehiculo, modelo, sufijo):
    """
    Busca archivos en [ruta_base]/[vehiculo]/[modelo]/V*/ARTES/
    cuyo nombre contenga el sufijo dado.

    Retorna lista de dicts:
        { 'version', 'archivo', 'ruta' }
    """
    encontrados = []
    ruta_vehiculo = os.path.join(ruta_base, vehiculo, modelo)

    if not os.path.isdir(ruta_vehiculo):
        return encontrados

    try:
        entradas = os.listdir(ruta_vehiculo)
    except PermissionError:
        return encontrados

    versiones = sorted(
        d for d in entradas
        if os.path.isdir(os.path.join(ruta_vehiculo, d))
        and d.upper().startswith("V")
    )

    sufijo_cmp = sufijo.upper().strip()

    for version in versiones:
        ruta_version = os.path.join(ruta_vehiculo, version)

        # Buscar carpeta ARTES (case-insensitive)
        ruta_artes = None
        try:
            for sub in os.listdir(ruta_version):
                if sub.upper() == "ARTES" and os.path.isdir(
                    os.path.join(ruta_version, sub)
                ):
                    ruta_artes = os.path.join(ruta_version, sub)
                    break
        except Exception:
            continue

        if not ruta_artes:
            continue

        try:
            archivos = os.listdir(ruta_artes)
        except Exception:
            continue

        for archivo in archivos:
            ext = os.path.splitext(archivo)[1].lower()
            if ext not in (".dwg", ".3dm"):
                continue
            nombre_cmp = os.path.splitext(archivo)[0].upper()
            if sufijo_cmp in nombre_cmp:
                encontrados.append(
                    {
                        "version": version,
                        "archivo": archivo,
                        "ruta":    os.path.join(ruta_artes, archivo),
                    }
                )

    return encontrados


def ruta_destino_arte(ruta_base, vehiculo, modelo, version, nombre_arte):
    """
    Construye la ruta completa de destino para guardar el arte.
    Crea la carpeta ARTES si no existe.
    """
    ruta_artes = os.path.join(ruta_base, vehiculo, modelo, version, "ARTES")
    os.makedirs(ruta_artes, exist_ok=True)
    return os.path.join(ruta_artes, nombre_arte)
