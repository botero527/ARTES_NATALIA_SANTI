"""agp_sso.py — token de sesión única para el ecosistema de herramientas AGP.

El Panel de Herramientas (launcher) hace login una sola vez contra
[mallas].[APP_USUARIOS] y firma un token con crear_token(). Al abrir una
app de escritorio, la lanza con "--sso-token=<token>". Cada app copia este
mismo archivo y usa verificar_token() para saber si puede saltarse su
propia pantalla de login.

Solo usa la librería estándar (hmac/hashlib/json/base64/time) — ninguna
dependencia nueva que instalar en ningún proyecto.

IMPORTANTE: SSO_SECRET debe ser IDÉNTICO en las 4 copias de este archivo
(Panel, Artes, Redibujado, PipeMirror). Si se cambia aquí, hay que
cambiarlo en las otras 3 también, o los tokens dejan de validar.
"""

import base64
import hashlib
import hmac
import json
import time

SSO_SECRET = b"6f87d58f0e1d75146747271849d8d35e064491d69a0d49b6d7a6ea1943bdd662"
TOKEN_TTL_SEGUNDOS = 8 * 60 * 60  # 8 horas — dura una jornada de trabajo


def _firmar(payload_b64: bytes) -> str:
    return hmac.new(SSO_SECRET, payload_b64, hashlib.sha256).hexdigest()


def crear_token(usuario: str, nombre: str = "", rol: str = "") -> str:
    """Genera un token firmado para `usuario`. Solo lo llama el Panel."""
    payload = {
        "usuario": usuario,
        "nombre": nombre,
        "rol": rol,
        "exp": time.time() + TOKEN_TTL_SEGUNDOS,
    }
    payload_b64 = base64.urlsafe_b64encode(json.dumps(payload).encode("utf-8"))
    firma = _firmar(payload_b64)
    return payload_b64.decode("ascii") + "." + firma


def verificar_token(token: str):
    """Devuelve el payload (dict) si el token es válido y no ha vencido.

    Devuelve None si el token es inválido, está corrupto, la firma no
    coincide, o ya expiró. Las apps NO deben confiar en `nombre`/`rol` del
    payload para nada sensible — son solo pistas; el dato real hay que
    volver a leerlo de APP_USUARIOS por `usuario` antes de dar acceso.
    """
    if not token or "." not in token:
        return None
    payload_b64_str, firma = token.rsplit(".", 1)
    payload_b64 = payload_b64_str.encode("ascii")
    firma_esperada = _firmar(payload_b64)
    if not hmac.compare_digest(firma, firma_esperada):
        return None
    try:
        payload = json.loads(base64.urlsafe_b64decode(payload_b64))
    except Exception:
        return None
    if payload.get("exp", 0) < time.time():
        return None
    return payload


def token_de_argv(argv):
    """Busca --sso-token=<token> en una lista de argumentos (sys.argv)."""
    prefijo = "--sso-token="
    for arg in argv:
        if arg.startswith(prefijo):
            return arg[len(prefijo):]
    return None
