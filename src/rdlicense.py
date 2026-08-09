"""
RegistroDoc Pro — Sistema de Licencias (firma asimétrica Ed25519)
==================================================================
Modelo de seguridad CORRECTO para venta de software offline:

  • La CLAVE PRIVADA vive SOLO en la herramienta del vendedor
    (C:\\RegistroDoc_Mis_Ventas\\clave_privada_vendedor.key) y NUNCA se distribuye.
  • La CLAVE PÚBLICA está embebida aquí (es seguro publicarla): solo permite
    VERIFICAR códigos, jamás generarlos. Un cliente con el .exe NO puede fabricar
    códigos válidos, porque no tiene la clave privada.

Los códigos son PRE-GENERADOS por lote (únicos, no repetidos) y se guardan en un
documento para entregarlos en cada venta, sin procesar nada en el momento.

© 2026 RegistroDoc Pro — Todos los derechos reservados
"""

import base64
import secrets

from cryptography.hazmat.primitives.asymmetric.ed25519 import (
    Ed25519PrivateKey, Ed25519PublicKey,
)

# ── Clave pública del vendedor (segura de distribuir; SOLO verifica) ──
PUBLIC_KEY_B64 = "gGW+K9Cc8STqwFXIDjXqRhSOHX5kjwpPmc7WmNt0T+k="

_PREFIJO = "RD"
_SERIAL_LEN = 5          # 40 bits aleatorios → unicidad práctica
_FIRMA_LEN = 64          # Ed25519 produce firmas de 64 bytes


# ── Formato legible del código (grupos de 6 con guiones) ──
def _fmt(body: str) -> str:
    grupos = [body[i:i + 6] for i in range(0, len(body), 6)]
    return _PREFIJO + "-" + "-".join(grupos)


def _unfmt(codigo: str) -> str:
    limpio = (codigo or "").strip().upper().replace("-", "").replace(" ", "")
    if limpio.startswith(_PREFIJO):
        limpio = limpio[len(_PREFIJO):]
    return limpio


def _b32(blob: bytes) -> str:
    return base64.b32encode(blob).decode("ascii").rstrip("=")


def _unb32(body: str) -> bytes:
    pad = "=" * ((8 - len(body) % 8) % 8)
    return base64.b32decode(body + pad)


# ══════════════════════════════════════════════════════════════
#  VENDEDOR — generación (requiere la CLAVE PRIVADA)
# ══════════════════════════════════════════════════════════════

def cargar_clave_privada(ruta: str) -> Ed25519PrivateKey:
    """Carga la clave privada del vendedor desde su archivo (base64 de 32 bytes)."""
    with open(ruta, "rb") as f:
        raw = base64.b64decode(f.read())
    return Ed25519PrivateKey.from_private_bytes(raw)


def generar_codigo(clave_privada: Ed25519PrivateKey) -> str:
    """[SOLO VENDEDOR] Genera UN código único firmado.
    serial aleatorio (único) + firma Ed25519 sobre ese serial."""
    serial = secrets.token_bytes(_SERIAL_LEN)
    firma = clave_privada.sign(serial)
    return _fmt(_b32(serial + firma))


def generar_lote(clave_privada: Ed25519PrivateKey, cantidad: int) -> list:
    """[SOLO VENDEDOR] Genera un LOTE de códigos únicos (sin repetidos)."""
    vistos = set()
    codigos = []
    while len(codigos) < cantidad:
        c = generar_codigo(clave_privada)
        s = serial_de(c)
        if s in vistos:
            continue  # colisión (astronómicamente rara) → reintentar
        vistos.add(s)
        codigos.append(c)
    return codigos


# ══════════════════════════════════════════════════════════════
#  APP — verificación (SOLO clave pública; no puede generar)
# ══════════════════════════════════════════════════════════════

def verificar_codigo(codigo: str, public_key_b64: str = None) -> bool:
    """[APP] True solo si el código fue firmado por la clave privada del vendedor.
    Imposible de falsificar sin dicha clave privada.
    `public_key_b64` permite inyectar otra pública (solo para pruebas); por defecto
    usa la clave pública embebida del vendedor."""
    try:
        blob = _unb32(_unfmt(codigo))
        if len(blob) != _SERIAL_LEN + _FIRMA_LEN:
            return False
        serial, firma = blob[:_SERIAL_LEN], blob[_SERIAL_LEN:]
        pub = Ed25519PublicKey.from_public_bytes(
            base64.b64decode(public_key_b64 or PUBLIC_KEY_B64))
        pub.verify(firma, serial)   # lanza si la firma no es válida
        return True
    except Exception:
        return False


def serial_de(codigo: str) -> str:
    """Identificador único del código (hex del serial). Sirve para detectar
    códigos repetidos y para registrar cuál se activó, sin exponer la firma."""
    try:
        blob = _unb32(_unfmt(codigo))
        return blob[:_SERIAL_LEN].hex().upper()
    except Exception:
        return ""
