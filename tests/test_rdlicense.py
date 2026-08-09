import os
import sys
import base64

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "src")))

import pytest
from cryptography.hazmat.primitives.asymmetric.ed25519 import Ed25519PrivateKey
from cryptography.hazmat.primitives import serialization

import rdlicense


def _keypair_efimero():
    priv = Ed25519PrivateKey.generate()
    pub_raw = priv.public_key().public_bytes(
        serialization.Encoding.Raw, serialization.PublicFormat.Raw)
    return priv, base64.b64encode(pub_raw).decode()


def test_generar_y_verificar_round_trip():
    priv, pub_b64 = _keypair_efimero()
    codigo = rdlicense.generar_codigo(priv)
    assert codigo.startswith("RD-")
    assert rdlicense.verificar_codigo(codigo, public_key_b64=pub_b64) is True


def test_codigo_de_otra_clave_no_verifica():
    priv1, _ = _keypair_efimero()
    _, pub2_b64 = _keypair_efimero()
    codigo = rdlicense.generar_codigo(priv1)
    # Firmado con priv1 pero verificado con la pública de OTRO par → inválido
    assert rdlicense.verificar_codigo(codigo, public_key_b64=pub2_b64) is False


def test_codigo_manipulado_no_verifica():
    priv, pub_b64 = _keypair_efimero()
    codigo = rdlicense.generar_codigo(priv)
    # Alterar un carácter del cuerpo del código
    partes = codigo.split("-")
    cuerpo = list(partes[-1])
    cuerpo[0] = "A" if cuerpo[0] != "A" else "B"
    partes[-1] = "".join(cuerpo)
    manipulado = "-".join(partes)
    assert rdlicense.verificar_codigo(manipulado, public_key_b64=pub_b64) is False


def test_basura_no_verifica():
    _, pub_b64 = _keypair_efimero()
    for basura in ["", "RD-XXXXX", "no-es-codigo", "RD-000000-000000"]:
        assert rdlicense.verificar_codigo(basura, public_key_b64=pub_b64) is False


def test_lote_sin_repetidos():
    priv, pub_b64 = _keypair_efimero()
    lote = rdlicense.generar_lote(priv, 200)
    assert len(lote) == 200
    seriales = [rdlicense.serial_de(c) for c in lote]
    assert len(set(seriales)) == 200          # todos únicos
    assert all(rdlicense.verificar_codigo(c, public_key_b64=pub_b64) for c in lote)


def test_clave_privada_del_vendedor_valida_con_publica_embebida():
    """Si existe la clave privada real del vendedor, un código generado con ella
    DEBE validar con la clave pública embebida en la app."""
    ruta = r"C:\RegistroDoc_Mis_Ventas\clave_privada_vendedor.key"
    if not os.path.exists(ruta):
        pytest.skip("Clave privada del vendedor no presente en este entorno")
    priv = rdlicense.cargar_clave_privada(ruta)
    codigo = rdlicense.generar_codigo(priv)
    assert rdlicense.verificar_codigo(codigo) is True   # usa la pública embebida
