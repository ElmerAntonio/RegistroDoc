import pytest
import os
import tempfile
import json
# Master salt dummy setup para pruebas locales DEBE IR ANTES DEL IMPORT
os.environ["REGISTRODOC_MASTER_SALT"] = "TEST_SALT"

import rdsecurity
from rdsecurity import (
    cifrar, descifrar, guardar_cifrado, cargar_cifrado,
    _hw_fingerprint, verificar_licencia, validar_nota_meduca,
    _leer_cedula_token, _guardar_cedula_token
)
from unittest.mock import patch

@pytest.fixture
def temp_file():
    fd, path = tempfile.mkstemp()
    os.close(fd)
    yield path
    try:
        os.remove(path)
    except OSError:
        pass

@patch('rdsecurity._hw_fingerprint')
def test_cryptography_roundtrip(mock_hw):
    mock_hw.return_value = b"MOCK_HW_123"
    password = b"test_password"
    data = b"informacion super secreta"

    # Cifrado real
    cifrado = cifrar(data, password)

    # Descifrado real
    descifrado = descifrar(cifrado, password)
    assert descifrado == data

    # Prueba fallos manipulando el tag GCM
    cifrado_roto = bytearray(cifrado)
    cifrado_roto[-1] ^= 0x01  # Flipping the last bit of the tag
    with pytest.raises(ValueError):
         descifrar(bytes(cifrado_roto), password)

@patch('rdsecurity._hw_fingerprint')
def test_guardar_cargar_cifrado(mock_hw, temp_file):
    mock_hw.return_value = b"MOCK_HW_123"
    datos_origen = {"usuario": "admin", "privilegio": 1}

    # Usar IO de archivo real cifrando y descifrando
    guardar_cifrado(temp_file, datos_origen)
    datos_leidos = cargar_cifrado(temp_file)
    assert datos_origen == datos_leidos

def test_cargar_cifrado_archivo_inexistente():
    # Sin mockear os.path.exists, usar ruta ficticia explícita
    ruta_ficticia = r"C:\ruta\totalmente\falsa\inexistente.enc"
    resultado = cargar_cifrado(ruta_ficticia)
    assert resultado == {}  # Debería devolver {} si no existe

@pytest.mark.parametrize("valor,esperado,esperado_val", [
    (None, False, 0.0),
    (0, False, 0.0),
    (0.0, False, 0.0),
    ("", False, 0.0),
    ("   ", False, 0.0),
    (False, False, 0.0),
    ([], False, 0.0),
    (3.5, True, 3.5),
    (5.0, True, 5.0),
    ("4,2", True, 4.2),  # Manejo de coma decimal
    (6.0, False, 6.0),   # Arriba del rango
    (0.9, False, 0.9)    # Abajo del rango
])
def test_validar_nota_meduca(valor, esperado, esperado_val):
    valido, nota, _ = validar_nota_meduca(valor)
    assert valido == esperado
    assert nota == esperado_val

def test_leer_cedula_token_corrupto(temp_file):
    # Mockear el nombre del archivo para usar un archivo temporal
    with patch('rdsecurity.open', create=True) as mock_open:
        # Simular archivo que existe pero tiene basura que descifrar fallará
        mock_open.return_value.__enter__.return_value.read.return_value = b"NOT_A_VALID_BLOB"
        with patch('os.path.exists', return_value=True):
            # Debería retornar {} al fallar descifrar o json.loads
            assert _leer_cedula_token() == {}

def test_cedula_token_roundtrip():
    cedula = "8-888-888"
    try:
        _guardar_cedula_token(cedula)
        datos = _leer_cedula_token()
        assert datos["hint"] == cedula
    finally:
        if os.path.exists("rd_token.bin"):
            os.remove("rd_token.bin")
