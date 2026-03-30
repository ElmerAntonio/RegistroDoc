import sys
from unittest.mock import MagicMock

# Mock cryptography before importing src.rdsecurity
mock_crypto = MagicMock()
sys.modules["cryptography"] = mock_crypto
sys.modules["cryptography.hazmat"] = mock_crypto
sys.modules["cryptography.hazmat.primitives"] = mock_crypto
sys.modules["cryptography.hazmat.primitives.ciphers"] = mock_crypto
sys.modules["cryptography.hazmat.primitives.ciphers.aead"] = mock_crypto
sys.modules["cryptography.hazmat.primitives.kdf"] = mock_crypto
sys.modules["cryptography.hazmat.primitives.kdf.pbkdf2"] = mock_crypto
sys.modules["cryptography.hazmat.backends"] = mock_crypto

import pytest
from src.rdsecurity import validar_nota_meduca

def test_validar_nota_meduca_happy_path():
    """Test happy path scenarios for MEDUCA grade validation."""
    # Integers as strings
    ok, nota, msg = validar_nota_meduca("4")
    assert ok is True
    assert nota == 4.0
    assert msg == "OK"

    # Floats as strings
    ok, nota, msg = validar_nota_meduca("3.5")
    assert ok is True
    assert nota == 3.5
    assert msg == "OK"

    # Boundary values
    ok, nota, msg = validar_nota_meduca("1.0")
    assert ok is True
    assert nota == 1.0
    assert msg == "OK"

    ok, nota, msg = validar_nota_meduca("5.0")
    assert ok is True
    assert nota == 5.0
    assert msg == "OK"

def test_validar_nota_meduca_normalization():
    """Test normalization of grade strings (whitespace and commas)."""
    # Whitespace
    ok, nota, msg = validar_nota_meduca("  4.5  ")
    assert ok is True
    assert nota == 4.5

    # Comma as decimal separator
    ok, nota, msg = validar_nota_meduca("4,2")
    assert ok is True
    assert nota == 4.2

def test_validar_nota_meduca_edge_cases():
    """Test edge cases for MEDUCA grade validation."""
    # Just below minimum
    ok, nota, msg = validar_nota_meduca("0.9")
    assert ok is False
    assert "mínima" in msg

    # Just above maximum
    ok, nota, msg = validar_nota_meduca("5.1")
    assert ok is False
    assert "máxima" in msg

    # Extra decimals (Python's round() uses banker's rounding)
    # round(4.55, 1) -> 4.5
    # round(4.65, 1) -> 4.6

    ok, nota, msg = validar_nota_meduca("4.55")
    assert ok is True
    assert nota == 4.5

    ok, nota, msg = validar_nota_meduca("4.56")
    assert ok is True
    assert nota == 4.6

def test_validar_nota_meduca_invalid_inputs():
    """Test invalid inputs for MEDUCA grade validation."""
    # Empty inputs
    ok, _, msg = validar_nota_meduca("")
    assert ok is False
    assert "vacía" in msg

    ok, _, msg = validar_nota_meduca("   ")
    assert ok is False
    assert "vacía" in msg

    # Non-numeric
    ok, _, msg = validar_nota_meduca("abc")
    assert ok is False
    assert "no es un número válido" in msg

    # Multiple dots
    ok, _, msg = validar_nota_meduca("4.5.5")
    assert ok is False
    assert "no es un número válido" in msg

    # None input
    ok, _, msg = validar_nota_meduca(None)
    assert ok is False
    assert "vacía" in msg

import os
from unittest.mock import patch
from src.rdsecurity import cargar_cifrado

@patch('src.rdsecurity._hw_fingerprint')
def test_cargar_cifrado_missing_file(mock_hw_fingerprint):
    # Mock hw_fingerprint to be deterministic
    mock_hw_fingerprint.return_value = b"test_deterministic_hw_fingerprint"

    # Pass a clearly non-existent path
    ruta_invalida = "ruta_invalida_que_no_existe.enc"

    # Ensure the file doesn't actually exist
    if os.path.exists(ruta_invalida):
        os.remove(ruta_invalida)

    resultado = cargar_cifrado(ruta_invalida)

    assert resultado == {}, "Expected empty dict for missing file"

@patch('src.rdsecurity._hw_fingerprint')
def test_cargar_cifrado_invalid_data(mock_hw_fingerprint, tmp_path):
    # Mock hw_fingerprint to be deterministic
    mock_hw_fingerprint.return_value = b"test_deterministic_hw_fingerprint"

    # Create a temporary file path
    temp_file = tmp_path / "basura.enc"

    # Write garbage data to the file (this should cause decryption or json parsing to fail)
    temp_file.write_bytes(os.urandom(100))

    # Call cargar_cifrado and assert it returns an empty dict
    resultado = cargar_cifrado(str(temp_file))

    assert resultado == {}, "Expected empty dict when reading garbage data"
