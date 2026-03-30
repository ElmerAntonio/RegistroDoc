import sys
from unittest.mock import MagicMock

try:
    import cryptography
    HAS_CRYPTOGRAPHY = True
except ImportError:
    HAS_CRYPTOGRAPHY = False
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
from src.rdsecurity import validar_nota_meduca, cifrar, descifrar

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

@pytest.mark.skipif(not HAS_CRYPTOGRAPHY, reason="Requires real cryptography module")
@pytest.mark.parametrize("invalid_blob", [
    None,
    "not a bytes object",
    [],
    b"",
])
def test_descifrar_invalid_types_raises_error(invalid_blob):
    """Test that descifrar properly rejects invalid data types and empty bytes."""
    password = b"secret_password"
    with pytest.raises((TypeError, ValueError)):
        descifrar(invalid_blob, password)


@pytest.mark.skipif(not HAS_CRYPTOGRAPHY, reason="Requires real cryptography module")
def test_descifrar_payload_too_small_raises_value_error():
    """Test that a payload smaller than minimum length (4 + 32 + 12 + 16 = 64 bytes) raises ValueError."""
    password = b"secret_password"
    # Create payload of 63 bytes (one byte short of 64)
    short_payload = b"A" * 63
    with pytest.raises(ValueError, match="Archivo demasiado pequeño — posiblemente corrupto"):
        descifrar(short_payload, password)

@pytest.mark.skipif(not HAS_CRYPTOGRAPHY, reason="Requires real cryptography module")
def test_descifrar_invalid_magic_bytes_raises_value_error():
    """Test that valid-length payload but wrong magic bytes raises ValueError."""
    password = b"secret_password"
    # Create a 64 byte payload with invalid magic bytes ("RD25" instead of "RD26")
    invalid_magic_payload = b"RD25" + (b"A" * 60)
    with pytest.raises(ValueError, match="Formato inválido — archivo no reconocido"):
        descifrar(invalid_magic_payload, password)

@pytest.mark.skipif(not HAS_CRYPTOGRAPHY, reason="Requires real cryptography module")
def test_descifrar_roundtrip_success():
    """Verify that a normal round-trip encrypt and decrypt works."""
    password = b"secret_password"
    data = b"This is a super secret payload!"
    encrypted = cifrar(data, password)
    decrypted = descifrar(encrypted, password)
    assert decrypted == data

@pytest.mark.skipif(not HAS_CRYPTOGRAPHY, reason="Requires real cryptography module")
def test_descifrar_tampered_ciphertext_raises_value_error():
    """Test that tampering with the encrypted payload portion raises ValueError."""
    password = b"secret_password"
    data = b"This is a super secret payload!"
    encrypted = bytearray(cifrar(data, password))

    # Modify one byte in the ciphertext (which is before the last 16 bytes for the GCM tag)
    # The GCM tag is the last 16 bytes.
    # encrypted format: 4 (magic) + 32 (salt) + 12 (nonce) + ciphertext + 16 (tag)
    # the ciphertext starts at index 48. Let's tamper with the first byte of ciphertext.
    tamper_index = 48
    # Just to be sure there's enough ciphertext to tamper with
    assert len(encrypted) > 48 + 16

    # Flip one bit in the ciphertext byte
    encrypted[tamper_index] ^= 0x01

    with pytest.raises(ValueError, match="Autenticación fallida — archivo manipulado o clave incorrecta"):
        descifrar(bytes(encrypted), password)

@pytest.mark.skipif(not HAS_CRYPTOGRAPHY, reason="Requires real cryptography module")
def test_descifrar_tampered_gcm_tag_raises_value_error():
    """Test that tampering with the GCM tag raises ValueError."""
    password = b"secret_password"
    data = b"This is a super secret payload!"
    encrypted = bytearray(cifrar(data, password))

    # The GCM tag is the last 16 bytes. Tamper with the last byte.
    tamper_index = -1
    encrypted[tamper_index] ^= 0x01

    with pytest.raises(ValueError, match="Autenticación fallida — archivo manipulado o clave incorrecta"):
        descifrar(bytes(encrypted), password)

@pytest.mark.skipif(not HAS_CRYPTOGRAPHY, reason="Requires real cryptography module")
def test_descifrar_tampered_salt_raises_value_error():
    """Test that tampering with the salt raises ValueError."""
    password = b"secret_password"
    data = b"This is a super secret payload!"
    encrypted = bytearray(cifrar(data, password))

    # Salt is from index 4 to 36 (4 + 32). Tamper with the 5th byte (index 4).
    tamper_index = 4
    encrypted[tamper_index] ^= 0x01

    with pytest.raises(ValueError, match="Autenticación fallida — archivo manipulado o clave incorrecta"):
        descifrar(bytes(encrypted), password)

@pytest.mark.skipif(not HAS_CRYPTOGRAPHY, reason="Requires real cryptography module")
def test_descifrar_tampered_nonce_raises_value_error():
    """Test that tampering with the nonce raises ValueError."""
    password = b"secret_password"
    data = b"This is a super secret payload!"
    encrypted = bytearray(cifrar(data, password))

    # Nonce is from index 36 to 48 (4 + 32 + 12). Tamper with the first byte of nonce (index 36).
    tamper_index = 36
    encrypted[tamper_index] ^= 0x01

    with pytest.raises(ValueError, match="Autenticación fallida — archivo manipulado o clave incorrecta"):
        descifrar(bytes(encrypted), password)

@pytest.mark.skipif(not HAS_CRYPTOGRAPHY, reason="Requires real cryptography module")
def test_descifrar_incorrect_password_raises_value_error():
    """Test that using an incorrect password to decrypt raises ValueError."""
    password = b"secret_password"
    wrong_password = b"wrong_password"
    data = b"This is a super secret payload!"
    encrypted = cifrar(data, password)

    with pytest.raises(ValueError, match="Autenticación fallida — archivo manipulado o clave incorrecta"):
        descifrar(encrypted, wrong_password)
