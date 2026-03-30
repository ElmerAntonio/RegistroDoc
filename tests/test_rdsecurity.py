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

def test_cifrar_descifrar_happy_path():
    """Test successful encryption and decryption flow."""
    datos = b"Datos secretos de prueba para AES-256-GCM"
    password = b"MiSuperClaveSecreta123"

    # Encrypt
    blob = cifrar(datos, password)

    # Verify blob structure briefly
    assert blob.startswith(b"RD26"), "Blob must start with magic bytes RD26"
    assert len(blob) > 4 + 32 + 12 + 16, "Blob is too small"

    # Decrypt
    resultado = descifrar(blob, password)

    # Ensure decrypted data exactly matches original data
    assert resultado == datos, "Decrypted data must exactly match original data"


def test_descifrar_incorrect_password():
    """Test decryption fails with incorrect password."""
    datos = b"Datos confidenciales"
    password = b"ClaveCorrecta!"
    password_incorrecta = b"ClaveIncorrecta!"

    # Encrypt with correct password
    blob = cifrar(datos, password)

    # Attempt to decrypt with wrong password
    with pytest.raises(ValueError, match="Autenticación fallida — archivo manipulado o clave incorrecta"):
        descifrar(blob, password_incorrecta)


def test_descifrar_integrity_failure_modified_blob():
    """Test decryption fails when the blob has been modified (tampered)."""
    datos = b"Datos financieros ultra secretos"
    password = b"ClaveParaFinanzas#2026"

    # Encrypt data normally
    blob = cifrar(datos, password)

    # Convert to mutable bytearray
    blob_mut = bytearray(blob)

    # Intentionally flip a bit in the ciphertext/tag area (last byte)
    # This simulates corruption or a hacker trying to manipulate the encrypted file
    blob_mut[-1] ^= 0x01

    blob_modificado = bytes(blob_mut)

    # Attempt to decrypt the modified blob with the correct password
    # The AES-GCM tag validation should catch the modification and fail
    with pytest.raises(ValueError, match="Autenticación fallida — archivo manipulado o clave incorrecta"):
        descifrar(blob_modificado, password)


def test_descifrar_invalid_blob_format():
    """Test decryption fails gracefully when the blob format is completely invalid."""
    password = b"CualquierClave"

    # Too small blob
    blob_pequeno = b"RD26" + (b"A" * 10)  # less than 4+32+12+16 bytes
    with pytest.raises(ValueError, match="Archivo demasiado pequeño — posiblemente corrupto"):
        descifrar(blob_pequeno, password)

    # Invalid magic bytes but correct size
    blob_invalido_magic = b"RD25" + (b"A" * (32 + 12 + 16 + 10))
    with pytest.raises(ValueError, match="Formato inválido — archivo no reconocido"):
        descifrar(blob_invalido_magic, password)
