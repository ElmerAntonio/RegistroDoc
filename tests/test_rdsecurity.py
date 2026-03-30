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

import os
import pytest

# Mock environment variable for tests
os.environ["REGISTRODOC_MASTER_SALT"] = "test_salt_for_pytest"

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

def test_missing_master_salt():
    """Test that importing rdsecurity without REGISTRODOC_MASTER_SALT raises an error."""
    import sys
    import importlib

    # Store old environment variable
    old_val = os.environ.get("REGISTRODOC_MASTER_SALT")

    try:
        # Remove from env
        if "REGISTRODOC_MASTER_SALT" in os.environ:
            del os.environ["REGISTRODOC_MASTER_SALT"]

        # Remove from sys.modules to force reload
        if "src.rdsecurity" in sys.modules:
            del sys.modules["src.rdsecurity"]

        # Import should fail
        with pytest.raises(RuntimeError) as excinfo:
            import src.rdsecurity

        assert "CRITICAL ERROR: REGISTRODOC_MASTER_SALT no encontrada" in str(excinfo.value)

    finally:
        # Restore env
        if old_val is not None:
            os.environ["REGISTRODOC_MASTER_SALT"] = old_val

        # Re-import for other tests just in case
        if "src.rdsecurity" in sys.modules:
            del sys.modules["src.rdsecurity"]
        import src.rdsecurity
