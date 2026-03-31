import pytest
from unittest.mock import patch
from src.rdsecurity import _derivar_clave

def test_derivar_clave_kat():
    """
    Known-Answer Test (KAT) for _derivar_clave.
    Ensures that for a specific password, salt, and PBKDF2 iterations,
    the output is strictly deterministic.
    Since PBKDF2_ITERS could change, we mock it for this specific test
    to guarantee the KAT remains stable and isolates testing to the mathematical
    logic and backend configuration.
    """
    password = b"password123"
    salt = b"somesalt456"
    expected_key_hex = "9774da0e83619a1228bf30670b4dcb0a9369c955b95e9f70c3355eee1faacbb7"

    with patch('src.rdsecurity.PBKDF2_ITERS', 1000):
        derived_key = _derivar_clave(password, salt)
        assert derived_key.hex() == expected_key_hex

def test_derivar_clave_empty_password():
    """Test behavior with empty password."""
    salt = b"randomsalt"

    with patch('src.rdsecurity.PBKDF2_ITERS', 1000):
        # Empty password is mathematically valid in PBKDF2
        derived_key = _derivar_clave(b"", salt)
        assert len(derived_key) == 32

def test_derivar_clave_empty_salt():
    """Test behavior with empty salt."""
    password = b"password123"

    with patch('src.rdsecurity.PBKDF2_ITERS', 1000):
        # Empty salt is mathematically valid in PBKDF2
        derived_key = _derivar_clave(password, b"")
        assert len(derived_key) == 32

def test_derivar_clave_invalid_types():
    """Test behavior with invalid input types."""
    salt = b"somesalt"
    password = b"password123"

    with patch('src.rdsecurity.PBKDF2_ITERS', 1000):
        # Passing strings instead of bytes should raise TypeError
        with pytest.raises(TypeError):
            _derivar_clave("string_password", salt)

        with pytest.raises(TypeError):
            _derivar_clave(password, "string_salt")
