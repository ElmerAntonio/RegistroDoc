import sys
import cryptography
import pytest
from src.rdsecurity import validar_nota_meduca
from src.rdsecurity import generar_codigo_licencia

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


def test_generar_codigo_licencia_determinism():
    """Test that generating a code for the same cedula always yields the exact same code."""
    cedula = "8-765-4321"
    codigo1 = generar_codigo_licencia(cedula)
    codigo2 = generar_codigo_licencia(cedula)
    codigo3 = generar_codigo_licencia(cedula)

    assert codigo1 == codigo2
    assert codigo1 == codigo3

def test_generar_codigo_licencia_format():
    """Test that the generated code has the correct length and format."""
    cedula = "4-123-456"
    codigo = generar_codigo_licencia(cedula)

    # Check length
    assert len(codigo) == 26

    # Check format using regular expression
    assert bool(re.match(r"^RD-[A-Z0-9]{5}-[A-Z0-9]{5}-[A-Z0-9]{5}-[A-Z0-9]{5}$", codigo))

def test_generar_codigo_licencia_uniqueness():
    """Test that different cedulas generate completely different codes."""
    cedula1 = "8-111-222"
    cedula2 = "8-111-223"

    codigo1 = generar_codigo_licencia(cedula1)
    codigo2 = generar_codigo_licencia(cedula2)

    assert codigo1 != codigo2

def test_generar_codigo_licencia_whitespace_handling():
    """Test that leading and trailing whitespaces in the cedula are ignored."""
    cedula = " 3-888-999 "
    codigo_with_spaces = generar_codigo_licencia(cedula)
    codigo_without_spaces = generar_codigo_licencia(cedula.strip())

    assert codigo_with_spaces == codigo_without_spaces

def test_generar_codigo_licencia_empty_string():
    """Test generating a code for an empty string (should still work deterministically)."""
    cedula = ""
    codigo = generar_codigo_licencia(cedula)

    assert len(codigo) == 26
    assert bool(re.match(r"^RD-[A-Z0-9]{5}-[A-Z0-9]{5}-[A-Z0-9]{5}-[A-Z0-9]{5}$", codigo))
