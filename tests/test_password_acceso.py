import os
import sys

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "src")))

from rdsecurity import hash_password, verify_password


def test_password_correcta_verifica():
    salt, h = hash_password("MiClave123")
    assert verify_password("MiClave123", salt, h) is True


def test_password_incorrecta_falla():
    salt, h = hash_password("MiClave123")
    assert verify_password("otra", salt, h) is False
    assert verify_password("", salt, h) is False
    assert verify_password("miclave123", salt, h) is False  # sensible a mayúsculas


def test_salts_distintos_por_llamada():
    salt1, h1 = hash_password("igual")
    salt2, h2 = hash_password("igual")
    assert salt1 != salt2   # salt aleatorio por llamada
    assert h1 != h2         # → hash distinto aunque la contraseña sea la misma
    # pero ambas verifican la misma contraseña
    assert verify_password("igual", salt1, h1) is True
    assert verify_password("igual", salt2, h2) is True


def test_hash_no_contiene_la_password():
    salt, h = hash_password("SecretoDelDocente")
    assert "SecretoDelDocente" not in h
    assert "SecretoDelDocente" not in salt


def test_verify_robusto_ante_datos_invalidos():
    # No debe lanzar excepción con entradas corruptas
    assert verify_password("x", "no-hex", "no-hex") is False
    assert verify_password("x", "", "") is False
