import pytest
import os
import tempfile
import sys
from unittest.mock import patch, MagicMock

# Pre-import setup
os.environ["REGISTRODOC_MASTER_SALT"] = "TEST_SALT_EXTENDED"

# Add src to sys.path
src_path = os.path.abspath(os.path.join(os.path.dirname(__file__), '../src'))
if src_path not in sys.path:
    sys.path.insert(0, src_path)

import rdsecurity

# Redirect files to temp dir
TEMP_DIR = tempfile.mkdtemp()
rdsecurity.CONFIG_FILE = os.path.join(TEMP_DIR, "config.enc")
rdsecurity.TOKEN_FILE = os.path.join(TEMP_DIR, "rd_token.bin")
rdsecurity.LICENSE_FILE = os.path.join(TEMP_DIR, "licencia.dat")
rdsecurity.BRUTE_FILE = os.path.join(TEMP_DIR, "rd_brute.bin")
rdsecurity.AUDIT_FILE = os.path.join(TEMP_DIR, "rd_audit.bin")
rdsecurity.EXCEL_HASH_FILE = os.path.join(TEMP_DIR, "rd_excel_hash.bin")

from rdsecurity import (
    registrar_intento_fallido, verificar_bloqueo, resetear_intentos,
    activar_licencia_completa, verificar_licencia, generar_codigo_licencia
)

def setup_function():
    # Clear temp files before each test
    for f in [rdsecurity.CONFIG_FILE, rdsecurity.TOKEN_FILE, rdsecurity.LICENSE_FILE, rdsecurity.BRUTE_FILE, rdsecurity.AUDIT_FILE, rdsecurity.EXCEL_HASH_FILE]:
        if os.path.exists(f):
            try: os.remove(f)
            except: pass

def test_brute_force_protection():
    # Initially not blocked
    bloqueado, msg = verificar_bloqueo()
    assert not bloqueado

    # 1. First 2 failures: no lock
    for i in range(1, 3):
        bloq, msg = registrar_intento_fallido()
        assert not bloq
        assert "incorrecto" in msg

    # 2. 3rd failure: 30s lock
    bloq, msg = registrar_intento_fallido()
    assert bloq
    assert "30 segundos" in msg

    bloqueado, msg_verif = verificar_bloqueo()
    assert bloqueado
    assert "TEMPORALMENTE" in msg_verif

    # 3. Reset attempts
    resetear_intentos()
    bloqueado, msg_verif = verificar_bloqueo()
    assert not bloqueado

def test_license_activation_flow():
    cedula = "8-999-999"
    codigo_correcto = generar_codigo_licencia(cedula)
    codigo_incorrecto = "RD-AAAAA-BBBBB-CCCCC-DDDDD"

    # Mock _hw_fingerprint to be stable
    with patch('rdsecurity._hw_fingerprint', return_value=b"STABLE_HW_123"):
        # Fails with wrong code
        ok, msg = activar_licencia_completa(cedula, codigo_incorrecto)
        assert not ok
        assert "incorrecto" in msg

        # Passes with correct code
        ok, msg = activar_licencia_completa(cedula, codigo_correcto)
        assert ok
        assert "exitosa" in msg

        # Verification succeeds
        valida, datos = verificar_licencia()
        assert valida
        assert datos["cedula"] == cedula

def test_license_hardware_mismatch():
    cedula = "8-777-777"
    codigo = generar_codigo_licencia(cedula)

    # 1. Activate on HW_A
    with patch('rdsecurity._hw_fingerprint', return_value=b"HW_SIGNATURE_A"):
        ok, msg = activar_licencia_completa(cedula, codigo)
        assert ok

    # 2. Try to verify on HW_B -> should fail mismatch check
    with patch('rdsecurity._hw_fingerprint', return_value=b"HW_SIGNATURE_B"):
        valida, datos = verificar_licencia()
        assert not valida
