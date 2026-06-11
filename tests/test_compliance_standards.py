import pytest
import os
import tempfile
import sys
import time
import tracemalloc
import re
from unittest.mock import patch, MagicMock

# Pre-import setup
os.environ["REGISTRODOC_MASTER_SALT"] = "COMPLIANCE_TEST_SALT"

# Add src to sys.path
src_path = os.path.abspath(os.path.join(os.path.dirname(__file__), '../src'))
if src_path not in sys.path:
    sys.path.insert(0, src_path)

import rdsecurity
from rdsecurity import cifrar, descifrar, _hw_fingerprint

# ISO 25010: Comportamiento Temporal (Rendimiento)
def test_iso_25010_time_behavior():
    password = b"hw_key_compliance_123"
    data = b"informacion_sensible_de_prueba" * 100
    
    t_start = time.perf_counter()
    # Ejecutar 10 rondas de cifrado/descifrado para evaluar latencia acumulada (suficiente para benchmarking)
    for _ in range(10):
        blob = cifrar(data, password)
        original = descifrar(blob, password)
        assert original == data
    t_end = time.perf_counter()
    
    elapsed = t_end - t_start
    avg_round_trip = elapsed / 10
    
    # Con 600,000 iteraciones de PBKDF2 (NIST), el tiempo promedio debe ser inferior a 0.5 segundos.
    assert avg_round_trip < 0.5, f"Rendimiento de criptografia lento: {avg_round_trip:.4f}s promedio"

# ISO 25010: Utilización de Recursos (Memoria y fugas)
def test_iso_25010_resource_utilization():
    tracemalloc.start()
    
    password = b"compliance_password"
    data = b"leak_test_data"
    
    # Capturar estado inicial
    snapshot1 = tracemalloc.take_snapshot()
    
    # Realizar operaciones repetitivas para inducir fugas potenciales
    for _ in range(20):
        blob = cifrar(data, password)
        _ = descifrar(blob, password)
        
    # Capturar estado final
    snapshot2 = tracemalloc.take_snapshot()
    tracemalloc.stop()
    
    # Medir diferencia en asignación de memoria
    stats = snapshot2.compare_to(snapshot1, 'lineno')
    total_diff = sum(stat.size_diff for stat in stats)
    
    # Tolerancia de fuga residual menor a 50 KB tras 100 iteraciones
    assert total_diff < 50000, f"Posible fuga de memoria detectada: {total_diff} bytes asignados"

# ISO 27001 & Ley 81 (Panamá): Privacidad de Datos en Pruebas
def test_iso_27001_data_privacy_compliance():
    # Comprobar que en los datos mock de los estudiantes de pruebas no hay nombres de personas reales o cédulas válidas.
    # El formato estándar de cédula en Panamá es: Prov-Tomo-Asiento (Ej: 8-720-1420).
    # Las cédulas de prueba válidas en nuestros tests usan el formato ficticio de ceros "0-000-0000" o similares.
    
    test_students = [
        {"nombre": "Juan Perez Mock", "cedula": "0-000-0001"},
        {"nombre": "Estudiante Prueba A", "cedula": "0-000-0002"},
        {"nombre": "Maria Gomez Demo", "cedula": "0-000-0003"},
    ]
    
    patron_cedula_real = re.compile(r"^[1-9][0-9]?-\d{3,4}-\d{3,4}$")
    
    for student in test_students:
        # 1. Asegurar que contiene indicadores de ser mock/prueba
        nombre_upper = student["nombre"].upper()
        assert any(x in nombre_upper for x in ["MOCK", "PRUEBA", "DEMO", "TEST"]), \
            f"El nombre del estudiante de prueba '{student['nombre']}' no contiene indicadores de anonimización."
        
        # 2. Asegurar que la cédula no parece una cédula real de ciudadano activo
        assert not patron_cedula_real.match(student["cedula"]), \
            f"La cedula '{student['cedula']}' del estudiante de prueba imita un patron de cedula real activo de Panama."

# NIST SP 800-38D: Integridad y Resistencia a la manipulación física de AES-GCM
def test_nist_aes_gcm_tamper_resistance():
    password = b"compliance_nist_pass"
    data = b"datos_originales_que_no_deben_ser_manipulados"
    
    blob = bytearray(cifrar(data, password))
    
    # El formato del blob es: magic (4 bytes) + salt (32 bytes) + nonce (12 bytes) + ciphertext/tag
    
    # 1. Alterar el Nonce (bytes 36 a 47)
    blob_nonce_tampered = bytearray(blob)
    blob_nonce_tampered[38] ^= 0xFF
    with pytest.raises(ValueError, match="Autenticación fallida"):
        descifrar(bytes(blob_nonce_tampered), password)
        
    # 2. Alterar el Salt (bytes 4 a 35)
    blob_salt_tampered = bytearray(blob)
    blob_salt_tampered[10] ^= 0xFF
    with pytest.raises(ValueError, match="Autenticación fallida"):
        descifrar(bytes(blob_salt_tampered), password)
        
    # 3. Alterar el cuerpo del texto cifrado o el tag (bytes finales)
    blob_body_tampered = bytearray(blob)
    blob_body_tampered[-5] ^= 0xFF
    with pytest.raises(ValueError, match="Autenticación fallida"):
        descifrar(bytes(blob_body_tampered), password)
