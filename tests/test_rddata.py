import pytest
import os
from src.rddata import DataEngine

def test_obtener_horario_returns_default_structure():
    """Test that obtener_horario returns the default structure when the file doesn't exist."""
    # Act
    # We use a dummy path that clearly doesn't exist to test the edge case
    engine = DataEngine("ruta_horario_inexistente.xlsx")
    horario = engine.obtener_horario()

    # Assert
    expected_structure = [
        {"horas": "", "lunes": "", "martes": "", "miercoles": "", "jueves": "", "viernes": ""}
        for _ in range(8)
    ]

    assert horario == expected_structure
    assert len(horario) == 8

def test_obtener_datos_generales_missing_file():
    """
    Test that obtener_datos_generales returns the exact expected default
    dictionary when initialized with a non-existent file path.
    """
    # Create an instance with a non-existent path
    engine = DataEngine(ruta_excel="ruta_ficticia_inexistente.xlsx")

    # Ensure the path really doesn't exist
    assert not os.path.exists(engine.ruta), "El archivo no debe existir para esta prueba"

    # Call the method
    resultado = engine.obtener_datos_generales()

    # Expected default dictionary
    expected_datos = {
        "docente_nombre": "", "docente_cedula": "", "seguro_social": "", "numero_posicion": "",
        "condicion_nombramiento": "", "escuela_nombre": "", "escuela_region": "", "distrito": "",
        "corregimiento": "", "zona_escolar": "", "director_nombre": "", "subdirector_nombre": "",
        "coordinador_nombre": "", "telefono": "", "correo": "", "ano_lectivo": "2026",
        "jornada": "", "fecha_t1": "", "fecha_t2": "", "fecha_t3": ""
    }

    # Strict equality assertion
    assert resultado == expected_datos
