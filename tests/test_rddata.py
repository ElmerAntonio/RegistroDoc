import pytest
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
