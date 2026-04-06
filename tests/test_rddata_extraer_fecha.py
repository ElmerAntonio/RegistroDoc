import pytest
from rddata import DataEngine

@pytest.mark.parametrize("input_val, expected", [
    ("ASISTENCIA DEL 15 DE MARZO", "15 DE MARZO"),
    ("asistencia del 10 de abril", "10 DE ABRIL"),
    ("  ASISTENCIA DEL   20 DE MAYO  ", "20 DE MAYO"),
    ("ASISTENCIA DEL 12_DE_JUNIO", "12 DE JUNIO"),
    ("ASISTENCIA DEL 12__DE__JUNIO", "12 DE JUNIO"),
    ("OTRO TEXTO CUALQUIERA", ""),
    ("", ""),
    (None, ""),
    (12345, ""),
    ("ASISTENCIA DEL", ""),
    ("ASISTENCIA DEL   ", ""),
])
def test_procesar_texto_asistencia(input_val, expected):
    """Prueba la lógica de extracción de fecha de asistencia con diversos casos."""
    assert DataEngine._procesar_texto_asistencia(input_val) == expected
