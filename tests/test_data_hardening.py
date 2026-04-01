import pytest
from rddata import DataEngine
import os

@pytest.fixture
def dummy_engine(tmp_path):
    import openpyxl
    file_path = tmp_path / "dummy_registro.xlsx"
    wb = openpyxl.Workbook()
    ws_m = wb.active
    ws_m.title = "MAESTRO"
    wb.save(file_path)
    # mock config methods
    engine = DataEngine(str(file_path), "primaria")
    # patch the limit to check our custom implementation
    return engine

def test_student_limit_primaria(dummy_engine):
    # Mocking obtener_estudiantes_completos
    dummy_engine.obtener_estudiantes_completos = lambda grado: [{"id": i} for i in range(34)]

    # Try adding 35th student
    res = dummy_engine.agregar_estudiante("7°", "NEW STUDENT")
    assert res == False, "Engine should prevent adding students beyond limit"

def test_student_limit_premedia(dummy_engine):
    dummy_engine.modalidad = "premedia"
    dummy_engine.obtener_estudiantes_completos = lambda grado: [{"id": i} for i in range(36)]

    res = dummy_engine.agregar_estudiante("7°", "NEW STUDENT")
    assert res == False, "Engine should prevent adding students beyond limit"
