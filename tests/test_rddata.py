import pytest
import os
from rddata import DataEngine

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

def test_obtener_grados_activos_missing_file():
    """
    Test that obtener_grados_activos returns an empty list
    when initialized with a non-existent file path.
    """
    # Create an instance with a non-existent path
    engine = DataEngine(ruta_excel="ruta_inexistente_grados.xlsx")

    # Call the method
    resultado = engine.obtener_grados_activos()

    # Expected empty list
    assert resultado == []

def test_atomic_save(tmp_path):
    import openpyxl
    file_path = str(tmp_path / "test_libreta.xlsx")
    wb = openpyxl.Workbook()
    wb.save(file_path)
    wb.close()
    
    engine = DataEngine(file_path)
    
    # 1. Modificar y guardar de forma atómica
    wb_write = openpyxl.load_workbook(file_path)
    wb_write.active.cell(row=1, column=1, value="NUEVO_VALOR")
    engine._save_wb(wb_write)
    wb_write.close()
    
    # 2. Comprobar que el original fue guardado y el backup se creó
    bak_path = file_path.replace(".xlsx", "_bak.xlsx")
    assert os.path.exists(file_path)
    assert os.path.exists(bak_path)
    
    # 3. Comprobar contenido guardado
    wb_check = openpyxl.load_workbook(file_path)
    assert wb_check.active.cell(row=1, column=1).value == "NUEVO_VALOR"
    wb_check.close()

def test_cache_invalidation_mtime(tmp_path):
    import openpyxl
    import time
    file_path = str(tmp_path / "test_cache.xlsx")
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "DASHBOARD"
    ws.cell(row=1, column=1, value="ORIGINAL")
    wb.save(file_path)
    wb.close()
    
    engine = DataEngine(file_path)
    assert engine._wb_cache.active.cell(row=1, column=1).value == "ORIGINAL"
    
    # Simular edición externa escribiendo al archivo
    time.sleep(0.01)  # Asegurar cambio de mtime en sistemas rápidos
    wb_ext = openpyxl.load_workbook(file_path)
    wb_ext.active.cell(row=1, column=1, value="EDITADO_EXTERNO")
    wb_ext.save(file_path)
    wb_ext.close()
    
    # El motor debería invalidar y recargar al comprobar
    engine._verificar_y_recargar_cache()
    assert engine._wb_cache.active.cell(row=1, column=1).value == "EDITADO_EXTERNO"
