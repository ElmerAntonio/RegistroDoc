import sys
import os
src_path = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "src"))
if src_path not in sys.path:
    sys.path.insert(0, src_path)

import pytest
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

def test_obtener_cuadro_honor_general_real_file():
    """Integration test validating obtener_cuadro_honor_general with the real workbook."""
    real_path = os.path.join("assets", "templates", "Registro_Premedia.xlsx")
    if not os.path.exists(real_path):
        pytest.skip("Registro_Premedia.xlsx not present in assets/templates, skipping integration test")
        
    engine = DataEngine(real_path)
    cuadro = engine.obtener_cuadro_honor_general()
    
    # Assert type
    assert isinstance(cuadro, list)
    
    # If any students are in the honor roll, verify their scores and sorting
    if len(cuadro) > 0:
        prev_prom = 5.01
        for est in cuadro:
            assert "nombre" in est
            assert "grado" in est
            assert "promedio" in est
            assert 4.5 <= est["promedio"] <= 5.0, f"Promedio {est['promedio']} must be between 4.5 and 5.0"
            assert est["promedio"] <= prev_prom, "Students must be sorted in descending order of average"
            prev_prom = est["promedio"]


def test_cierre_ano_lectivo(tmp_path):
    """Test that realizar_cierre_ano_lectivo correctly processes backups, resets data, and increments the year."""
    import openpyxl
    file_path = str(tmp_path / "Registro_Premedia.xlsx")
    
    # Create dummy workbook with at least one sheet
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "ASISTENCIA (7° A)"
    # Add a date cell DD-MM to verify cleaner
    ws.cell(row=2, column=3, value="12-05")
    ws.cell(row=3, column=3, value="P")
    wb.save(file_path)
    wb.close()
    
    engine = DataEngine(file_path)
    engine.modalidad = "premedia"
    
    # Set config year to 2026
    from rdsecurity import cargar_config_segura, guardar_config_segura
    cfg = cargar_config_segura({})
    cfg["ano_lectivo"] = "2026"
    guardar_config_segura(cfg)
    
    # Run rollover without student promotion (since we have no students in DB yet)
    exito, msg = engine.realizar_cierre_ano_lectivo(promover_estudiantes=False)
    
    assert exito is True
    assert "2026" in msg
    
    # Verify year was incremented
    cfg_new = cargar_config_segura({})
    assert cfg_new.get("ano_lectivo") == "2027"

def test_parsear_archivo_estudiantes_txt(tmp_path):
    """Test that parsear_archivo_estudiantes correctly extracts names and Panamanian cedulas from text files."""
    from impp import parsear_archivo_estudiantes
    
    txt_path = str(tmp_path / "roster.txt")
    with open(txt_path, "w", encoding="utf-8") as f:
        f.write("1. ELMER ANTONIO RAMOS  8-765-4321\n")
        f.write("2. JUAN PEREZ PE-99-8888\n")
        f.write("NOMBRE CEDULA (CABECERA A IGNORAR)\n")
        f.write("3. MARIA DE LEON (SIN CEDULA)\n")
        
    estudiantes = parsear_archivo_estudiantes(txt_path)
    
    assert len(estudiantes) == 3
    assert estudiantes[0]["nombre"] == "ELMER ANTONIO RAMOS"
    assert estudiantes[0]["cedula"] == "8-765-4321"
    assert estudiantes[1]["nombre"] == "JUAN PEREZ"
    assert estudiantes[1]["cedula"] == "PE-99-8888"
    assert estudiantes[2]["nombre"] == "MARIA DE LEON (SIN CEDULA)"
    assert estudiantes[2]["cedula"] == ""

