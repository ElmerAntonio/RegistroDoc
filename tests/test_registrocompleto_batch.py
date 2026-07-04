import os
import sys
import pytest
from unittest.mock import MagicMock, patch

# Resolviendo directorios para linter e importaciones
src_path = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "src"))
if src_path not in sys.path:
    sys.path.insert(0, src_path)

# Mock standard Tkinter/CTk classes to allow importing and instantiating
# RegistroCompletoFrame without running GUI loop
@pytest.fixture(autouse=True)
def mock_ctk_and_gui():
    with patch('customtkinter.CTkFrame.__init__', lambda *args, **kwargs: None), \
         patch('customtkinter.CTkFrame.grid_columnconfigure', lambda *args, **kwargs: None), \
         patch('customtkinter.CTkFrame.grid_rowconfigure', lambda *args, **kwargs: None), \
         patch('registrocompletoapp.RegistroCompletoFrame.crear_filtros', lambda *args, **kwargs: None), \
         patch('registrocompletoapp.RegistroCompletoFrame.crear_tabs', lambda *args, **kwargs: None), \
         patch('registrocompletoapp.RegistroCompletoFrame.actualizar_vista', lambda *args, **kwargs: None):
        yield

def test_obtener_datos_de_excel_batch(tmp_path):
    from rdsql import SQLDatabaseManager
    from registrocompletoapp import RegistroCompletoFrame
    
    # 1. Setup temporary database using clean_db_manager logic
    db_manager = SQLDatabaseManager()
    
    conn = db_manager.conectar()
    cursor = conn.cursor()
    
    # Insert mock grades
    cursor.execute("INSERT INTO grados (id, nombre, seccion, modalidad) VALUES (?, ?, ?, ?);", (10, "8° B", "A", "premedia"))
    
    # Insert mock students
    cursor.execute("INSERT INTO estudiantes (id, nombre, grado_id) VALUES (?, ?, ?);", ("1", "Estudiante Uno", 10))
    cursor.execute("INSERT INTO estudiantes (id, nombre, grado_id) VALUES (?, ?, ?);", ("2", "Estudiante Dos", 10))
    
    # Insert mock materias
    cursor.execute("INSERT INTO materias (id, nombre, grado_id) VALUES (?, ?, ?);", (101, "Español", 10))
    
    # Insert mock grades
    # Columns in SQL: estudiante_id, materia_id, trimestre, tipo, descripcion, valor, puntos_obtenidos, puntos_maximos, fecha
    cursor.execute("""
        INSERT INTO notas (estudiante_id, materia_id, trimestre, tipo, descripcion, valor, fecha)
        VALUES (?, ?, ?, ?, ?, ?, ?);
    """, ("1", 101, 1, "Examen", "Examen 1", 4.5, "2023-10-10"))
    cursor.execute("""
        INSERT INTO notas (estudiante_id, materia_id, trimestre, tipo, descripcion, valor, fecha)
        VALUES (?, ?, ?, ?, ?, ?, ?);
    """, ("2", 101, 1, "Examen", "Examen 1", 3.2, "2023-10-10"))
    
    # Insert mock attendance
    # Columns: estudiante_id, fecha, estado, trimestre
    cursor.execute("""
        INSERT INTO asistencia (estudiante_id, fecha, estado, trimestre)
        VALUES (?, ?, ?, ?);
    """, ("1", "2023-10-10", ".", 1))
    cursor.execute("""
        INSERT INTO asistencia (estudiante_id, fecha, estado, trimestre)
        VALUES (?, ?, ?, ?);
    """, ("2", "2023-10-10", "-", 1))
    
    # Insert mock habits
    # Columns: estudiante_id, trimestre, criterio_codigo, nota, frecuencia
    cursor.execute("""
        INSERT INTO habitos (estudiante_id, trimestre, criterio_codigo, nota, frecuencia)
        VALUES (?, ?, ?, ?, ?);
    """, ("1", 1, "Responsabilidad", "S", "Trimestral"))
    cursor.execute("""
        INSERT INTO habitos (estudiante_id, trimestre, criterio_codigo, nota, frecuencia)
        VALUES (?, ?, ?, ?, ?);
    """, ("2", 1, "Responsabilidad", "R", "Trimestral"))
    
    conn.commit()
    
    print("DEBUG - grados:", cursor.execute("SELECT * FROM grados").fetchall())
    print("DEBUG - materias:", cursor.execute("SELECT * FROM materias").fetchall())
    print("DEBUG - estudiantes:", cursor.execute("SELECT * FROM estudiantes").fetchall())
    print("DEBUG - notas:", cursor.execute("SELECT * FROM notas").fetchall())
    
    # Mock DataEngine and its methods
    mock_engine = MagicMock()
    mock_engine.db_manager = db_manager
    mock_engine.conectar_db.return_value = conn
    mock_engine.db_conn = conn
    mock_engine.modalidad = "premedia"
    
    mock_engine.obtener_estudiantes_completos.return_value = [
        {"id": 1, "nombre": "Estudiante Uno"},
        {"id": 2, "nombre": "Estudiante Dos"}
    ]
    mock_engine.obtener_promedios_reales.return_value = {
        "Estudiante Uno": 4.5,
        "Estudiante Dos": 3.2
    }
    mock_engine.obtener_fechas_asistencia.return_value = ["2023-10-10"]
    
    # Instantiate app/frame
    frame = RegistroCompletoFrame(master=None, engine=mock_engine)
    
    # Mock tasks list for grades
    frame.combo_materia = MagicMock()
    frame.combo_materia.get.return_value = "Español"
    
    # Call _obtener_datos_de_excel
    # params: grado, trimestre, materia
    result = frame._obtener_datos_de_excel("8° B", "Trimestre 1", "Español")
    
    print("DEBUG - result:", result)
    
    # Assert result structure and correctness
    assert result is not None
    
    # 1. Calificaciones assertions
    assert len(result["headers_notas"]) == 1
    assert result["headers_notas"][0][0] == "Examen 1\n(2023-10-10)"
    assert len(result["rows_notas"]) == 2
    assert result["rows_notas"][0]["nombre"] == "Estudiante Uno"
    assert result["rows_notas"][0]["notas"] == [4.5]
    assert result["rows_notas"][0]["promedio"] == 4.5
    assert result["rows_notas"][1]["notas"] == [3.2]
    
    # 2. Asistencia assertions
    assert result["headers_asis"] == ["2023-10-10"]
    assert len(result["rows_asis"]) == 2
    assert result["rows_asis"][0]["asistencia"] == ["P"]
    assert result["rows_asis"][0]["porcentaje"] == 100.0
    assert result["rows_asis"][1]["asistencia"] == ["A"]
    assert result["rows_asis"][1]["porcentaje"] == 0.0
    
    # 3. Hábitos assertions
    assert "Responsabilidad" in result["habitos_criterios"]
    assert len(result["rows_habitos"]) == 2
    assert result["rows_habitos"][0]["evals_actual"]["Responsabilidad"]["final"] == "S"
    assert result["rows_habitos"][1]["evals_actual"]["Responsabilidad"]["final"] == "R"
    
    # Cleanup DB
    db_manager.cerrar()
