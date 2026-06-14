import sys
import os

# Resolviendo directorios para linter e importaciones
src_path = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "src"))
if src_path not in sys.path:
    sys.path.insert(0, src_path)

import pytest
import sqlite3
from rdsql import SQLDatabaseManager, obtener_dir_datos_usuario

@pytest.fixture
def clean_db_manager(tmp_path):
    # Usar una llave de prueba
    test_key = b"RD_TEST_KEY_32_BYTES_FOR_SECURITY"
    manager = SQLDatabaseManager(key=test_key)
    
    # Redireccionar las rutas al directorio temporal del test para evitar colisión de bloqueos
    manager.db_enc_path = str(tmp_path / "test_registro.db.enc")
    manager.db_temp_path = str(tmp_path / "test_sqlite_temp.db")
    
    # Asegurar que no hay residuos anteriores
    if os.path.exists(manager.db_enc_path):
        os.remove(manager.db_enc_path)
    if os.path.exists(manager.db_temp_path):
        os.remove(manager.db_temp_path)
        
    yield manager
    
    # Limpieza post-test
    manager.cerrar()
    if os.path.exists(manager.db_enc_path):
        os.remove(manager.db_enc_path)
    if os.path.exists(manager.db_temp_path):
        try:
            os.remove(manager.db_temp_path)
        except Exception:
            pass

def test_database_initialization_and_table_creation(clean_db_manager):
    manager = clean_db_manager
    conn = manager.conectar()
    
    assert isinstance(conn, sqlite3.Connection)
    
    # Verificar que las tablas existen
    cursor = conn.cursor()
    cursor.execute("SELECT name FROM sqlite_master WHERE type='table';")
    tables = [r[0] for r in cursor.fetchall()]
    
    expected_tables = [
        "configuracion", "grados", "estudiantes", "materias", 
        "horario", "notas", "asistencia", "observaciones", 
        "habitos", "tareas", "auditoria"
    ]
    for table in expected_tables:
        assert table in tables

def test_database_encryption_cycle_and_persistence(clean_db_manager):
    manager = clean_db_manager
    
    # 1. Conectar e insertar datos
    conn = manager.conectar()
    cursor = conn.cursor()
    cursor.execute("INSERT INTO configuracion (clave, valor) VALUES (?, ?);", ("docente_nombre", "Elmer Antonio"))
    conn.commit()
    
    # 2. Cerrar debe cifrar la base y borrar el temp_db
    manager.cerrar()
    assert not os.path.exists(manager.db_temp_path)
    assert os.path.exists(manager.db_enc_path)
    
    # 3. Validar que el archivo cifrado contiene el magic header de rdsecurity (RD26)
    with open(manager.db_enc_path, "rb") as f:
        magic = f.read(4)
    assert magic == b"RD26"
    
    # 4. Reconectar con el mismo manager
    conn2 = manager.conectar()
    cursor2 = conn2.cursor()
    cursor2.execute("SELECT valor FROM configuracion WHERE clave = ?;", ("docente_nombre",))
    res = cursor2.fetchone()
    
    assert res is not None
    assert res[0] == "Elmer Antonio"
    
    manager.cerrar()
