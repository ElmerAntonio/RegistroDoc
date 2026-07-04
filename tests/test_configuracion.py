import pytest
import os
import tempfile
import json
import sqlite3
from unittest.mock import patch, MagicMock

# Configurar variables de entorno antes de importar módulos
os.environ["REGISTRODOC_MASTER_SALT"] = "TEST_SALT"
os.environ["REGISTRODOC_DEV_MODE"] = "1"  # Permitir conexión SQLite en pruebas

import rdsecurity
from rdsecurity import guardar_config_segura, cargar_config_segura
import config
from rdsql import SQLDatabaseManager
from rddata import DataEngine

# Configurar directorio temporal para redirigir archivos cifrados y base de datos
TEMP_DIR = tempfile.mkdtemp()
rdsecurity.CONFIG_FILE = os.path.join(TEMP_DIR, "config.enc")
rdsecurity.TOKEN_FILE = os.path.join(TEMP_DIR, "rd_token.bin")
config.CONFIG_FILE = os.path.join(TEMP_DIR, "perfil.json")

# Redirigir bases de datos SQLite en SQLDatabaseManager
ORIG_INIT = SQLDatabaseManager.__init__

def mock_sql_init(self, key=None):
    self.db_dir = os.path.join(TEMP_DIR, "data")
    os.makedirs(self.db_dir, exist_ok=True)
    self.db_enc_path = os.path.join(self.db_dir, "registro.db.enc")
    
    self.temp_dir = os.path.join(TEMP_DIR, "temp")
    os.makedirs(self.temp_dir, exist_ok=True)
    self.db_temp_path = os.path.join(self.temp_dir, "sqlite_temp.db")
    
    import threading
    self.db_lock = threading.RLock()
    self.conn = None
    self.raw_hw_key = key if key else b"MOCK_HW_123"
    self.key = self.raw_hw_key
    
    # Crear estructura básica si no existe en la db temporal o cifrada
    if not os.path.exists(self.db_temp_path) and not os.path.exists(self.db_enc_path):
        conn = sqlite3.connect(self.db_temp_path)
        conn.execute("CREATE TABLE configuracion (clave TEXT PRIMARY KEY, valor TEXT NOT NULL);")
        conn.commit()
        conn.close()

# Parchear constructor y métodos de obtención de rutas
@pytest.fixture(autouse=True)
def setup_mocks(monkeypatch):
    monkeypatch.setattr(SQLDatabaseManager, "__init__", mock_sql_init)
    monkeypatch.setattr(rdsecurity, "obtener_dir_datos_usuario", lambda: TEMP_DIR)
    
    # Asegurar limpieza de archivos temporales entre pruebas
    for f in [rdsecurity.CONFIG_FILE, rdsecurity.TOKEN_FILE, config.CONFIG_FILE]:
        if os.path.exists(f):
            try: os.remove(f)
            except: pass
            
    db_enc = os.path.join(TEMP_DIR, "data", "registro.db.enc")
    db_temp = os.path.join(TEMP_DIR, "temp", "sqlite_temp.db")
    for f in [db_enc, db_temp]:
        if os.path.exists(f):
            try: os.remove(f)
            except: pass

@patch('rdsecurity._hw_fingerprint', return_value=b"MOCK_HW_123")
def test_guardar_config_segura_sincroniza_sqlite(mock_hw):
    # 1. Guardar una configuración
    test_cfg = {
        "docente_nombre": "Carlos Perez",
        "docente_cedula": "8-111-222",
        "tema": "light"
    }
    guardar_config_segura(test_cfg)
    
    # 2. Comprobar que config.enc existe
    assert os.path.exists(rdsecurity.CONFIG_FILE)
    
    # 3. Comprobar que los campos están guardados en la tabla SQLite configuracion
    import rdsql as _rdsql
    db = _rdsql.SQLDatabaseManager()
    conn = db.conectar()
    cursor = conn.cursor()
    cursor.execute("SELECT clave, valor FROM configuracion;")
    rows = dict(cursor.fetchall())
    db.cerrar()
    
    assert rows.get("docente_nombre") == "Carlos Perez"
    assert rows.get("docente_cedula") == "8-111-222"
    assert rows.get("tema") == "light"

@patch('rdsecurity._hw_fingerprint', return_value=b"MOCK_HW_123")
def test_cargar_config_segura_fallback_sqlite(mock_hw):
    # 1. Preparar datos en SQLite directamente simulando pérdida de config.enc
    test_cfg = {
        "docente_nombre": "Maria Gomez",
        "docente_cedula": "9-999-999",
        "tema": "dark"
    }
    
    # Escribir en SQLite directamente
    import rdsql as _rdsql
    db = _rdsql.SQLDatabaseManager()
    conn = db.conectar()
    cursor = conn.cursor()
    for k, v in test_cfg.items():
        cursor.execute("INSERT OR REPLACE INTO configuracion (clave, valor) VALUES (?, ?);", (k, v))
    conn.commit()
    db.cerrar()
    
    # Asegurar que config.enc NO existe
    if os.path.exists(rdsecurity.CONFIG_FILE):
        os.remove(rdsecurity.CONFIG_FILE)
        
    # 2. Cargar la configuración -> Debería disparar el fallback a SQLite
    cfg_cargada = cargar_config_segura({"tema": "light", "docente_nombre": "Default"})
    
    assert cfg_cargada.get("docente_nombre") == "Maria Gomez"
    assert cfg_cargada.get("docente_cedula") == "9-999-999"
    assert cfg_cargada.get("tema") == "dark"

@patch('rdsecurity._hw_fingerprint', return_value=b"MOCK_HW_123")
def test_obtener_datos_generales_prioriza_sqlite(mock_hw):
    # 1. Guardar datos en SQLite y cerrar para generar base cifrada
    test_cfg = {
        "docente_nombre": "Prof. Ana Lopez",
        "docente_cedula": "4-444-444",
        "ano_lectivo": "2027"
    }
    
    import rdsql as _rdsql2
    db = _rdsql2.SQLDatabaseManager()
    conn = db.conectar()
    cursor = conn.cursor()
    for k, v in test_cfg.items():
        cursor.execute("INSERT OR REPLACE INTO configuracion (clave, valor) VALUES (?, ?);", (k, v))
    conn.commit()
    db.cerrar()
    
    # 2. Instanciar DataEngine con una ruta ficticia y probar obtener_datos_generales
    engine = DataEngine("falsa_ruta.xlsx", modalidad="premedia")
    
    generales = engine.obtener_datos_generales()
    
    assert generales.get("docente_nombre") == "Prof. Ana Lopez"
    assert generales.get("docente_cedula") == "4-444-444"
    assert generales.get("ano_lectivo") == "2027"
