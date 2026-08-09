import os
import sys
import sqlite3

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "src")))

from rdsql import SQLDatabaseManager


def test_sqlite_integro_acepta_base_valida(tmp_path):
    """Una BD SQLite válida debe pasar quick_check → True."""
    ruta = str(tmp_path / "valida.db")
    c = sqlite3.connect(ruta)
    c.execute("CREATE TABLE estudiantes (id INTEGER PRIMARY KEY, nombre TEXT);")
    c.execute("INSERT INTO estudiantes (nombre) VALUES ('Alumno');")
    c.commit()
    c.close()
    # _sqlite_integro no usa estado de instancia → se puede llamar con self=None
    assert SQLDatabaseManager._sqlite_integro(None, ruta) is True


def test_sqlite_integro_rechaza_archivo_corrupto(tmp_path):
    """Un archivo que no es una BD SQLite válida debe fallar → False
    (esto dispara la restauración automática desde backup)."""
    ruta = str(tmp_path / "corrupta.db")
    with open(ruta, "wb") as f:
        f.write(b"esto no es una base de datos sqlite valida " * 100)
    assert SQLDatabaseManager._sqlite_integro(None, ruta) is False


def test_sqlite_integro_rechaza_inexistente(tmp_path):
    """Un archivo inexistente debe fallar de forma segura → False."""
    ruta = str(tmp_path / "no_existe.db")
    assert SQLDatabaseManager._sqlite_integro(None, ruta) is False


def test_sqlite_integro_rechaza_encabezado_falso(tmp_path):
    """Archivo con encabezado SQLite falso pero contenido inválido → False."""
    ruta = str(tmp_path / "fake_header.db")
    with open(ruta, "wb") as f:
        f.write(b"SQLite format 3\x00")          # encabezado real
        f.write(os.urandom(4096))                 # páginas basura
    assert SQLDatabaseManager._sqlite_integro(None, ruta) is False
