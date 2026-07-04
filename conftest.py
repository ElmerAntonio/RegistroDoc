"""
conftest.py — Configuración global de pytest para RegistroDoc Pro.
Resuelve el problema de que SQLDatabaseManager necesita la clave
de hardware+cédula del docente para funcionar. En tests se usa
una BD SQLite en memoria completamente limpia.
"""
import os
import sys
import sqlite3
import pytest

# Asegurar modo DEV para que la validación de proceso pase
os.environ["REGISTRODOC_DEV_MODE"] = "1"

# Añadir src al path de Python
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "src"))


def _build_inmemory_manager():
    """Crea una instancia de SQLDatabaseManager que opera en memoria."""
    from rdsql import SQLDatabaseManager

    class _InMemoryDBManager(SQLDatabaseManager):
        """Subclase que reemplaza conectar() por una BD in-memory."""

        def __init__(self):
            # Inicializar campos mínimos sin llamar al __init__ real
            # (que necesita hardware fingerprint y clave del docente)
            import threading
            self.db_lock = threading.RLock()
            self._save_in_progress = False
            self._save_lock = threading.Lock()
            self.conn = None
            self.key = b"\x00" * 32
            self.raw_hw_key = b"\x00" * 32
            self.claves_candidatas = [self.key]
            self.user_dir = os.path.join(os.path.expanduser("~"), "AppData", "Local", "RegistroDoc")
            self.db_enc_path = ":memory:"
            self.db_temp_path = ":memory:"

        def conectar(self):
            if self.conn:
                return self.conn
            self.conn = sqlite3.connect(":memory:", check_same_thread=False)
            self.conn.execute("PRAGMA foreign_keys = ON;")
            self.inicializar_tablas(self.conn)
            return self.conn

        def encriptar_campo(self, texto: str) -> str:
            return texto or ""

        def desencriptar_campo(self, valor: str) -> str:
            return valor or ""

        def guardar_cifrado(self):
            pass  # No-op en tests

        def cerrar(self):
            if self.conn:
                self.conn.close()

    return _InMemoryDBManager


@pytest.fixture(autouse=True)
def patch_sql_manager(monkeypatch, request):
    """Parchea SQLDatabaseManager con BD en memoria compartida para todos los tests.
    Todas las instancias creadas dentro del mismo test comparten la misma BD en memoria.
    """
    if "test_configuracion" in str(request.node.fspath):
        yield
        return

    import rdsql
    import sqlite3

    original_cls = rdsql.SQLDatabaseManager

    # Estado compartido para el test: una sola conexión en memoria y flag de init
    _state = {"conn": None, "initialized": False}

    InMemory = _build_inmemory_manager()

    class _SharedInMemoryDBManager(InMemory):
        """Todos los tests comparten la misma BD en memoria (singleton de conexión)."""

        def __init__(self):
            super().__init__()

        def conectar(self):
            if _state["conn"] is None:
                conn = sqlite3.connect(":memory:", check_same_thread=False)
                conn.execute("PRAGMA foreign_keys = ON;")
                self.inicializar_tablas(conn)
                _state["conn"] = conn
                _state["initialized"] = True
            self.conn = _state["conn"]
            return _state["conn"]

        def cerrar(self):
            pass  # No cerrar la conexión compartida hasta el teardown del fixture

    monkeypatch.setattr(rdsql, "SQLDatabaseManager", _SharedInMemoryDBManager)

    yield

    monkeypatch.setattr(rdsql, "SQLDatabaseManager", original_cls)
    if _state["conn"] is not None:
        try:
            _state["conn"].close()
        except Exception:
            pass
        _state["conn"] = None
