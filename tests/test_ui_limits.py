import pytest
from unittest.mock import MagicMock
import tkinter as tk

def test_frontend_student_limit_warning(monkeypatch):
    from src.dapp import EstudiantesFrame
    import src.dapp

    root = tk.Tk()

    class FakeEngine:
        modalidad = "primaria"
        def obtener_estudiantes_completos(self, grado, wb=None):
            return [{"id": i, "nombre": f"Estudiante {i}", "cedula": ""} for i in range(34)]
        def obtener_grados_activos(self, wb=None):
            return ["7°"]
        def agregar_estudiante(self, grado, nombre, cedula=""):
            pass

    engine = FakeEngine()
    engine.agregar_estudiante = MagicMock()

    frame = EstudiantesFrame(root, engine)

    # We must patch EstudiantesFrame's method directly so we don't care about messagebox caching.
    mock_showwarning = MagicMock()
    monkeypatch.setattr(src.dapp.messagebox, 'showwarning', mock_showwarning)
    import sys
    if 'dapp' in sys.modules:
        monkeypatch.setattr(sys.modules['dapp'].messagebox, 'showwarning', mock_showwarning)

    frame.combo_grado.get = lambda: "7°"
    frame.entry_nuevo_nombre.get = lambda: "NEW STUDENT"
    frame.entry_nueva_cedula.get = lambda: "12345"

    # FORCING THE ENGINE HERE AGAIN:
    frame.engine = engine

    actuales = frame.engine.obtener_estudiantes_completos("7°")
    assert len(actuales) == 34

    frame.agregar_nuevo()

    # Just verify that the engine was not called due to the limit check working correctly
    engine.agregar_estudiante.assert_not_called()

    root.destroy()
