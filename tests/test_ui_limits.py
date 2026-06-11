import os
import sys
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "src")))

import pytest
from unittest.mock import MagicMock

def test_frontend_student_limit_warning(monkeypatch):

    import tkinter as tk
    import customtkinter as ctk
    from dapp import EstudiantesFrame
    import dapp

    try:
        root = ctk.CTk()
    except Exception as e:
        pytest.skip(f"Tkinter/Tcl no está completamente operativo en este entorno: {e}")

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

    # Patch the messagebox in the imported dapp module to prevent UI modal popups
    mock_showwarning = MagicMock()
    monkeypatch.setattr(dapp.messagebox, 'showwarning', mock_showwarning)

    frame = EstudiantesFrame(root, engine)

    frame.combo_grado.get = lambda: "7°"
    frame.entry_nuevo_nombre.get = lambda: "NEW STUDENT"
    frame.entry_nueva_cedula.get = lambda: "12345"

    actuales = frame.engine.obtener_estudiantes_completos("7°")
    assert len(actuales) == 34

    frame.agregar_nuevo()

    # Just verify that the engine was not called due to the limit check working correctly
    engine.agregar_estudiante.assert_not_called()

    root.destroy()

