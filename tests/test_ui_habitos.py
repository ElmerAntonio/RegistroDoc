import os
import sys
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "src")))

import pytest
from unittest.mock import MagicMock

def test_ui_habitos_initialization_and_update(monkeypatch):
    import tkinter as tk
    import customtkinter as ctk
    from habapp import HabitosFrame

    try:
        root = ctk.CTk()
    except Exception as e:
        pytest.skip(f"Tkinter/Tcl no está completamente operativo en este entorno: {e}")

    class FakeCursor:
        def execute(self, *args, **kwargs):
            return self
        def fetchall(self):
            return []
        def fetchone(self):
            return None

    class FakeConn:
        def cursor(self):
            return FakeCursor()
        def commit(self):
            pass

    class FakeEngine:
        modalidad = "premedia"
        db_conn = FakeConn()
        db_manager = MagicMock()
        
        def obtener_grados_activos(self, wb=None):
            return ["7° A"]
            
        def obtener_estudiantes_completos(self, grado, wb=None):
            return [{"id": 102, "nombre": "RAMOS, Elmer", "cedula": "", "sexo": "M"}]

    engine = FakeEngine()
    frame = HabitosFrame(root, engine)

    # Trigger view update
    frame.actualizar_vista()

    # Assertions
    assert frame.combo_grado.get() == "7° A"
    assert frame.combo_trimestre.get() in ["Trimestre 1", "Trimestre 2", "Trimestre 3"]
    assert frame.combo_frecuencia.get() == "Trimestral"
    assert len(frame.estudiantes) == 1
    assert frame.estudiantes[0]["nombre"] == "RAMOS, Elmer"

    root.destroy()
