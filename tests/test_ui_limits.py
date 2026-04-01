import pytest
from unittest.mock import MagicMock
from dapp import EstudiantesFrame
import tkinter as tk

def test_frontend_student_limit_warning():
    # Setup mock root
    root = tk.Tk()

    # Setup mock engine
    engine = MagicMock()
    engine.modalidad = "primaria"
    # mock a full class list
    engine.obtener_estudiantes_completos.return_value = [{"id": i} for i in range(34)]

    frame = EstudiantesFrame(root, engine)

    # Mock messagebox
    from tkinter import messagebox
    messagebox.showwarning = MagicMock()

    # Simulate adding
    frame.combo_grado.set("7°")
    frame.entry_nuevo_nombre.insert(0, "NEW STUDENT")
    frame.agregar_nuevo()

    # Should show warning and NOT call engine.agregar_estudiante
    messagebox.showwarning.assert_called_once()
    assert "Límite Alcanzado" in messagebox.showwarning.call_args[0][0]
    engine.agregar_estudiante.assert_not_called()

    root.destroy()
