import os
import sys
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "src")))

import pytest
from unittest.mock import MagicMock

def test_registrodoc_app_lifecycle(monkeypatch):

    import tkinter as tk
    import customtkinter as ctk

    # Mock DataEngine to isolate the UI tests from files
    mock_engine_class = MagicMock()
    mock_engine = MagicMock()
    mock_engine_class.return_value = mock_engine
    
    # Set up dummy methods on mock_engine
    mock_engine.obtener_grados_activos.return_value = ["8° B"]
    mock_engine.obtener_materias_por_grado.return_value = ["Español"]
    mock_engine.obtener_estudiantes_completos.return_value = []
    mock_engine.obtener_fechas_asistencia.return_value = []
    mock_engine.obtener_promedios_reales.return_value = {"Estudiante 1": 4.5}
    mock_engine.get_dashboard_stats.return_value = {
        "total": 92,
        "riesgo": 7,
        "honor": "SANTOS FIDEL",
        "honor_prom": "4.9",
        "asistencia": "98%",
        "grados": ["8° B"],
        "habitos": {"S": 10, "R": 2, "X": 1}
    }
    
    # Patch DataEngine in app
    monkeypatch.setattr("app.DataEngine", mock_engine_class)
    
    # Patch image loading in MainApplication sidebar to avoid PIL issues if any
    monkeypatch.setattr("app.PIL_OK", False)
    
    # Disable the background schedule updates to keep tests fast and synchronous
    monkeypatch.setattr("app.MainApplication._actualizar_horario_header", lambda self: None)
    
    # Disable background data loading in RegistroCompletoFrame to avoid race conditions in tests
    import registrocompletoapp
    monkeypatch.setattr(registrocompletoapp.RegistroCompletoFrame, "cargar_datos_asincrono", lambda self: None)
    
    import app
    
    # Instantiate the app
    try:
        reg_app = app.RegistroDocApp(modalidad_inicial="premedia")
    except Exception as e:
        pytest.skip(f"Tkinter/Tcl no está completamente operativo en este entorno: {e}")
    
    try:
        # Verify app properties
        assert reg_app._destroyed is False
        assert reg_app.title() == "RegistroDoc Pro v3.0 — MEDUCA Panamá"
        
        # Test navigation functions
        reg_app.mostrar_dashboard()
        assert reg_app.main_app.menu_activo == "Inicio"
        
        reg_app.mostrar_estudiantes()
        assert reg_app.main_app.menu_activo == "Estudiantes"
        
        reg_app.mostrar_notas()
        assert reg_app.main_app.menu_activo == "Notas y Asistencia"
        
        reg_app.mostrar_asistencia()
        assert reg_app.main_app.menu_activo == "Notas y Asistencia"
        
        reg_app.mostrar_registro_completo()
        assert reg_app.main_app.menu_activo == "Registro Completo"
        
        reg_app.mostrar_observaciones()
        assert reg_app.main_app.menu_activo == "Observaciones"
        
        reg_app.mostrar_habitos()
        assert reg_app.main_app.menu_activo == "Hábitos"
        
        reg_app.mostrar_reportes()
        assert reg_app.main_app.menu_activo == "Reportes y Gráficos"
        
        reg_app.mostrar_impresion()
        assert reg_app.main_app.menu_activo == "Impresión"
        
        reg_app.mostrar_configuracion()
        assert reg_app.main_app.menu_activo == "Configuración"
        
        reg_app.mostrar_ayuda()
        assert reg_app.main_app.menu_activo == "Ayuda y Guía"
        
        reg_app.mostrar_tareas()
        assert reg_app.main_app.menu_activo == "Tareas"
        
        reg_app.mostrar_reuniones()
        assert reg_app.main_app.menu_activo == "Reuniones"
        
        # Test closing
        reg_app.on_closing()
        assert reg_app._destroyed is True
        
    finally:
        # Ensure cleanup in case of failures
        try:
            reg_app.destroy()
        except Exception:
            pass
