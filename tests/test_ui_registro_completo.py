import os
import sys
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "src")))

import pytest
from unittest.mock import MagicMock


def _mock_engine():
    e = MagicMock()
    e.obtener_grados_activos.return_value = ["8° B"]
    e.obtener_materias_por_grado.return_value = ["Español"]
    e.obtener_estudiantes_completos.return_value = []
    e.obtener_fechas_asistencia.return_value = []
    e.obtener_promedios_reales.return_value = {}
    e.get_dashboard_stats.return_value = {
        "total": 0, "riesgo": 0, "honor": "-", "honor_prom": "0",
        "asistencia": "0%", "grados": ["8° B"], "habitos": {"S": 0, "R": 0, "X": 0},
    }
    return e


def _make_app(monkeypatch):
    mock_engine_class = MagicMock()
    mock_engine_class.return_value = _mock_engine()
    monkeypatch.setattr("app.DataEngine", mock_engine_class)
    monkeypatch.setattr("app.PIL_OK", False)
    monkeypatch.setattr("app.MainApplication._actualizar_horario_header", lambda self: None)
    import registrocompletoapp
    monkeypatch.setattr(registrocompletoapp.RegistroCompletoFrame, "cargar_datos_asincrono", lambda self: None)
    import app
    try:
        return app.RegistroDocApp(modalidad_inicial="premedia")
    except Exception as e:  # Tk/Tcl no operativo en el entorno headless
        pytest.skip(f"Tkinter/Tcl no está completamente operativo en este entorno: {e}")


def test_registro_completo_separa_tipos_de_nota(monkeypatch):
    """La tabla de notas debe renderizar columnas con TIPOS mezclados/desordenados
    sin lanzar excepción (protege la separación visual por tipo de nota)."""
    reg_app = _make_app(monkeypatch)
    try:
        reg_app.mostrar_registro_completo()
        frame = reg_app._frames["RegistroCompletoFrame"]

        # Aislar la parte de NOTAS (lo que se modificó): stats/hábitos como no-op
        for m in ("_renderizar_stats_notas", "_renderizar_stats_asis",
                  "_renderizar_stats_habitos", "_renderizar_habitos"):
            monkeypatch.setattr(type(frame), m, lambda self, *a, **k: None)

        data = {
            "headers_notas": [
                ("Examen 1", "Examen"),
                ("Diaria 1", "Diaria / Parcial"),
                ("Aprec 1", "Apreciación"),
                ("Diaria 2", "Diaria / Parcial"),
            ],
            "rows_notas": [
                {"id": 101, "nombre": "Alumno Uno", "notas": [4.0, 3.5, 5.0, 4.2], "promedio": 4.1},
                {"id": 102, "nombre": "Alumna Dos", "notas": [2.5, 4.8, "", 4.0], "promedio": 3.8},
            ],
            "stats_notas": {"actividades": 4, "promedio_grupal": 4.0, "tasa_aprobacion": 100.0, "max_nota": 5.0},
            "headers_asis": [], "rows_asis": [], "stats_asis": {"total_dias": 0},
            "habitos_criterios": [], "rows_habitos": [], "stats_habitos": {},
        }

        # No debe lanzar excepción al agrupar/colorear/separar por tipo
        frame._renderizar_datos(data)

        # Debe haber renderizado contenido en la tabla de notas
        assert len(frame.scroll_notas.winfo_children()) > 0
    finally:
        try:
            reg_app.destroy()
        except Exception:
            pass


def test_registro_completo_sin_notas_no_crashea(monkeypatch):
    """Sin calificaciones registradas, el render no debe fallar."""
    reg_app = _make_app(monkeypatch)
    try:
        reg_app.mostrar_registro_completo()
        frame = reg_app._frames["RegistroCompletoFrame"]
        for m in ("_renderizar_stats_notas", "_renderizar_stats_asis",
                  "_renderizar_stats_habitos", "_renderizar_habitos"):
            monkeypatch.setattr(type(frame), m, lambda self, *a, **k: None)
        data = {
            "headers_notas": [], "rows_notas": [],
            "stats_notas": {"actividades": 0, "promedio_grupal": 0.0, "tasa_aprobacion": 0.0, "max_nota": 0.0},
            "headers_asis": [], "rows_asis": [], "stats_asis": {"total_dias": 0},
            "habitos_criterios": [], "rows_habitos": [], "stats_habitos": {},
        }
        frame._renderizar_datos(data)  # no debe lanzar
    finally:
        try:
            reg_app.destroy()
        except Exception:
            pass


def test_metodos_de_ventana_no_crashean(monkeypatch):
    """Los helpers de ventana (DWM y eventos map/unmap/configure) no deben lanzar."""
    reg_app = _make_app(monkeypatch)
    try:
        class _Ev:
            widget = reg_app
            width = 1280
            height = 720

        # Desactivar animaciones DWM: seguro en Windows, no-op fuera de Windows
        reg_app._deshabilitar_animaciones_dwm()
        reg_app._on_window_map(_Ev())
        reg_app._on_window_configure(_Ev())
        reg_app._on_window_unmap(_Ev())
        assert reg_app._is_minimized is True
    finally:
        try:
            reg_app.destroy()
        except Exception:
            pass
