import unittest
from unittest.mock import patch, MagicMock
import sys
import os

# Add src to path
src_path = os.path.abspath(os.path.join(os.path.dirname(__file__), '../src'))
if src_path not in sys.path:
    sys.path.insert(0, src_path)

from rddata import DataEngine

class MockCursor:
    def __init__(self):
        self.query = ""
        self.params = ()

    def execute(self, sql, params=()):
        self.query = sql
        self.params = params

    def fetchone(self):
        if "FROM grados" in self.query:
            return (1,)
        return None

    def fetchall(self):
        if "FROM estudiantes" in self.query:
            if "cedula" in self.query:
                return [
                    (10, "Estudiante 1", "8-123-456"),
                    (11, "Estudiante 2", "8-789-012")
                ]
            else:
                return [
                    (10, "Estudiante 1"),
                    (11, "Estudiante 2")
                ]
        elif "FROM materias" in self.query:
            return [
                (101, "Materia 1")
            ]
        elif "FROM notas" in self.query:
            est_id = self.params[0]
            trim = self.params[2]
            print(f"DEBUG FROM notas: est_id={est_id}, trim={trim}")
            if est_id == 10:
                return [("Diaria / Parcial", 4.5)]
            elif est_id == 11:
                if trim == 2:
                    return [("Diaria / Parcial", 2.8)]
                else:
                    return [("Diaria / Parcial", 3.2)]
        return []

class TestReportesTrimester(unittest.TestCase):

    def setUp(self):
        self.engine = DataEngine("dummy.xlsx")

    def test_obtener_datos_reportes_trimester_filtering(self):
        mock_conn = MagicMock()
        mock_cursor = MockCursor()
        mock_conn.cursor.return_value = mock_cursor

        self.engine.db_manager = MagicMock()
        self.engine.db_manager.desencriptar_campo.side_effect = lambda x: x
        self.engine.db_conn = mock_conn

        res_t1 = self.engine.obtener_datos_reportes("8° B", "Trimestre 1")
        print("DEBUG res_t1:", res_t1)
        self.assertEqual(len(res_t1["docente"]), 2)
        self.assertEqual(res_t1["docente"][0][2], "4.50")

    def test_obtener_datos_reportes_reprobado(self):
        mock_conn = MagicMock()
        mock_cursor = MockCursor()
        mock_conn.cursor.return_value = mock_cursor

        self.engine.db_manager = MagicMock()
        self.engine.db_manager.desencriptar_campo.side_effect = lambda x: x
        self.engine.db_conn = mock_conn

        res = self.engine.obtener_datos_reportes("8° B", "Trimestre 2")
        self.assertEqual(len(res["aprobados"]), 2)
        self.assertEqual(res["aprobados"][1][1], "REPROBADO")

if __name__ == '__main__':
    unittest.main()
