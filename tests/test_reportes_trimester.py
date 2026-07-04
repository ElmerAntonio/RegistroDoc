import unittest
from unittest.mock import patch, MagicMock
import sys
import os

# Add src to path
src_path = os.path.abspath(os.path.join(os.path.dirname(__file__), '../src'))
if src_path not in sys.path:
    sys.path.insert(0, src_path)

from rddata import DataEngine

class TestReportesTrimester(unittest.TestCase):

    def setUp(self):
        self.engine = DataEngine("dummy.xlsx")

    def test_obtener_datos_reportes_trimester_filtering(self):
        """
        Verifica que obtener_datos_reportes filtra correctamente por trimestre
        usando consultas SQL. El método ya no acepta wb — es SQL puro.
        """
        # Mock del cursor SQL para simular la BD
        mock_cursor = MagicMock()
        mock_conn = MagicMock()
        mock_conn.cursor.return_value = mock_cursor

        # Grado encontrado
        mock_cursor.fetchone.side_effect = [
            (1,),           # SELECT id FROM grados → grado_id=1
            (4.5,),         # AVG nota Estudiante 1 - Trimestre 1
            (3.2,),         # AVG nota Estudiante 2 - Trimestre 1
        ]
        # Estudiantes del grado
        mock_cursor.fetchall.return_value = [
            (10, "Estudiante 1", "8-123-456"),
            (11, "Estudiante 2", "8-789-012"),
        ]

        # db_manager.descifrar_campo devuelve el valor sin cambios (no hay cifrado en test)
        self.engine.db_manager = MagicMock()
        self.engine.db_manager.descifrar_campo.side_effect = lambda x: x

        self.engine.db_conn = mock_conn

        res_t1 = self.engine.obtener_datos_reportes("8° B", "Trimestre 1")
        self.assertEqual(len(res_t1["docente"]), 2)
        self.assertEqual(res_t1["docente"][0][2], "4.50")
        self.assertEqual(res_t1["docente"][1][2], "3.20")
        self.assertEqual(res_t1["aprobados"][0][1], "APROBADO")
        self.assertEqual(res_t1["aprobados"][1][1], "APROBADO")

    def test_obtener_datos_reportes_reprobado(self):
        """Verifica que un promedio menor a 3.0 se marca como REPROBADO."""
        mock_cursor = MagicMock()
        mock_conn = MagicMock()
        mock_conn.cursor.return_value = mock_cursor

        mock_cursor.fetchone.side_effect = [
            (1,),    # grado_id
            (2.8,),  # promedio Estudiante 2 - reprobado
        ]
        mock_cursor.fetchall.return_value = [
            (11, "Estudiante 2", "8-789-012"),
        ]

        self.engine.db_manager = MagicMock()
        self.engine.db_manager.descifrar_campo.side_effect = lambda x: x
        self.engine.db_conn = mock_conn

        res = self.engine.obtener_datos_reportes("8° B", "Trimestre 2")
        self.assertEqual(len(res["aprobados"]), 1)
        self.assertEqual(res["aprobados"][0][1], "REPROBADO")
        self.assertEqual(res["docente"][0][2], "2.80")

if __name__ == '__main__':
    unittest.main()
