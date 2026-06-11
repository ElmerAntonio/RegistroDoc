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

    @patch('rddata.DataEngine.obtener_estudiantes_completos')
    @patch('rddata.DataEngine.obtener_promedios_reales')
    def test_obtener_datos_reportes_trimester_filtering(self, mock_proms, mock_estudiantes):
        # Setup mock data
        mock_estudiantes.return_value = [
            {"nombre": "Estudiante 1", "cedula": "8-123-456"},
            {"nombre": "Estudiante 2", "cedula": "8-789-012"}
        ]
        
        # When called with "Trimestre 1", it should fetch that trimester's grades
        def side_effect_proms(grado, materia, periodo, wb=None):
            if periodo == "Trimestre 1":
                return {"Estudiante 1": 4.5, "Estudiante 2": 3.2}
            elif periodo == "Trimestre 2":
                return {"Estudiante 1": 3.8, "Estudiante 2": 2.8}
            return {}

        mock_proms.side_effect = side_effect_proms

        # Test Trimestre 1 with mock workbook passed to bypass path check
        mock_wb = MagicMock()
        res_t1 = self.engine.obtener_datos_reportes("8° B", "Trimestre 1", wb=mock_wb)
        self.assertEqual(res_t1["docente"][0][2], "4.50")
        self.assertEqual(res_t1["docente"][1][2], "3.20")
        self.assertEqual(res_t1["aprobados"][0][1], "APROBADO")
        self.assertEqual(res_t1["aprobados"][1][1], "APROBADO")

        # Test Trimestre 2
        res_t2 = self.engine.obtener_datos_reportes("8° B", "Trimestre 2", wb=mock_wb)
        self.assertEqual(res_t2["docente"][0][2], "3.80")
        self.assertEqual(res_t2["docente"][1][2], "2.80")
        self.assertEqual(res_t2["aprobados"][0][1], "APROBADO")
        self.assertEqual(res_t2["aprobados"][1][1], "REPROBADO")

if __name__ == '__main__':
    unittest.main()
