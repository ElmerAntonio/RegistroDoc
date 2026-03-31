import customtkinter as ctk
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

# Para probar los gráficos
class DummyEngine:
    modalidad = "premedia"
    def obtener_grados_activos(self): return ["7°", "8°"]
    def obtener_materias_por_grado(self, grado): return ["Matemáticas", "Español"]
    def obtener_estudiantes_completos(self, grado):
        return [
            {"id": "1", "nombre": "Juan", "estado": "aprobado", "promedio": 4.5},
            {"id": "2", "nombre": "Pedro", "estado": "reprobado", "promedio": 2.5},
            {"id": "3", "nombre": "Maria", "estado": "aprobado", "promedio": 3.8},
            {"id": "4", "nombre": "Ana", "estado": "reprobado", "promedio": 2.9},
        ]
    def get_dashboard_stats(self):
        return {'total': 36, 'riesgo': 8}

    # Nuevo metodo para simular la data de reportes de una materia/trimestre
    def obtener_promedios_materia(self, grado, materia, trimestre):
        return [
            {"id": "1", "nombre": "Juan", "promedio": 4.5},
            {"id": "2", "nombre": "Pedro", "promedio": 2.5},
            {"id": "3", "nombre": "Maria", "promedio": 3.8},
            {"id": "4", "nombre": "Ana", "promedio": 2.9},
            {"id": "5", "nombre": "Luis", "promedio": 3.0},
            {"id": "6", "nombre": "Sofia", "promedio": 4.0},
        ]

print("Dummy Engine created")
