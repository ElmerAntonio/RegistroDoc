import sys
import os
sys.path.insert(0, os.path.abspath("src"))
from rddata import DataEngine

engine = DataEngine("Registro_2026.xlsx")
cuadro = engine.obtener_cuadro_honor_general()
print("Cuadro de honor length:", len(cuadro))
for est in cuadro[:10]:
    print(f"- {est['nombre']} ({est['grado']}): {est['promedio']:.2f}")
