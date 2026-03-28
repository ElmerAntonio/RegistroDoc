from src.rddata import DataEngine
import os

# Prueba con el de Premedia
archivo_premedia = "Registro_2026.xlsx"
if os.path.exists(archivo_premedia):
    engine = DataEngine(archivo_premedia, "premedia")
    print("Limpiando Premedia...")
    engine.reiniciar_libreta()
    print("✅ Premedia Limpio.")

# Prueba con el de Primaria
archivo_primaria = "Registro_Primaria.xlsx"
if os.path.exists(archivo_primaria):
    engine = DataEngine(archivo_primaria, "primaria")
    print("Limpiando Primaria...")
    engine.reiniciar_libreta()
    print("✅ Primaria Limpio.")