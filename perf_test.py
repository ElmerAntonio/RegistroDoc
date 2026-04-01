import time
import os
import gc
from src.rddata import DataEngine
from src.config import BASE_DIR

# Asumiendo que Registro_2026.xlsx tiene datos suficientes
ruta = os.path.join(BASE_DIR, "..", "Registro_2026.xlsx")

if not os.path.exists(ruta):
    print("Archivo Excel no encontrado, creando uno ligero para test de rendimiento...")
    engine = DataEngine(ruta)
    engine._crear_estructura_base()

# Test 1: Instanciación e inicialización del motor
start = time.time()
engine = DataEngine(ruta)
end = time.time()
print(f"DataEngine Init: {(end - start)*1000:.2f} ms")

# Test 2: Carga en memoria cache (si aplica)
start = time.time()
grados = engine.obtener_grados_activos()
end = time.time()
print(f"Obtener grados activos: {(end - start)*1000:.2f} ms")

# Test 3: Obtener materias por grado
if grados:
    start = time.time()
    materias = engine.obtener_materias_por_grado(grados[0])
    end = time.time()
    print(f"Obtener materias para {grados[0]}: {(end - start)*1000:.2f} ms")

# Test 4: Obtener estudiantes completos
if grados:
    start = time.time()
    estudiantes = engine.obtener_estudiantes_completos(grados[0])
    end = time.time()
    print(f"Obtener estudiantes completos ({len(estudiantes)}): {(end - start)*1000:.2f} ms")

# Test 5: Obtener Resumen Dashboard Completo
start = time.time()
dashboard = engine.obtener_resumen_dashboard()
end = time.time()
print(f"Obtener Resumen Dashboard Completo: {(end - start)*1000:.2f} ms")
