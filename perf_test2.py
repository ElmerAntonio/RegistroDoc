import time
import os
import gc
from src.rddata import DataEngine
from src.config import BASE_DIR

ruta = os.path.join(BASE_DIR, "..", "Registro_2026.xlsx")
engine = DataEngine(ruta)

print(f"==================== RENDIMIENTO ====================")
# Test 2: Carga en memoria cache (si aplica)
start = time.time()
grados = engine.obtener_grados_activos()
end = time.time()
print(f"Obtener grados activos: {(end - start)*1000:.2f} ms")

if grados:
    # Test 3: Obtener materias por grado
    start = time.time()
    materias = engine.obtener_materias_por_grado(grados[0])
    end = time.time()
    print(f"Obtener materias para {grados[0]}: {(end - start)*1000:.2f} ms")

    # Test 4: Obtener estudiantes completos
    start = time.time()
    estudiantes = engine.obtener_estudiantes_completos(grados[0])
    end = time.time()
    print(f"Obtener estudiantes completos ({len(estudiantes)}): {(end - start)*1000:.2f} ms")

    if materias:
        start = time.time()
        promedios = engine.obtener_promedios_reales(grados[0], materias[0], "Trimestre 1")
        end = time.time()
        print(f"Obtener promedios reales ({grados[0]}, {materias[0]}): {(end - start)*1000:.2f} ms")

print(f"=====================================================")
