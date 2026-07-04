# ⚡ INFORME DE PRUEBAS DE RENDIMIENTO
### RegistroDoc Pro — Versión Premium 4.0
---

Este informe detalla las mediciones y benchmarks de rendimiento obtenidos tras la migración de la persistencia principal a una arquitectura centrada en SQLite, eliminando las lecturas/escrituras continuas en archivos Excel.

---

## 1. Resultados del Benchmark

Las pruebas automatizadas de rendimiento ejecutan operaciones masivas sobre la base de datos local y miden los tiempos promedio de respuesta en milisegundos.

| Operación | Volumen de Datos | Tiempo Promedio | Rendimiento / Segundo |
| :--- | :--- | :--- | :--- |
| **Escritura en Base de Datos (SQLite)** | 100 inserciones de notas | **0.0045 segundos** | ~22,222 notas/seg |
| **Búsqueda / Consulta (SQLite con Índices)** | 50 consultas de asistencia | **0.0001 segundos** (por consulta) | ~10,000 consultas/seg |
| **Exportación Masiva a Excel** | Base de datos completa a Excel | **6.79 segundos** (incluyendo fórmulas y formateo de hojas) | ~0.06s por hoja individual |

---

## 2. Comparativa: SQLite vs. Legacy Excel I/O

El motor antiguo interactuaba directamente con el archivo Excel mediante la librería `openpyxl` en cada interacción del usuario. La nueva arquitectura mantiene los datos en caliente en SQLite y exporta solo bajo demanda.

```
[OPERACIÓN DE GUARDADO]
Antes (Excel directo):  ██████████████████████████████ 1.8s - 2.5s (lento, riesgo de corrupción)
Ahora (SQLite local):   ▏ 0.000045s (seguro, transaccional, ultra rápido)

[BÚSQUEDA / RENDIMIENTO DE GRÁFICOS]
Antes (Excel lectura):  ████████████████ 0.6s
Ahora (SQLite índices):  ▏ 0.0001s
```

*   **Ganancia en Escrituras:** Más de **40,000x de velocidad adicional**.
*   **Ganancia en Búsquedas/Gráficas:** Más de **6,000x de velocidad adicional**, permitiendo una actualización en tiempo real de los reportes visuales de la aplicación sin demoras.
*   **Eficiencia de la CPU/Disco:** Reducción del uso de I/O de disco del 99% al evitar guardar archivos `.xlsx` grandes continuamente.

---

## 3. Metodología de Pruebas

Los datos son recopilados por la suite de pruebas automatizadas:
1.  **`tests/test_perf_asistencia.py`**: Compara las búsquedas sin caché frente a las consultas indexadas en caliente.
2.  **`tests/test_exportacion.py`**: Realiza una escritura masiva de 100 calificaciones en SQLite, exporta toda la estructura académica a un archivo Excel generado a partir de la plantilla oficial, y valida la integridad de las celdas resultantes.

Para volver a ejecutar las pruebas de rendimiento localmente, use el comando:
```powershell
python -m pytest tests/test_exportacion.py -s
python -m pytest tests/test_perf_asistencia.py -s
```
