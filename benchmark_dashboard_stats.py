import time
import os
from src.rddata import DataEngine

def benchmark():
    # Use Registro_2026.xlsx as it's the larger file
    ruta_excel = "Registro_2026.xlsx"
    if not os.path.exists(ruta_excel):
        print(f"Error: {ruta_excel} not found.")
        return

    engine = DataEngine(ruta_excel, modalidad="premedia")

    # Warm up
    engine.get_dashboard_stats()

    start_time = time.time()
    iterations = 5
    for _ in range(iterations):
        stats = engine.get_dashboard_stats()
    end_time = time.time()

    avg_time = (end_time - start_time) / iterations
    print(f"Average time for get_dashboard_stats: {avg_time:.4f} seconds")
    print(f"Stats: {stats}")

if __name__ == "__main__":
    benchmark()
