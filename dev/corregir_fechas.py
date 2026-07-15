"""
RegistroDoc Pro — Utilidad para corregir fechas de asistencia en el segundo trimestre.
Corrige el formato de Día-Mes a Mes-Día (MM-DD) para las fechas con ambigüedad.
"""

import os
import sys

# Habilitar el modo de desarrollo para omitir restricciones del proceso ejecutor
os.environ["REGISTRODOC_DEV_MODE"] = "1"

# Agregar la ruta del código fuente al path de Python
sys.path.append(os.path.abspath("src"))

from rdsql import SQLDatabaseManager

def main():
    print("=== Corrector de Fechas de Asistencia (Trimestre 2) ===")
    
    try:
        db_mgr = SQLDatabaseManager()
    except Exception as e:
        print(f"Error al inicializar el Database Manager: {e}")
        return

    if not os.path.exists(db_mgr.db_enc_path):
        print("[-] Error: No se encontró el archivo de base de datos cifrado.")
        return

    # Conectar y descifrar la base de datos
    try:
        conn = db_mgr.conectar()
    except Exception as e:
        print(f"[-] Error al conectar con la base de datos: {e}")
        return

    # Mapeo de fechas de formato erróneo (Día-Mes) a correcto (Mes-Día)
    mapping = {
        '29-06': '06-29', # 29 de junio
        '30-06': '06-30', # 30 de junio
        '06-07': '07-06', # 06 de julio
        '08-06': '06-08', # 08 de junio
        '08-07': '07-08', # 08 de julio
        '09-06': '06-09', # 09 de junio
        '09-07': '07-09', # 09 de julio
        '10-06': '06-10', # 10 de junio
        '10-07': '07-10', # 10 de julio
    }

    try:
        cursor = conn.cursor()
        
        # Consultar antes de actualizar
        cursor.execute("SELECT DISTINCT fecha FROM asistencia WHERE trimestre = 2 ORDER BY fecha;")
        antes = [r[0] for r in cursor.fetchall()]
        print(f"[*] Fechas antes de corregir: {antes}")
        
        # Realizar actualizaciones
        total_updated = 0
        for old_date, new_date in mapping.items():
            cursor.execute(
                "UPDATE asistencia SET fecha = ? WHERE fecha = ? AND trimestre = 2;",
                (new_date, old_date)
            )
            count = cursor.rowcount
            if count > 0:
                print(f"[+] Corregida fecha '{old_date}' -> '{new_date}' ({count} registros actualizados)")
                total_updated += count
                
        # Confirmar los cambios en SQLite
        conn.commit()
        
        # Forzar el guardado y cifrado de vuelta al disco
        print("[*] Guardando y cifrando base de datos corregida en disco...")
        db_mgr.guardar_cifrado()
        
        # Consultar después de actualizar para verificar
        cursor.execute("SELECT DISTINCT fecha FROM asistencia WHERE trimestre = 2 ORDER BY fecha;")
        despues = [r[0] for r in cursor.fetchall()]
        print(f"[*] Fechas después de corregir: {despues}")
        
        print(f"\n[+] ¡Éxito! Se actualizaron un total de {total_updated} registros de asistencia.")
        
    except Exception as e:
        print(f"[-] Ocurrió un error al actualizar las fechas: {e}")
        try:
            conn.rollback()
        except:
            pass
    finally:
        db_mgr.cerrar()

if __name__ == "__main__":
    main()
