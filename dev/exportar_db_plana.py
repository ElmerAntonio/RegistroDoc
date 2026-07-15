"""
RegistroDoc Pro — Utilidad para exportar base de datos descifrada a un archivo SQLite plano.
Esta utilidad crea una copia local 'registro_plano.db' con todos los campos desencriptados
(incluyendo nombres de estudiantes, cédulas y comentarios) para su inspección y depuración.

Uso:
  python dev/exportar_db_plana.py
"""

import os
import sys
import shutil
import sqlite3

# Habilitar el modo de desarrollo para omitir restricciones del proceso ejecutor
os.environ["REGISTRODOC_DEV_MODE"] = "1"

# Agregar la ruta del código fuente al path de Python
sys.path.append(os.path.abspath("src"))

from rdsql import SQLDatabaseManager

def main():
    print("=== Exportador de Base de Datos RegistroDoc (Plana y Descifrada) ===")
    
    # 1. Instanciar el administrador de base de datos
    try:
        db_mgr = SQLDatabaseManager()
    except Exception as e:
        print(f"Error al inicializar el Database Manager: {e}")
        return

    print(f"[*] Base cifrada origen: {db_mgr.db_enc_path}")
    if not os.path.exists(db_mgr.db_enc_path):
        print("[-] Error: No se encontró el archivo cifrado. ¿Ya has iniciado la aplicación al menos una vez?")
        return

    # 2. Conectar (esto descifra la BD cifrada en el archivo temporal local de SQLite)
    try:
        print("[*] Descifrando base de datos en repositorio temporal...")
        conn = db_mgr.conectar()
    except Exception as e:
        print(f"[-] Error al descifrar la base de datos: {e}")
        return

    target_db = "registro_plano.db"
    
    try:
        # 3. Copiar la base de datos temporal descifrada a nuestro archivo destino plano
        print(f"[*] Duplicando base temporal a '{target_db}'...")
        if os.path.exists(target_db):
            os.remove(target_db)
        shutil.copy2(db_mgr.db_temp_path, target_db)
        
        # Cerrar conexión temporal del manager
        db_mgr.cerrar()
        
        # 4. Abrir la nueva base de datos plana para desencriptar columnas internas
        print("[*] Desencriptando campos internos a nivel de columna (nombres, cédulas, observaciones)...")
        plano_conn = sqlite3.connect(target_db)
        cursor = plano_conn.cursor()
        
        # A. Desencriptar estudiantes
        cursor.execute("SELECT id, nombre, cedula, sexo, nombre_acudiente FROM estudiantes;")
        estudiantes = cursor.fetchall()
        for est_id, nombre_enc, cedula_enc, sexo_enc, acudiente_enc in estudiantes:
            nombre_dec = db_mgr.desencriptar_campo(nombre_enc)
            cedula_dec = db_mgr.desencriptar_campo(cedula_enc) if cedula_enc else ""
            sexo_dec = db_mgr.desencriptar_campo(sexo_enc) if sexo_enc else ""
            acudiente_dec = db_mgr.desencriptar_campo(acudiente_enc) if acudiente_enc else ""
            
            cursor.execute(
                "UPDATE estudiantes SET nombre = ?, cedula = ?, sexo = ?, nombre_acudiente = ? WHERE id = ?;",
                (nombre_dec, cedula_dec, sexo_dec, acudiente_dec, est_id)
            )
            
        # B. Desencriptar observaciones
        try:
            cursor.execute("SELECT id, comentario FROM observaciones;")
            obs = cursor.fetchall()
            for obs_id, comentario_enc in obs:
                comentario_dec = db_mgr.desencriptar_campo(comentario_enc)
                cursor.execute(
                    "UPDATE observaciones SET comentario = ? WHERE id = ?;",
                    (comentario_dec, obs_id)
                )
        except sqlite3.OperationalError:
            # Si no existe la tabla observaciones, omitir
            pass

        plano_conn.commit()
        plano_conn.close()
        
        print("\n[+] ¡Éxito!")
        print(f"[+] Se ha creado el archivo: {os.path.abspath(target_db)}")
        print("[+] Puedes abrir este archivo usando cualquier visor de SQLite (ej. 'DB Browser for SQLite')")
        print("[+] Advertencia: Mantén este archivo plano protegido, ya que contiene datos personales en texto claro.")
        
    except Exception as e:
        print(f"[-] Ocurrió un error durante el procesamiento: {e}")
        if os.path.exists(target_db):
            try:
                os.remove(target_db)
            except:
                pass

if __name__ == "__main__":
    main()
