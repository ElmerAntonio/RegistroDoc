import os
import shutil
import glob

def main():
    workspace_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    package_dir = os.path.join(workspace_dir, "package_data")
    
    # 1. Crear directorios destino
    pkg_data_dir = os.path.join(package_dir, "data")
    pkg_appdata_data = os.path.join(package_dir, "appdata", "data")
    pkg_appdata_temp = os.path.join(package_dir, "appdata", "temp")
    
    for d in [pkg_data_dir, pkg_appdata_data, pkg_appdata_temp]:
        os.makedirs(d, exist_ok=True)
        
    print(f"Estructura de package_data creada en: {package_dir}")
    
    # 2. Copiar archivos de configuración de la raíz (data/)
    src_data_dir = os.path.join(workspace_dir, "data")
    if os.path.exists(src_data_dir):
        for item in os.listdir(src_data_dir):
            src_file = os.path.join(src_data_dir, item)
            if os.path.isfile(src_file) and not item.endswith(".db.enc"):
                shutil.copy2(src_file, os.path.join(pkg_data_dir, item))
                print(f"Copiado archivo de configuración: {item}")
    else:
        print("ADVERTENCIA: Directorio data/ del workspace no encontrado.")

    # 3. Copiar base de datos activa desde LOCALAPPDATA
    local_appdata = os.environ.get("LOCALAPPDATA", os.path.expanduser("~\\AppData\\Local"))
    src_appdata_db = os.path.join(local_appdata, "RegistroDoc", "data", "registro.db.enc")
    
    if os.path.exists(src_appdata_db):
        shutil.copy2(src_appdata_db, os.path.join(pkg_appdata_data, "registro.db.enc"))
        print(f"Copiada base de datos activa desde: {src_appdata_db}")
    else:
        print("ADVERTENCIA: Base de datos activa no encontrada en LOCALAPPDATA/RegistroDoc/data/")
        
    # 4. Copiar planillas Excel de trabajo activas desde LOCALAPPDATA temp/
    src_appdata_temp = os.path.join(local_appdata, "RegistroDoc", "temp")
    if os.path.exists(src_appdata_temp):
        excel_files = glob.glob(os.path.join(src_appdata_temp, "*.xlsx"))
        for xls in excel_files:
            # Ignorar backups y archivos temporales/temporales de Excel
            base_name = os.path.basename(xls)
            if base_name in ("Registro_Premedia.xlsx", "Registro_Primaria.xlsx"):
                shutil.copy2(xls, os.path.join(pkg_appdata_temp, base_name))
                print(f"Copiada planilla Excel activa: {base_name}")
    else:
        print("ADVERTENCIA: Directorio temporal no encontrado en LOCALAPPDATA/RegistroDoc/temp/")

    print("Preparación de datos completada con éxito.")

if __name__ == "__main__":
    main()
