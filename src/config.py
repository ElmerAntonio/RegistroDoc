import os

# Directorio base del código fuente (src)
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# Archivo de configuración principal (perfil.json)
CONFIG_FILE = os.path.abspath(os.path.join(BASE_DIR, "..", "perfil.json"))
