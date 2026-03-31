import re

with open("tests/test_rdsecurity.py", "r") as f:
    content = f.read()

# Make sure we set the environment variable BEFORE importing rdsecurity
pattern = r"""import rdsecurity
from rdsecurity import \(
    cifrar, descifrar, guardar_cifrado, cargar_cifrado,
    _hw_fingerprint, verificar_licencia, validar_nota_meduca
\)
from unittest\.mock import patch

# Master salt dummy setup para pruebas locales
os\.environ\["REGISTRODOC_MASTER_SALT"\] = "TEST_SALT\""""

new_imports = """# Master salt dummy setup para pruebas locales DEBE IR ANTES DEL IMPORT
os.environ["REGISTRODOC_MASTER_SALT"] = "TEST_SALT"

import rdsecurity
from rdsecurity import (
    cifrar, descifrar, guardar_cifrado, cargar_cifrado,
    _hw_fingerprint, verificar_licencia, validar_nota_meduca
)
from unittest.mock import patch"""

content = re.sub(pattern, new_imports, content, flags=re.DOTALL)

with open("tests/test_rdsecurity.py", "w") as f:
    f.write(content)
