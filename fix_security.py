import re

with open("src/rdsecurity.py", "r") as f:
    content = f.read()

# Already reads from .env using dotenv, but let's make sure it's strictly fail-fast with no fallbacks.
# Checking current implementation:

security_pattern = r"""# Cargar configuración segura desde .env
load_dotenv\(\)

# Clave maestra interna \(Cargada desde variable de entorno de forma segura\)
_master_salt_env = os\.environ\.get\("REGISTRODOC_MASTER_SALT"\)
if not _master_salt_env:
    raise RuntimeError\(
        "CRITICAL ERROR: REGISTRODOC_MASTER_SALT no encontrada en el entorno o archivo \.env\. "
        "El programa no puede iniciar de forma segura\."
    \)

_MASTER_SALT = _master_salt_env\.encode\("utf-8"\)"""

# It looks like the code already correctly implements the fail-fast approach for the MASTER SALT:
# `raise RuntimeError("CRITICAL ERROR...`
# However, we must ensure the `load_dotenv()` correctly resolves the `.env` file since it might be in root.
# Let's adjust load_dotenv path.

new_security = """# Cargar configuración segura desde .env
env_path = os.path.join(os.path.dirname(os.path.dirname(__file__)), '.env')
load_dotenv(dotenv_path=env_path)

# Clave maestra interna (Cargada desde variable de entorno de forma segura)
_master_salt_env = os.environ.get("REGISTRODOC_MASTER_SALT")
if not _master_salt_env:
    raise RuntimeError(
        "CRITICAL ERROR: REGISTRODOC_MASTER_SALT no encontrada en el entorno o archivo .env. "
        "El programa no puede iniciar de forma segura."
    )

_MASTER_SALT = _master_salt_env.encode("utf-8")"""

content = re.sub(security_pattern, new_security, content, flags=re.DOTALL)

with open("src/rdsecurity.py", "w") as f:
    f.write(content)
