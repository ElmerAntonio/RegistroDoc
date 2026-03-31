import re

with open("tests/test_rdsecurity_crypto.py", "r") as f:
    content = f.read()

pattern = r"""import pytest
from unittest\.mock import patch
from src\.rdsecurity import _derivar_clave"""

new_imports = """import pytest
import os
os.environ["REGISTRODOC_MASTER_SALT"] = "TEST_SALT"

from unittest.mock import patch
from rdsecurity import _derivar_clave"""

content = re.sub(pattern, new_imports, content, flags=re.DOTALL)

with open("tests/test_rdsecurity_crypto.py", "w") as f:
    f.write(content)
