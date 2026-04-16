import re

with open("tests/test_rdsecurity_crypto.py", "r") as f:
    content = f.read()

content = content.replace("from rdsecurity import _derivar_clave", "from src.rdsecurity import _derivar_clave")

with open("tests/test_rdsecurity_crypto.py", "w") as f:
    f.write(content)
