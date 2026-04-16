import re

with open("tests/test_rdsecurity_crypto.py", "r") as f:
    content = f.read()

content = content.replace("with patch('rdsecurity.PBKDF2_ITERS', 1000):", "with patch('src.rdsecurity.PBKDF2_ITERS', 1000):")

with open("tests/test_rdsecurity_crypto.py", "w") as f:
    f.write(content)
