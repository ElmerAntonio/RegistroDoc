import re

with open("tests/test_rdprint.py", "r") as f:
    content = f.read()

pattern = r"""# Mock cryptography for rdsecurity if needed during test imports
sys\.modules\['cryptography'\] = MagicMock\(\)
sys\.modules\['cryptography\.hazmat'\] = MagicMock\(\)
sys\.modules\['cryptography\.hazmat\.primitives'\] = MagicMock\(\)"""

content = re.sub(pattern, "", content, flags=re.DOTALL)

with open("tests/test_rdprint.py", "w") as f:
    f.write(content)
