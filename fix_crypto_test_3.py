import re

with open("tests/test_rdsecurity_crypto.py", "r") as f:
    content = f.read()

pattern = r"""        with patch\('rdsecurity\.PBKDF2_ITERS', 1000\):
            derived_key = _derivar_clave\(password, salt\)
            # Update with the actual key derived for the test conditions.
            # The previous expected key might have used different conditions \(SHA1 vs SHA256\)
            assert derived_key\.hex\(\) == '2722f1cbc9f5e61fbdd5eb1704b907e302a724f2d17043a566d1f8595906007f'"""

new_test = """    with patch('rdsecurity.PBKDF2_ITERS', 1000):
        derived_key = _derivar_clave(password, salt)
        assert derived_key.hex() == expected_key_hex"""

content = re.sub(pattern, new_test, content, flags=re.DOTALL)

with open("tests/test_rdsecurity_crypto.py", "w") as f:
    f.write(content)
