with open("tests/test_rdsecurity.py", "r") as f:
    content = f.read()
content = content.replace('C:\\ruta\\totalmente\\falsa\\inexistente.enc', r'C:\ruta\totalmente\falsa\inexistente.enc')
with open("tests/test_rdsecurity.py", "w") as f:
    f.write(content)
