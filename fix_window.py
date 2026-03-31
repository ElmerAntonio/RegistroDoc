import re

with open("src/app.py", "r", encoding="utf-8") as f:
    content = f.read()

# Eliminar el intento de iniciar maximizado que causa el "espectáculo"
content = re.sub(
    r'        # Intentar iniciar maximizado en Windows\n        try:\n            self.state\("zoomed"\)\n        except Exception:\n            pass\n',
    '',
    content
)

with open("src/app.py", "w", encoding="utf-8") as f:
    f.write(content)
