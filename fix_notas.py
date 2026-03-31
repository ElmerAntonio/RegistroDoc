import re

with open("src/eapp.py", "r", encoding="utf-8") as f:
    content = f.read()

# Make sure buscar_modificar populates the FIRST entry of the row, and enables visual cues
content = re.sub(
    r'        for id_est, entry in self\.entradas_notas\.items\(\):\n            entry\.delete\(0, \'end\'\)\n            if id_est in notas_existentes:\n                entry\.insert\(0, str\(notas_existentes\[id_est\]\)\)\n',
    '        for id_est, entries_list in self.entradas_notas.items():\n            for entry in entries_list: entry.delete(0, \'end\')\n            if id_est in notas_existentes:\n                entries_list[0].insert(0, str(notas_existentes[id_est]))\n',
    content
)

# En actualizar_notas(), after cleaning, also clean all entries
content = re.sub(
    r'        for entry in self\.entradas_notas\.values\(\): entry\.delete\(0, \'end\'\)',
    '        for entries_list in self.entradas_notas.values():\n            for entry in entries_list: entry.delete(0, \'end\')',
    content
)


with open("src/eapp.py", "w", encoding="utf-8") as f:
    f.write(content)
