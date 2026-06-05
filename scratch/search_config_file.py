with open("src/sapp.py", "r", encoding="utf-8") as f:
    lines = f.readlines()

for idx, line in enumerate(lines, 1):
    if "CONFIG_FILE" in line:
        print(f"Line {idx}: {line.strip()}")
