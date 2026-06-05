with open("src/rddata.py", "r", encoding="utf-8") as f:
    lines = f.readlines()

for idx, line in enumerate(lines, 1):
    if "_wb_cache" in line:
        print(f"Line {idx}: {line.strip()}")
