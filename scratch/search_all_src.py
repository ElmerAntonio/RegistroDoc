import os

for root, dirs, files in os.walk("src"):
    for file in files:
        if file.endswith(".py"):
            path = os.path.join(root, file)
            with open(path, "r", encoding="utf-8") as f:
                lines = f.readlines()
            for idx, line in enumerate(lines, 1):
                if "CONFIG_FILE" in line and "from rdsecurity import" not in line and "CONFIG_FILE =" not in line:
                    print(f"{path} Line {idx}: {line.strip()}")
