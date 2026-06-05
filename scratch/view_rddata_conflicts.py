with open("src/rddata.py", "r", encoding="utf-8") as f:
    lines = f.readlines()

in_conflict = False
conflict_lines = []
for idx, line in enumerate(lines, 1):
    if line.startswith("<<<<<<<"):
        in_conflict = True
        conflict_lines.append(f"--- Conflict Start at Line {idx} ---")
    if in_conflict:
        conflict_lines.append(f"{idx}: {line.strip()}")
    if line.startswith(">>>>>>>"):
        in_conflict = False
        conflict_lines.append("--- Conflict End ---\n")

print("\n".join(conflict_lines[:150]))
if len(conflict_lines) > 150:
    print(f"... and {len(conflict_lines) - 150} more lines of conflict.")
