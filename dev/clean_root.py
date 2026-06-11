import os
import shutil

ROOT_DIR = r"c:\RegistroDoc"
DATA_DIR = os.path.join(ROOT_DIR, "data")

if os.path.exists(DATA_DIR):
    if not os.path.isdir(DATA_DIR):
        print(f"Warning: {DATA_DIR} exists but is not a directory! Removing it.")
        try:
            os.remove(DATA_DIR)
            os.makedirs(DATA_DIR)
        except Exception as e:
            print(f"Error recreating directory: {e}")
    else:
        print("data/ directory already exists.")
else:
    try:
        os.makedirs(DATA_DIR)
    except Exception as e:
        print(f"Error creating directory: {e}")

files_to_move = [
    "config.enc",
    "rd_audit.bin",
    "observaciones_plantillas.json",
    "tareas_pendientes.json"
]

for filename in files_to_move:
    src = os.path.join(ROOT_DIR, filename)
    dst = os.path.join(DATA_DIR, filename)
    if os.path.exists(src):
        try:
            if not os.path.exists(dst):
                print(f"Moving {filename} to data/")
                shutil.move(src, dst)
            else:
                print(f"File {filename} already in data/, deleting from root")
                os.remove(src)
        except Exception as e:
            print(f"Error handling {filename}: {e}")

# Move Registro_2026_bak.xlsx and backups/ to Respaldos_Locales/
respaldos_locales_dir = os.path.join(ROOT_DIR, "Respaldos_Locales")
os.makedirs(respaldos_locales_dir, exist_ok=True)

bak_file = os.path.join(ROOT_DIR, "Registro_2026_bak.xlsx")
if os.path.exists(bak_file):
    try:
        shutil.move(bak_file, os.path.join(respaldos_locales_dir, "Registro_2026_bak.xlsx"))
        print("Moved Registro_2026_bak.xlsx to Respaldos_Locales/")
    except Exception as e:
        print(f"Error moving backup file: {e}")

backups_dir = os.path.join(ROOT_DIR, "backups")
if os.path.exists(backups_dir):
    try:
        for f in os.listdir(backups_dir):
            shutil.move(os.path.join(backups_dir, f), os.path.join(respaldos_locales_dir, f))
        os.rmdir(backups_dir)
        print("Moved contents of backups/ and deleted the folder")
    except Exception as e:
        print(f"Error handling backups folder: {e}")

# Delete obsolete files and folders
obsolete_items = [
    "MagicMock",
    "plan.md",
    "Registro_2026_test.xlsx",
    "Registro_2026_test2.xlsx",
    "img"
]

for item in obsolete_items:
    path = os.path.join(ROOT_DIR, item)
    if os.path.exists(path):
        try:
            if os.path.isdir(path):
                shutil.rmtree(path)
                print(f"Deleted directory: {item}")
            else:
                os.remove(path)
                print(f"Deleted file: {item}")
        except Exception as e:
            print(f"Error deleting {item}: {e}")

# Clean docs folder of old backups
docs_dir = os.path.join(ROOT_DIR, "docs")
for f in ["Registro Primaria-back.xlsx", "Registro_2026_back.xlsx"]:
    path = os.path.join(docs_dir, f)
    if os.path.exists(path):
        try:
            os.remove(path)
            print(f"Deleted obsolete file in docs/: {f}")
        except Exception as e:
            print(f"Error deleting docs/{f}: {e}")

print("Clean root procedure completed.")
