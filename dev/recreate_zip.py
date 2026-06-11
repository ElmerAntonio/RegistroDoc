import os
import zipfile

def recreate_zip():
    base_dir = r"c:\RegistroDoc"
    zip_path = os.path.join(base_dir, "RegistroDoc_Limpio.zip")
    
    print(f"Creating clean ISO-standard zip archive at: {zip_path}")
    
    # Folders to recursively include in the distribution package
    include_dirs = ["src", "dependencies", "assets", "tests", "docs"]
    
    # Root level files to include
    include_files = [
        "Iniciar_RegistroDoc.bat",
        ".env.example",
        "Registro Primaria.xlsx",
        "Registro_2026.xlsx",
        "requirements.txt",
        "pytest.ini"
    ]
    
    # Temporary/User-specific files to exclude completely from package
    exclude_extensions = [".enc", ".bin", ".dat", ".zip", ".log", ".pyc"]
    
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as z:
        # 1. Add root files
        for filename in include_files:
            file_path = os.path.join(base_dir, filename)
            if os.path.exists(file_path):
                print(f"Adding root file: {filename}")
                z.write(file_path, filename)
                
        # 2. Add directories recursively
        for dirname in include_dirs:
            dir_path = os.path.join(base_dir, dirname)
            if os.path.exists(dir_path):
                print(f"Adding directory: {dirname}...")
                for root, dirs, files in os.walk(dir_path):
                    # Skip python cache directories
                    if "__pycache__" in root:
                        continue
                    for file in files:
                        # Skip compiled files or files matching excluded extensions
                        if any(file.endswith(ext) for ext in exclude_extensions):
                            continue
                        full_path = os.path.join(root, file)
                        rel_path = os.path.relpath(full_path, base_dir)
                        z.write(full_path, rel_path)
                        
    print("Clean distribution ZIP file created successfully!")

if __name__ == "__main__":
    recreate_zip()
