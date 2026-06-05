import re

with open("src/rddata.py", "r", encoding="utf-8") as f:
    content = f.read()

# 1. Add self._last_mtime to __init__
init_target = 'self._wb_cache = None\n        self._cargar_en_memoria()'
init_replacement = 'self._wb_cache = None\n        self._last_mtime = 0.0\n        self._cargar_en_memoria()'
content = content.replace(init_target, init_replacement)

# 2. Update _cargar_en_memoria and add _verificar_y_recargar_cache, _save_wb
cargar_target = """    def _cargar_en_memoria(self):
        if self._wb_cache is not None:
            try:
                self._wb_cache.close()
            except Exception:
                pass
            self._wb_cache = None

        if os.path.exists(self.ruta):
            try:
                self._wb_cache = openpyxl.load_workbook(self.ruta, data_only=True)
            except Exception:
                self._wb_cache = None"""

cargar_replacement = """    def _cargar_en_memoria(self):
        if self._wb_cache is not None:
            try:
                self._wb_cache.close()
            except Exception:
                pass
            self._wb_cache = None
            self._last_mtime = 0.0

        if os.path.exists(self.ruta):
            try:
                self._wb_cache = openpyxl.load_workbook(self.ruta, data_only=True)
                self._last_mtime = os.path.getmtime(self.ruta)
            except Exception:
                self._wb_cache = None
                self._last_mtime = 0.0

    def _verificar_y_recargar_cache(self):
        \"\"\"Invalida la caché del workbook si el archivo en disco ha cambiado.\"\"\"
        if os.path.exists(self.ruta):
            try:
                mtime_actual = os.path.getmtime(self.ruta)
                if self._wb_cache is None or mtime_actual > self._last_mtime:
                    self._cargar_en_memoria()
            except Exception:
                pass

    def _save_wb(self, wb):
        \"\"\"Guarda un libro de calificaciones de forma atómica (Write-then-Replace).\"\"\"
        import shutil
        import tempfile
        
        # 1. Backup de seguridad antes de guardar
        bak_ruta = self.ruta.replace(".xlsx", "_bak.xlsx")
        if os.path.exists(self.ruta):
            try:
                shutil.copy2(self.ruta, bak_ruta)
            except Exception:
                pass

        # 2. Guardado atómico
        temp_dir = os.path.dirname(os.path.abspath(self.ruta))
        fd, temp_path = tempfile.mkstemp(suffix=".tmp", dir=temp_dir)
        os.close(fd)

        try:
            wb.save(temp_path)
            os.replace(temp_path, self.ruta)
        except Exception as e:
            if os.path.exists(temp_path):
                try:
                    os.remove(temp_path)
                except Exception:
                    pass
            raise e"""

# Use string replace for cargar_en_memoria
content = content.replace(cargar_target, cargar_replacement)

# 3. Replace all wb.save(self.ruta) with self._save_wb(wb)
content = content.replace("wb.save(self.ruta)", "self._save_wb(wb)")
content = content.replace("wb_write.save(self.ruta)", "self._save_wb(wb_write)")

# 4. Inject _verificar_y_recargar_cache() call to reading/writing methods
# Let's target all: "wb = self._wb_cache if self._wb_cache else openpyxl.load_workbook(self.ruta, data_only=True)"
# and insert "self._verificar_y_recargar_cache()" right before them.
wb_pattern = r"(\s+)(wb\s*=\s*self\._wb_cache\s+if\s+self\._wb_cache\s+else\s+openpyxl\.load_workbook\(self\.ruta,\s*data_only=True\))"
content = re.sub(wb_pattern, r"\1self._verificar_y_recargar_cache()\1\2", content)

with open("src/rddata.py", "w", encoding="utf-8") as f:
    f.write(content)

print("Refactoring complete.")
