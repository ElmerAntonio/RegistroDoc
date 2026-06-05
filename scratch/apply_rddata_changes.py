import re
import os

with open("src/rddata.py", "r", encoding="utf-8") as f:
    content = f.read()

# 1. Update __init__ to include _last_mtime
init_pattern = r"(self\._wb_cache\s*=\s*None\n\s*self\._cargar_en_memoria\(\))"
init_replacement = r"self._wb_cache = None\n        self._last_mtime = 0.0\n        self._cargar_en_memoria()"
content = re.sub(init_pattern, init_replacement, content)

# 2. Update _cargar_en_memoria to update _last_mtime
cargar_pattern = r"def _cargar_en_memoria\(self\):.*?self\._wb_cache\s*=\s*None"
# Let's see the exact _cargar_en_memoria implementation to do a clean replace.
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

        if os.path.exists(self.ruta):
            try:
                self._wb_cache = openpyxl.load_workbook(self.ruta, data_only=True)
                self._last_mtime = os.path.getmtime(self.ruta)
            except Exception:
                self._wb_cache = None
                self._last_mtime = 0.0"""

content = content.replace(cargar_target, cargar_replacement)

# 3. Add _verificar_y_recargar_cache and _save_wb methods after _cargar_en_memoria
new_methods = """
    def _verificar_y_recargar_cache(self):
        if os.path.exists(self.ruta):
            try:
                current_mtime = os.path.getmtime(self.ruta)
                if current_mtime != self._last_mtime:
                    self._cargar_en_memoria()
            except Exception:
                pass

    def _save_wb(self, wb):
        import shutil
        import tempfile

        # 1. Backup de seguridad antes de guardar
        bak_ruta = self.ruta.replace(".xlsx", "_bak.xlsx")
        if os.path.exists(self.ruta):
            try:
                shutil.copy2(self.ruta, bak_ruta)
            except Exception:
                pass

        # 2. Guardar a través de archivo temporal de forma atómica
        dir_base = os.path.dirname(os.path.abspath(self.ruta))
        fd, temp_path = tempfile.mkstemp(suffix=".xlsx", dir=dir_base)
        os.close(fd)

        try:
            wb.save(temp_path)
            # Reemplazar de forma atómica
            if os.path.exists(self.ruta):
                os.replace(temp_path, self.ruta)
            else:
                shutil.move(temp_path, self.ruta)
        except Exception as e:
            if os.path.exists(temp_path):
                try: os.remove(temp_path)
                except: pass
            raise e

        # Actualizar mtime para evitar recarga de caché redundante por nuestra propia escritura
        try:
            self._last_mtime = os.path.getmtime(self.ruta)
        except Exception:
            self._last_mtime = 0.0
"""

# Place new_methods right after _cargar_en_memoria
content = content.replace(cargar_replacement, cargar_replacement + new_methods)

# 4. Inject self._verificar_y_recargar_cache() in methods taking wb=None
# Find matches like: def name(self, ..., wb=None):
# and insert:
#         if wb is None:
#             self._verificar_y_recargar_cache()
# right after the def statement.
def inject_cache_check(match):
    full_def = match.group(0)
    indent = "    "
    # Find number of spaces of indentation of the def
    # Let's count leading spaces of the match
    m = re.match(r"^([ \t]*)def", full_def)
    if m:
        indent = m.group(1) + "    "
    
    # We want to add the check after the def line
    check_code = f"\n{indent}if wb is None:\n{indent}    self._verificar_y_recargar_cache()"
    return full_def + check_code

# Regex to match def lines of methods with wb=None
pattern = r"^[ \t]*def\s+\w+\([^)]*wb=None[^)]*\):"
content = re.sub(pattern, inject_cache_check, content, flags=re.MULTILINE)

# 5. Replace all wb.save(self.ruta) and wb_write.save(self.ruta) with self._save_wb(wb)
content = content.replace("wb.save(self.ruta)", "self._save_wb(wb)")
content = content.replace("wb_write.save(self.ruta)", "self._save_wb(wb_write)")

with open("src/rddata.py", "w", encoding="utf-8") as f:
    f.write(content)

print("Applied all changes successfully!")
