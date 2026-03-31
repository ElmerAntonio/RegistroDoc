import re

with open("src/rddata.py", "r") as f:
    content = f.read()

# Make DataEngine use a singleton-like pre-loaded workbook for reading
# Add self._wb_cache to initialization and handle it
init_pattern = r"""class DataEngine:
    def __init__\(self, ruta_excel, modalidad="premedia"\):
        self\.ruta = ruta_excel
        self\.modalidad = modalidad\.lower\(\)
        self\.cleaner = ExcelCleaner\(\)
        self\.fila_desc = 39 if self\.modalidad == "primaria" else 42"""

new_init = """class DataEngine:
    def __init__(self, ruta_excel, modalidad="premedia"):
        self.ruta = ruta_excel
        self.modalidad = modalidad.lower()
        self.cleaner = ExcelCleaner()
        self.fila_desc = 39 if self.modalidad == "primaria" else 42
        self._wb_cache = None
        self._cargar_en_memoria()

    def _cargar_en_memoria(self):
        if os.path.exists(self.ruta):
            try:
                self._wb_cache = openpyxl.load_workbook(self.ruta, data_only=True)
            except Exception:
                self._wb_cache = None"""

content = re.sub(init_pattern, new_init, content, flags=re.DOTALL)


# Replace openpyxl.load_workbook() calls in read methods to use self._wb_cache
# Example: _obtener_columna_nombres
col_nombres_pattern = r"""    def _obtener_columna_nombres\(self, grado, wb=None\):
        if not os\.path\.exists\(self\.ruta\) and wb is None: return 2

        should_close = False
        if wb is None:
            wb = openpyxl\.load_workbook\(self\.ruta, data_only=True\)
            should_close = True"""

new_col_nombres = """    def _obtener_columna_nombres(self, grado, wb=None):
        if not os.path.exists(self.ruta) and wb is None and self._wb_cache is None: return 2

        should_close = False
        if wb is None:
            if self._wb_cache:
                wb = self._wb_cache
            else:
                wb = openpyxl.load_workbook(self.ruta, data_only=True)
                should_close = True"""

content = re.sub(col_nombres_pattern, new_col_nombres, content, flags=re.DOTALL)


# General function optimization: replace explicit wb close in read-only methods if it's cached
datos_generales_pattern = r"""        if not os\.path\.exists\(self\.ruta\): return datos
        wb = openpyxl\.load_workbook\(self\.ruta, data_only=True\)"""

new_datos_generales = """        if not os.path.exists(self.ruta) and self._wb_cache is None: return datos
        wb = self._wb_cache if self._wb_cache else openpyxl.load_workbook(self.ruta, data_only=True)
        should_close = False if self._wb_cache else True"""

content = re.sub(datos_generales_pattern, new_datos_generales, content, flags=re.DOTALL)

# Fix wb.close() inside _obtener_datos_generales
close_general = r"""                break

        wb\.close\(\)
        return datos"""

new_close_general = """                break

        if should_close: wb.close()
        return datos"""
content = re.sub(close_general, new_close_general, content, flags=re.DOTALL)


# Fix `obtener_horario`
horario_pattern = r"""        if not os\.path\.exists\(self\.ruta\): return horario

        wb = openpyxl\.load_workbook\(self\.ruta, data_only=True\)"""
new_horario = """        if not os.path.exists(self.ruta) and self._wb_cache is None: return horario

        wb = self._wb_cache if self._wb_cache else openpyxl.load_workbook(self.ruta, data_only=True)
        should_close = False if self._wb_cache else True"""
content = re.sub(horario_pattern, new_horario, content, flags=re.DOTALL)

close_horario = r"""                    idx \+= 1

        wb\.close\(\)
        return horario"""
new_close_horario = """                    idx += 1

        if should_close: wb.close()
        return horario"""
content = re.sub(close_horario, new_close_horario, content, flags=re.DOTALL)

# And similarly for `obtener_grados_activos`, `obtener_materias_por_grado`, `obtener_estudiantes_completos`, `get_dashboard_stats`
def regex_replace(content, pattern, replacement):
    return re.sub(pattern, replacement, content, flags=re.DOTALL)

content = regex_replace(content,
r"""    def obtener_grados_activos\(self, wb=None\):
        if not os\.path\.exists\(self\.ruta\) and wb is None: return \[\]

        should_close = False
        if wb is None:
            wb = openpyxl\.load_workbook\(self\.ruta, read_only=True\)
            should_close = True""",
"""    def obtener_grados_activos(self, wb=None):
        if not os.path.exists(self.ruta) and wb is None and self._wb_cache is None: return []

        should_close = False
        if wb is None:
            if self._wb_cache:
                wb = self._wb_cache
            else:
                wb = openpyxl.load_workbook(self.ruta, read_only=True)
                should_close = True""")

content = regex_replace(content,
r"""    def obtener_materias_por_grado\(self, grado\):
        if not os\.path\.exists\(self\.ruta\): return \[\]
        wb = openpyxl\.load_workbook\(self\.ruta, read_only=True\)""",
"""    def obtener_materias_por_grado(self, grado):
        if not os.path.exists(self.ruta) and self._wb_cache is None: return []
        wb = self._wb_cache if self._wb_cache else openpyxl.load_workbook(self.ruta, read_only=True)
        should_close = False if self._wb_cache else True""")

content = regex_replace(content, r"""                        materias\.append\(mat\.title\(\)\)
        wb\.close\(\)
        return""",
"""                        materias.append(mat.title())
        if should_close: wb.close()
        return""")

content = regex_replace(content,
r"""    def obtener_estudiantes_completos\(self, grado, wb=None\):
        if not os\.path\.exists\(self\.ruta\) and wb is None: return \[\]

        should_close = False
        if wb is None:
            wb = openpyxl\.load_workbook\(self\.ruta, data_only=True\)
            should_close = True""",
"""    def obtener_estudiantes_completos(self, grado, wb=None):
        if not os.path.exists(self.ruta) and wb is None and self._wb_cache is None: return []

        should_close = False
        if wb is None:
            if self._wb_cache:
                wb = self._wb_cache
            else:
                wb = openpyxl.load_workbook(self.ruta, data_only=True)
                should_close = True""")


content = regex_replace(content,
r"""    def get_dashboard_stats\(self\):
        if not os\.path\.exists\(self\.ruta\): return \{"total": 0, "riesgo": 0, "honor": "N/A", "asistencia": "0%"\}

        # Optimization: Use a single workbook load for all dashboard stats
        wb = openpyxl\.load_workbook\(self\.ruta, data_only=True\)
        try:""",
"""    def get_dashboard_stats(self):
        if not os.path.exists(self.ruta) and self._wb_cache is None: return {"total": 0, "riesgo": 0, "honor": "N/A", "asistencia": "0%"}

        # Optimization: Use a single workbook load for all dashboard stats
        wb = self._wb_cache if self._wb_cache else openpyxl.load_workbook(self.ruta, data_only=True)
        should_close = False if self._wb_cache else True
        try:""")

content = regex_replace(content, r"""            total = sum\(len\(self\.obtener_estudiantes_completos\(g, wb=wb\)\) for g in grados\)
        finally:
            wb\.close\(\)""",
"""            total = sum(len(self.obtener_estudiantes_completos(g, wb=wb)) for g in grados)
        finally:
            if should_close: wb.close()""")

with open("src/rddata.py", "w") as f:
    f.write(content)
