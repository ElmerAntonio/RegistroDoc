import re

with open("src/app.py", "r", encoding="utf-8") as f:
    content = f.read()

# 1. Agregar la importación de grapp.py
content = re.sub(
    r'from sapp   import ConfigFrame\n',
    'from sapp   import ConfigFrame\nfrom grapp  import GraficosFrame\n',
    content
)

# 2. Agregar "Gráficos" a la lista items en _sb_renderizar
content = re.sub(
    r'\("📋", "Reportes",\s*self\._ir_reportes,\s*False\),\n\s*\("⚙️", "Configuración", self\._ir_configuracion, False\)',
    '("📋", "Reportes",      self._ir_reportes,     False),\n            ("📊", "Gráficos",      self._ir_graficos,     False),\n            ("⚙️", "Configuración", self._ir_configuracion, False)',
    content
)

# 3. Agregar la función _ir_graficos
content = re.sub(
    r'    def _ir_reportes\(self\):\n        pass\n',
    '    def _ir_reportes(self):\n        pass\n\n    def _ir_graficos(self):\n        if self.app:\n            try: self.app.mostrar_graficos()\n            except Exception: pass\n',
    content
)

# 4. Agregar la función mostrar_graficos
content = re.sub(
    r'    def mostrar_configuracion\(self\):\n        self\.limpiar_pantalla\(\)\n        ConfigFrame\(self\.main_app\.main_content_frame,\n                    self\.engine, self\)\.pack\(fill="both", expand=True\)\n',
    '    def mostrar_graficos(self):\n        self.limpiar_pantalla()\n        GraficosFrame(self.main_app.main_content_frame,\n                      self.engine).pack(fill="both", expand=True)\n\n    def mostrar_configuracion(self):\n        self.limpiar_pantalla()\n        ConfigFrame(self.main_app.main_content_frame,\n                    self.engine, self).pack(fill="both", expand=True)\n',
    content
)


with open("src/app.py", "w", encoding="utf-8") as f:
    f.write(content)
