import re

with open("src/app.py", "r") as f:
    content = f.read()

# Remove the search bar from the header in app.py
search_bar_pattern = r"""        # Centro: barra de búsqueda
        centro = ctk\.CTkFrame\(hdr, fg_color="transparent"\)
        centro\.grid\(row=0, column=1, sticky="ew", padx=40\)
        centro\.columnconfigure\(0, weight=1\)

        self\._search_frame = ctk\.CTkFrame\(
            centro, fg_color=C\["input"\],
            border_width=1, border_color=C\["borde"\],
            corner_radius=20, height=34\)
        self\._search_frame\.grid\(row=0, column=0, sticky="ew"\)
        self\._search_frame\.grid_propagate\(False\)
        self\._search_frame\.columnconfigure\(1, weight=1\)

        ctk\.CTkLabel\(self\._search_frame, text="🔍", fg_color="transparent",
                     font=ctk\.CTkFont\(size=14\), text_color=C\["texto_sec"\],
                     width=30\)\.grid\(row=0, column=0, padx=\(10, 0\)\)

        self\._search_var = tk\.StringVar\(\)
        self\._search_var\.trace_add\("write", self\._buscar_estudiante\)
        self\._search_entry = ctk\.CTkEntry\(
            self\._search_frame, textvariable=self\._search_var,
            fg_color="transparent", border_width=0,
            placeholder_text="Buscar alumno o documento...",
            font=ctk\.CTkFont\("Segoe UI", 13\),
            text_color=C\["texto"\]\)
        self\._search_entry\.grid\(row=0, column=1, sticky="ew",
                                padx=\(4, 10\), pady=4\)

        # Panel de resultados \(oculto por defecto\)
        self\._resultados_frame = ctk\.CTkFrame\(
            centro, fg_color=C\["card"\],
            border_width=1, border_color=self\._acento,
            corner_radius=8\)
        # no se muestra hasta que haya búsqueda"""

new_search_bar = """        # Centro: barra de búsqueda (Moved to DashApp)
        centro = ctk.CTkFrame(hdr, fg_color="transparent")
        centro.grid(row=0, column=1, sticky="ew", padx=40)
        centro.columnconfigure(0, weight=1)"""

content = re.sub(search_bar_pattern, new_search_bar, content, flags=re.DOTALL)

# Remove the related search methods from app.py
search_methods_pattern = r"""    def _buscar_estudiante\(self, \*_\):
.*?
    def _seleccionar_resultado\(self, nombre\):
        self\._search_var\.set\(nombre\)
        self\._resultados_frame\.grid_remove\(\)"""

content = re.sub(search_methods_pattern, "", content, flags=re.DOTALL)

with open("src/app.py", "w") as f:
    f.write(content)
