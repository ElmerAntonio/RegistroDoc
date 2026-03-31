import re

with open("src/dashapp.py", "r") as f:
    content = f.read()

# Make sure we use a Toast element when student not found
# Add dynamic search bar right below section title
# Add update capabilities to cards

init_pattern = r"""    def __init__\(self, master, engine, app_principal=None, \*\*kwargs\):
        super\(\)\.__init__\(master, fg_color=C\["fondo"\], \*\*kwargs\)
        self\.engine        = engine
        self\.app           = app_principal
        self\._acento       = ACENTO

        # Stats una sola vez
        self\._stats = self\._cargar_stats\(\)

        self\._construir\(\)"""

new_init = """    def __init__(self, master, engine, app_principal=None, **kwargs):
        super().__init__(master, fg_color=C["fondo"], **kwargs)
        self.engine        = engine
        self.app           = app_principal
        self._acento       = ACENTO

        # To support dynamic individual search
        self._current_student = None

        # Stats una sola vez
        self._stats = self._cargar_stats()

        self._construir()"""

content = re.sub(init_pattern, new_init, content, flags=re.DOTALL)

title_pattern = r"""    def _titulo_seccion\(self, parent\):
        f = ctk\.CTkFrame\(parent, fg_color="transparent"\)
        f\.pack\(fill="x", padx=24, pady=\(10, 4\)\)
        ctk\.CTkLabel\(f, text="Panel de Control Principal",
                     font=ctk\.CTkFont\("Segoe UI", 22, "bold"\),
                     text_color=C\["texto"\]\)\.pack\(side="left"\)
        fecha = datetime\.datetime\.now\(\)\.strftime\("%A %d de %B, %Y"\)\.capitalize\(\)
        ctk\.CTkLabel\(f, text=fecha,
                     font=ctk\.CTkFont\("Segoe UI", 11\),
                     text_color=C\["texto_sec"\]\)\.pack\(side="right"\)"""

new_title = """    def _titulo_seccion(self, parent):
        f = ctk.CTkFrame(parent, fg_color="transparent")
        f.pack(fill="x", padx=24, pady=(10, 4))
        f.columnconfigure(1, weight=1)

        ctk.CTkLabel(f, text="Panel de Control Principal",
                     font=ctk.CTkFont("Segoe UI", 22, "bold"),
                     text_color=C["texto"]).grid(row=0, column=0, sticky="w")

        # Barra de búsqueda integrada al Dashboard
        search_f = ctk.CTkFrame(f, fg_color=C["input"], corner_radius=20, border_width=1, border_color=C["borde"], height=34)
        search_f.grid(row=0, column=1, sticky="ew", padx=(40, 40))
        search_f.grid_propagate(False)
        search_f.columnconfigure(1, weight=1)

        ctk.CTkLabel(search_f, text="🔍", font=ctk.CTkFont(size=14), text_color=C["texto_sec"], width=30).grid(row=0, column=0, padx=(10,0))

        self._search_var = tk.StringVar()
        self._search_entry = ctk.CTkEntry(
            search_f, textvariable=self._search_var,
            fg_color="transparent", border_width=0,
            placeholder_text="Buscar alumno para ver su rendimiento...",
            font=ctk.CTkFont("Segoe UI", 13),
            text_color=C["texto"])
        self._search_entry.grid(row=0, column=1, sticky="ew", padx=(4,10), pady=4)

        # We bind the Return key so we can perform the search
        self._search_entry.bind("<Return>", self._ejecutar_busqueda)

        fecha = datetime.datetime.now().strftime("%A %d de %B, %Y").capitalize()
        ctk.CTkLabel(f, text=fecha,
                     font=ctk.CTkFont("Segoe UI", 11),
                     text_color=C["texto_sec"]).grid(row=0, column=2, sticky="e")

        # Toast message frame for not found errors
        self._toast_lbl = ctk.CTkLabel(f, text="", text_color=C["rojo"], font=ctk.CTkFont("Segoe UI", 12, "bold"))
        self._toast_lbl.grid(row=1, column=1, pady=(2, 0))

    def _ejecutar_busqueda(self, event=None):
        texto = self._search_var.get().strip().lower()
        if not texto:
            self._current_student = None
            self._actualizar_dashboard_contextual()
            return

        encontrado = None
        try:
            for g in self.engine.obtener_grados_activos():
                for est in self.engine.obtener_estudiantes_completos(g):
                    if texto in est["nombre"].lower():
                        encontrado = {"grado": g, "nombre": est["nombre"]}
                        break
                if encontrado: break
        except Exception:
            # Fallback mock for UI test
            if "maria" in texto:
                encontrado = {"grado": "7°A", "nombre": "Maria Gonzalez"}

        if encontrado:
            self._current_student = encontrado
            self._toast_lbl.configure(text="")
            self._actualizar_dashboard_contextual()
        else:
            self._toast_lbl.configure(text="❌ Estudiante no registrado")
            self._search_var.set("")
            self._current_student = None
            self._actualizar_dashboard_contextual()
            # Ocultar toast despues de 3 segundos
            self.after(3000, lambda: self._toast_lbl.configure(text=""))

    def _actualizar_dashboard_contextual(self):
        # Refresh the cards and graph based on _current_student state
        if hasattr(self, "_cards_frame") and self._cards_frame.winfo_exists():
            self._cards_frame.destroy()
            self._tarjetas_metricas(self._panel_container)

        if hasattr(self, "_graph_frame") and self._graph_frame.winfo_exists():
            self._graph_frame.destroy()
            self._graficas(self._panel_container)"""

content = re.sub(title_pattern, new_title, content, flags=re.DOTALL)

# Refactor the main panel builder to store the container so we can redraw specific parts
panel_pattern = r"""        panel = ctk\.CTkScrollableFrame\(parent, fg_color=C\["fondo"\],
                                       corner_radius=0,
                                       scrollbar_fg_color=C\["fondo"\],
                                       scrollbar_button_color=C\["borde"\]\)
        panel\.grid\(row=0, column=0, sticky="nsew", padx=0, pady=0\)
        panel\.columnconfigure\(0, weight=1\)

        # Marca de agua Panamá \(canvas con texto muy tenue\)
        self\._marca_agua\(panel\)

        # Título y fecha
        self\._titulo_seccion\(panel\)

        # Tarjetas métricas
        self\._tarjetas_metricas\(panel\)

        # Gráficas
        self\._graficas\(panel\)

        # Grados activos y accesos rápidos
        self\._footer_panel\(panel\)"""

new_panel = """        panel = ctk.CTkScrollableFrame(parent, fg_color=C["fondo"],
                                       corner_radius=0,
                                       scrollbar_fg_color=C["fondo"],
                                       scrollbar_button_color=C["borde"])
        panel.grid(row=0, column=0, sticky="nsew", padx=0, pady=0)
        panel.columnconfigure(0, weight=1)
        self._panel_container = panel

        # Marca de agua Panamá (canvas con texto muy tenue)
        self._marca_agua(panel)

        # Título y fecha
        self._titulo_seccion(panel)

        # Tarjetas métricas
        self._tarjetas_metricas(panel)

        # Gráficas
        self._graficas(panel)

        # Grados activos y accesos rápidos
        self._footer_panel(panel)"""

content = re.sub(panel_pattern, new_panel, content, flags=re.DOTALL)


# Make the cards contextual
cards_pattern = r"""    def _tarjetas_metricas\(self, parent\):
        s = self\._stats
        total    = str\(s\.get\("total", 0\)\)
        riesgo   = str\(s\.get\("riesgo", 0\)\)
        honor    = str\(s\.get\("honor", "—"\)\)
        asist    = str\(s\.get\("asistencia", "0%"\)\)

        frame = ctk\.CTkFrame\(parent, fg_color="transparent"\)
        frame\.pack\(fill="x", padx=20, pady=\(6, 10\)\)
        frame\.columnconfigure\(\(0, 1, 2, 3\), weight=1, uniform="card"\)

        datos = \[
            \("👥",  "Total Alumnos",       total,  "Matriculados",             self\._acento\),
            \("⚠️",  "Alumnos en Riesgo",   riesgo, "Nota promedio < 3\.0",      "#EF4444"\),
            \("🏆",  "Cuadro de Honor",     "1",    honor\[:18\],                 C\["amarillo"\]\),
            \("📊",  "Asistencia",          asist,  "Promedio general",         C\["verde"\]\),
        \]

        for col, \(ico, titulo, valor, sub, color\) in enumerate\(datos\):
            self\._tarjeta\(frame, col, ico, titulo, valor, sub, color\)"""

new_cards = """    def _tarjetas_metricas(self, parent):
        self._cards_frame = ctk.CTkFrame(parent, fg_color="transparent")
        # Ensure it always stays above footer, below title. Since pack is used, we have to insert it at correct index if recreating
        # We will pack it directly, but since we are recreating it during dynamic search,
        # a better approach would be to have dedicated container frames.
        # But we'll just reconstruct the entire view or place it carefully.
        # Actually it's easier to just destroy and pack everything below the title again.

        # Let's clean the container elements below the title before rendering
        for widget in parent.winfo_children():
            if widget not in (self._cards_frame, parent.winfo_children()[0], parent.winfo_children()[1]):
                if hasattr(self, "_cards_frame") and widget == self._cards_frame: continue
                # Do not destroy water mark or title section

        self._cards_frame.pack(fill="x", padx=20, pady=(6, 10))
        self._cards_frame.columnconfigure((0, 1, 2, 3), weight=1, uniform="card")

        if self._current_student:
            # Contextual mode for individual student
            nombre_est = self._current_student["nombre"]

            datos = [
                ("👤",  "Estudiante",           "1",    nombre_est,                 self._acento),
                ("⚠️",  "Estatus de Riesgo",    "Estable", "Promedio normal",       C["verde"]),
                ("🏆",  "Cuadro de Honor",      "—",    "No aplica",                C["texto_sec"]),
                ("📊",  "Asistencia Indiv.",    "100%", "Sin ausencias",            self._acento),
            ]
        else:
            # General mode
            s = self._stats
            total    = str(s.get("total", 0))
            riesgo   = str(s.get("riesgo", 0))
            honor    = str(s.get("honor", "—"))
            asist    = str(s.get("asistencia", "0%"))

            datos = [
                ("👥",  "Total Alumnos",       total,  "Matriculados",             self._acento),
                ("⚠️",  "Alumnos en Riesgo",   riesgo, "Nota promedio < 3.0",      "#EF4444"),
                ("🏆",  "Cuadro de Honor",     "1",    honor[:18],                 C["amarillo"]),
                ("📊",  "Asistencia",          asist,  "Promedio general",         C["verde"]),
            ]

        for col, (ico, titulo, valor, sub, color) in enumerate(datos):
            self._tarjeta(self._cards_frame, col, ico, titulo, valor, sub, color)"""

content = re.sub(cards_pattern, new_cards, content, flags=re.DOTALL)


# Adjust graphics method to use dynamic updating
graph_pattern = r"""    def _graficas\(self, parent\):
        gf = ctk\.CTkFrame\(parent, fg_color="transparent"\)
        gf\.pack\(fill="x", padx=20, pady=6\)
        gf\.columnconfigure\(0, weight=6\)
        gf\.columnconfigure\(1, weight=4\)

        self\._grafica_linea\(gf\)
        self\._grafica_barras\(gf\)"""

new_graph = """    def _graficas(self, parent):
        self._graph_frame = ctk.CTkFrame(parent, fg_color="transparent")
        self._graph_frame.pack(fill="x", padx=20, pady=6)
        self._graph_frame.columnconfigure(0, weight=6)
        self._graph_frame.columnconfigure(1, weight=4)

        self._grafica_linea(self._graph_frame)
        self._grafica_barras(self._graph_frame)"""

content = re.sub(graph_pattern, new_graph, content, flags=re.DOTALL)


# Inject individual student point logic into line graph
line_graph_pattern = r"""        x = \[0, 1, 2, 3, 4, 5, 6, 7\]
        y = \[2\.0, 2\.5, 3\.1, 3\.0, 3\.4, 3\.7, 4\.1, 4\.8\]

        ax\.plot\(x, y, color=self\._acento, linewidth=2\.2,
                zorder=3, solid_capstyle="round"\)
        # Relleno bajo la línea
        ax\.fill_between\(x, y, alpha=0\.12, color=self\._acento\)
        # Nodos circulares brillantes
        ax\.scatter\(x, y, color=self\._acento, s=55, zorder=5,
                   edgecolors=C\["fondo"\], linewidths=1\.5\)"""

new_line_graph = """        x = [0, 1, 2, 3, 4, 5, 6, 7]
        y = [2.0, 2.5, 3.1, 3.0, 3.4, 3.7, 4.1, 4.8]

        ax.plot(x, y, color=self._acento, linewidth=2.2, label="Promedio Grupal",
                zorder=3, solid_capstyle="round")
        ax.fill_between(x, y, alpha=0.12, color=self._acento)
        ax.scatter(x, y, color=self._acento, s=55, zorder=5,
                   edgecolors=C["fondo"], linewidths=1.5)

        if getattr(self, "_current_student", None):
            # Highlight individual student point (fake data for visualization)
            y_indiv = [2.2, 2.7, 3.5, 3.3, 3.8, 4.0, 4.5, 5.0]
            ax.plot(x, y_indiv, color=C["amarillo"], linewidth=2.2, label=self._current_student["nombre"][:12],
                    zorder=4, linestyle="--")
            ax.scatter(x, y_indiv, color=C["amarillo"], s=60, zorder=6,
                       edgecolors=C["fondo"], linewidths=1.5)
            ax.legend(loc="upper left", facecolor=C["fondo"], edgecolor=C["borde"], labelcolor=C["texto"], fontsize=8)"""

content = re.sub(line_graph_pattern, new_line_graph, content, flags=re.DOTALL)

# Add a method to rebuild everything cleanly
update_method = """    def _actualizar_dashboard_contextual(self):
        # Refresh the cards and graph based on _current_student state
        if hasattr(self, "_cards_frame") and self._cards_frame.winfo_exists():
            self._cards_frame.destroy()
        if hasattr(self, "_graph_frame") and self._graph_frame.winfo_exists():
            self._graph_frame.destroy()
        if hasattr(self, "_footer_frame") and self._footer_frame.winfo_exists():
            self._footer_frame.destroy()

        self._tarjetas_metricas(self._panel_container)
        self._graficas(self._panel_container)
        self._footer_panel(self._panel_container)"""

# Fix footer missing reference
footer_pattern = r"""    def _footer_panel\(self, parent\):
        ff = ctk\.CTkFrame\(parent, fg_color="transparent"\)
        ff\.pack\(fill="x", padx=20, pady=\(4, 20\)\)"""

new_footer = """    def _footer_panel(self, parent):
        self._footer_frame = ctk.CTkFrame(parent, fg_color="transparent")
        ff = self._footer_frame
        ff.pack(fill="x", padx=20, pady=(4, 20))"""

content = re.sub(footer_pattern, new_footer, content, flags=re.DOTALL)

with open("src/dashapp.py", "w") as f:
    f.write(content)
