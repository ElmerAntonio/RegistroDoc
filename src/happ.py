import customtkinter as ctk

class ReportesFrame(ctk.CTkFrame):
    def __init__(self, master, engine, **kwargs):
        super().__init__(master, fg_color="transparent", **kwargs)
        self.engine = engine
        
        self.pack_propagate(False)

        ctk.CTkLabel(self, text="Reportes y Estadísticas", font=("Segoe UI", 24, "bold")).pack(anchor="w", pady=(0, 10))

        # Panel de Controles
        self.frame_controles = ctk.CTkFrame(self, fg_color="#1E2D42", corner_radius=8)
        self.frame_controles.pack(fill="x", pady=5, ipadx=10, ipady=10)

        ctk.CTkLabel(self.frame_controles, text="Seleccione Grado:", font=("Segoe UI", 14)).pack(side="left", padx=10)

        opciones_grado = self.engine.obtener_grados_activos()
        if not opciones_grado: opciones_grado = ["7°", "8°", "9°"]
        self.combo_grado = ctk.CTkOptionMenu(self.frame_controles, values=opciones_grado, command=self.cargar_reportes)
        self.combo_grado.pack(side="left", padx=10)

        # Tabs for different reports
        self.tabs = ctk.CTkTabview(self)
        self.tabs.pack(fill="both", expand=True, pady=10)

        self.tab_docente = self.tabs.add("Reporte Docente")
        self.tab_aprobados = self.tabs.add("Aprobados / Reprobados")
        self.tab_direccion = self.tabs.add("Reporte Dirección")

        # Docente
        self.scroll_docente = ctk.CTkScrollableFrame(self.tab_docente, fg_color="transparent")
        self.scroll_docente.pack(fill="both", expand=True)

        # Aprobados
        self.scroll_aprobados = ctk.CTkScrollableFrame(self.tab_aprobados, fg_color="transparent")
        self.scroll_aprobados.pack(fill="both", expand=True)

        # Direccion
        self.scroll_direccion = ctk.CTkScrollableFrame(self.tab_direccion, fg_color="transparent")
        self.scroll_direccion.pack(fill="both", expand=True)

        self.cargar_reportes(self.combo_grado.get())

    def cargar_reportes(self, grado):
        self._limpiar(self.scroll_docente)
        self._limpiar(self.scroll_aprobados)
        self._limpiar(self.scroll_direccion)



        reportes = getattr(self.engine, 'obtener_datos_reportes', lambda x: {"docente": [], "aprobados": [], "direccion": []})(grado)

        if not reportes["docente"]:
            ctk.CTkLabel(self.scroll_docente, text="No hay datos de resumen para mostrar.", font=("Segoe UI", 16)).pack(pady=20)
            ctk.CTkLabel(self.scroll_aprobados, text="No hay datos de resumen para mostrar.", font=("Segoe UI", 16)).pack(pady=20)
            ctk.CTkLabel(self.scroll_direccion, text="No hay datos de resumen para mostrar.", font=("Segoe UI", 16)).pack(pady=20)
            return

        # DOCENTE HEADER
        f_h_doc = ctk.CTkFrame(self.scroll_docente, fg_color="#253650")
        f_h_doc.pack(fill="x", pady=2)
        ctk.CTkLabel(f_h_doc, text="Estudiante", width=250, anchor="w", font=("Segoe UI", 13, "bold")).pack(side="left", padx=10)
        ctk.CTkLabel(f_h_doc, text="Cédula", width=150, anchor="w", font=("Segoe UI", 13, "bold")).pack(side="left", padx=10)
        ctk.CTkLabel(f_h_doc, text="Nota Final", width=100, anchor="w", font=("Segoe UI", 13, "bold")).pack(side="left", padx=10)

        for fila in reportes["docente"]:
            row_d = ctk.CTkFrame(self.scroll_docente, fg_color="#1A2638")
            row_d.pack(fill="x", pady=1)
            ctk.CTkLabel(row_d, text=fila[0], width=250, anchor="w").pack(side="left", padx=10)
            ctk.CTkLabel(row_d, text=fila[1] if fila[1] else "N/A", width=150, anchor="w").pack(side="left", padx=10)
            ctk.CTkLabel(row_d, text=str(fila[2]), width=100, anchor="w").pack(side="left", padx=10)

        # APROBADOS HEADER
        f_h_apr = ctk.CTkFrame(self.scroll_aprobados, fg_color="#253650")
        f_h_apr.pack(fill="x", pady=2)
        ctk.CTkLabel(f_h_apr, text="Estudiante", width=250, anchor="w", font=("Segoe UI", 13, "bold")).pack(side="left", padx=10)
        ctk.CTkLabel(f_h_apr, text="Estado", width=150, anchor="w", font=("Segoe UI", 13, "bold")).pack(side="left", padx=10)

        for fila in reportes["aprobados"]:
            row_a = ctk.CTkFrame(self.scroll_aprobados, fg_color="#1A2638")
            row_a.pack(fill="x", pady=1)
            ctk.CTkLabel(row_a, text=fila[0], width=250, anchor="w").pack(side="left", padx=10)
            color_est = "#10B981" if str(fila[1]).upper() == "APROBADO" else "#EF4444"
            ctk.CTkLabel(row_a, text=str(fila[1]), width=150, anchor="w", text_color=color_est, font=("Segoe UI", 12, "bold")).pack(side="left", padx=10)

        # DIRECCION HEADER
        f_h_dir = ctk.CTkFrame(self.scroll_direccion, fg_color="#253650")
        f_h_dir.pack(fill="x", pady=2)
        ctk.CTkLabel(f_h_dir, text="Estudiante", width=250, anchor="w", font=("Segoe UI", 13, "bold")).pack(side="left", padx=10)
        ctk.CTkLabel(f_h_dir, text="Cédula", width=150, anchor="w", font=("Segoe UI", 13, "bold")).pack(side="left", padx=10)
        ctk.CTkLabel(f_h_dir, text="Nota Final", width=100, anchor="w", font=("Segoe UI", 13, "bold")).pack(side="left", padx=10)
        ctk.CTkLabel(f_h_dir, text="Estado", width=150, anchor="w", font=("Segoe UI", 13, "bold")).pack(side="left", padx=10)

        for fila in reportes["direccion"]:
            row_dir = ctk.CTkFrame(self.scroll_direccion, fg_color="#1A2638")
            row_dir.pack(fill="x", pady=1)
            ctk.CTkLabel(row_dir, text=fila[0], width=250, anchor="w").pack(side="left", padx=10)
            ctk.CTkLabel(row_dir, text=fila[1] if fila[1] else "N/A", width=150, anchor="w").pack(side="left", padx=10)
            ctk.CTkLabel(row_dir, text=str(fila[2]), width=100, anchor="w").pack(side="left", padx=10)
            color_est = "#10B981" if str(fila[3]).upper() == "APROBADO" else "#EF4444"
            ctk.CTkLabel(row_dir, text=str(fila[3]), width=150, anchor="w", text_color=color_est, font=("Segoe UI", 12, "bold")).pack(side="left", padx=10)

    def _limpiar(self, frame):
        for w in frame.winfo_children(): w.destroy()
