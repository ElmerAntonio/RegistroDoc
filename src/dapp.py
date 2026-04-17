import customtkinter as ctk
import tkinter.messagebox as messagebox

class EstudiantesFrame(ctk.CTkFrame):
    def __init__(self, master, engine, **kwargs):
        super().__init__(master, fg_color="transparent", **kwargs)
        self.engine = engine
        self.entradas_estudiantes = {}

        ctk.CTkLabel(self, text="Gestión de Estudiantes", font=("Segoe UI", 24, "bold")).pack(anchor="w", pady=(0, 10))

        # --- PANEL NUEVO: AGREGAR ESTUDIANTE (Verde) ---
        self.frame_agregar = ctk.CTkFrame(self, fg_color="#10B981", corner_radius=8)
        self.frame_agregar.pack(fill="x", pady=5, ipadx=10, ipady=10)
        
        ctk.CTkLabel(self.frame_agregar, text="➕ Agregar Alumno:", font=("Segoe UI", 14, "bold"), text_color="white").pack(side="left", padx=10)
        
        self.entry_nuevo_nombre = ctk.CTkEntry(self.frame_agregar, width=250, placeholder_text="APELLIDO, Nombre (Requerido)")
        self.entry_nuevo_nombre.pack(side="left", padx=10)
        
        self.entry_nueva_cedula = ctk.CTkEntry(self.frame_agregar, width=150, placeholder_text="Cédula (Opcional)")
        self.entry_nueva_cedula.pack(side="left", padx=10)
        
        ctk.CTkButton(self.frame_agregar, text="Guardar Nuevo", fg_color="#059669", hover_color="#047857", command=self.agregar_nuevo).pack(side="left", padx=10)

        # --- PANEL DE CONTROLES (Filtro por grado) ---
        self.frame_controles = ctk.CTkFrame(self, fg_color="#1E2D42", corner_radius=8)
        self.frame_controles.pack(fill="x", pady=10, ipadx=10, ipady=10)

        ctk.CTkLabel(self.frame_controles, text="Seleccione Grado:", font=("Segoe UI", 14)).pack(side="left", padx=10)
        
        opciones_grado = self.engine.obtener_grados_activos() or ["Sin datos"]
        self.combo_grado = ctk.CTkOptionMenu(self.frame_controles, values=opciones_grado, command=self.cargar_lista)
        self.combo_grado.pack(side="left", padx=10)

        # --- ENCABEZADOS DE LA LISTA ---
        self.frame_encabezados = ctk.CTkFrame(self, fg_color="#253650", corner_radius=5)
        self.frame_encabezados.pack(fill="x", pady=(10, 0), ipady=5)
        
        ctk.CTkLabel(self.frame_encabezados, text="N°", width=40, font=("Segoe UI", 14, "bold")).pack(side="left", padx=10)
        ctk.CTkLabel(self.frame_encabezados, text="APELLIDO, NOMBRE (Editable)", width=350, anchor="w", font=("Segoe UI", 14, "bold")).pack(side="left", padx=10)
        ctk.CTkLabel(self.frame_encabezados, text="CÉDULA (Editable)", width=150, anchor="w", font=("Segoe UI", 14, "bold")).pack(side="left", padx=10)

        self.scroll_lista = ctk.CTkScrollableFrame(self, fg_color="transparent")
        self.scroll_lista.pack(fill="both", expand=True, pady=5)

        # Botón Guardar Cambios Masivos
        self.btn_guardar = ctk.CTkButton(self, text="💾 GUARDAR MODIFICACIONES DE LA LISTA", fg_color="#3B82F6", hover_color="#2563EB", 
                                         font=("Segoe UI", 15, "bold"), height=40, command=self.guardar_cambios)
        self.btn_guardar.pack(pady=10)

        self.cargar_lista(self.combo_grado.get())

    def cargar_lista(self, grado):
        for w in self.scroll_lista.winfo_children(): w.destroy()
        self.entradas_estudiantes.clear()
        
        estudiantes = self.engine.obtener_estudiantes_completos(grado)

        if not estudiantes:
            ctk.CTkLabel(self.scroll_lista, text=f"La lista de {grado} está vacía. ¡Agrega un alumno arriba ☝️!", text_color="#94A3B8").pack(pady=20)
            return

        for est in estudiantes:
            fila = ctk.CTkFrame(self.scroll_lista, fg_color="#1A2638", corner_radius=5)
            fila.pack(fill="x", pady=2)

            ctk.CTkLabel(fila, text=str(est["id"]), width=40).pack(side="left", padx=10)
            
            # Nombre ahora es Editable
            entry_nom = ctk.CTkEntry(fila, width=350, fg_color="#0F1923")
            entry_nom.insert(0, est["nombre"])
            entry_nom.pack(side="left", padx=10)

            # Cédula Editable
            entry_ced = ctk.CTkEntry(fila, width=150, fg_color="#0F1923", placeholder_text="Ej: 4-123-456")
            if est["cedula"]: entry_ced.insert(0, est["cedula"])
            entry_ced.pack(side="left", padx=10)
            
            self.entradas_estudiantes[est["id"]] = {"nombre": entry_nom, "cedula": entry_ced}

    def agregar_nuevo(self):
        grado = self.combo_grado.get()
        nombre = self.entry_nuevo_nombre.get().strip().upper()
        cedula = self.entry_nueva_cedula.get().strip()

        if not nombre:
            messagebox.showwarning("Faltan datos", "El Apellido y Nombre son obligatorios.")
            return

        # Friendly Limit Warning
        max_estudiantes = 34 if self.engine.modalidad == "primaria" else 36
        actuales = self.engine.obtener_estudiantes_completos(grado)
        if len(actuales) >= max_estudiantes:
            messagebox.showwarning("Límite Alcanzado", f"El grado {grado} ha alcanzado el límite máximo de {max_estudiantes} estudiantes permitido para {self.engine.modalidad.capitalize()}.")
            return

        exito = self.engine.agregar_estudiante(grado, nombre, cedula)
        if exito:
            messagebox.showinfo("Éxito", f"Estudiante {nombre} agregado correctamente.")
            self.entry_nuevo_nombre.delete(0, 'end')
            self.entry_nueva_cedula.delete(0, 'end')
            self.cargar_lista(grado) # Recargar la lista para que aparezca
        else:
            messagebox.showerror("Error", f"No se pudo agregar. La lista podría estar llena (Máx. {max_estudiantes} alumnos).")

    def guardar_cambios(self):
        grado = self.combo_grado.get()
        datos_modificados = {}
        
        for id_est, entries in self.entradas_estudiantes.items():
            nom = entries["nombre"].get().strip().upper()
            ced = entries["cedula"].get().strip()
            
            if nom: # Solo guarda si el nombre no está en blanco
                datos_modificados[id_est] = {"nombre": nom, "cedula": ced}

        if self.engine.guardar_cambios_estudiantes(grado, datos_modificados):
            messagebox.showinfo("Éxito", "Nombres y cédulas guardados correctamente en el Excel.")
            self.cargar_lista(grado)
        else:
            messagebox.showerror("Error", "Hubo un problema al guardar los cambios.")