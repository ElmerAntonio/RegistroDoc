import customtkinter as ctk
from tkinter import messagebox, filedialog
import os
from config import BASE_DIR, CONFIG_FILE
import json
import threading
import datetime


class ConfigFrame(ctk.CTkFrame):
    def __init__(self, master, engine, app_principal, **kwargs):
        super().__init__(master, fg_color="transparent", **kwargs)
        self.engine = engine
        self.app_principal = app_principal 

        ctk.CTkLabel(self, text="Panel de Control Maestro", font=("Segoe UI", 24, "bold"), text_color="#3B82F6").pack(anchor="center", pady=(0, 5))
        
        self.tabs = ctk.CTkTabview(self)
        self.tabs.pack(fill="both", expand=True, pady=10)
        
        self.tab_gen = self.tabs.add("Datos de Portada y Carátula")
        self.tab_hor = self.tabs.add("Horarios")
        self.tab_gra = self.tabs.add("Gestión de Grados")
        self.tab_mat = self.tabs.add("Gestión de Materias")

        self.crear_panel_general()
        self.crear_panel_horarios()
        self.crear_panel_grados()
        self.crear_panel_materias()
        self.actualizar_listas_ui() 

    # ================= PESTAÑA 1: DATOS (GRID RESPONSIVO) =================
    def crear_panel_general(self):
        data = self.engine.obtener_datos_generales()

        self.scroll_gen = ctk.CTkScrollableFrame(self.tab_gen, fg_color="transparent")
        self.scroll_gen.pack(fill="both", expand=True)

        f1 = ctk.CTkFrame(self.scroll_gen, fg_color="#1E2D42", corner_radius=10)
        f1.pack(fill="x", padx=10, pady=10, ipadx=10, ipady=10)
        ctk.CTkLabel(f1, text="Modalidad Activa:", font=("Segoe UI", 14, "bold")).pack(anchor="w", padx=20, pady=(10, 5))

        row_mod = ctk.CTkFrame(f1, fg_color="transparent")
        row_mod.pack(fill="x")
        self.var_modalidad = ctk.StringVar(value=self.engine.modalidad.capitalize())
        ctk.CTkRadioButton(row_mod, text="Premedia", variable=self.var_modalidad, value="Premedia").pack(side="left", padx=40)
        ctk.CTkRadioButton(row_mod, text="Primaria", variable=self.var_modalidad, value="Primaria").pack(side="left", padx=40)
        ctk.CTkButton(row_mod, text="Cambiar Modalidad", fg_color="#F59E0B", command=self.cambiar_modalidad).pack(side="right", padx=20)

        f2 = ctk.CTkFrame(self.scroll_gen, fg_color="#1A2638", corner_radius=10)
        f2.pack(fill="both", expand=True, padx=10, pady=10, ipadx=10, ipady=10)

        # Doble columna real: (label, entry) x 2
        f2.grid_columnconfigure(0, weight=0, minsize=250)
        f2.grid_columnconfigure(1, weight=1)
        f2.grid_columnconfigure(2, weight=0, minsize=250)
        f2.grid_columnconfigure(3, weight=1)

        ctk.CTkLabel(
            f2,
            text="Informacion Academica (Portada y Caratula)",
            font=("Segoe UI", 16, "bold"),
            text_color="#10B981",
        ).grid(row=0, column=0, columnspan=4, sticky="w", padx=20, pady=10)

        cargo_unificado = (
            str(data.get("condicion_nombramiento", "")).strip()
            or str(data.get("titulo_caratula", "")).strip()
            or "Instructor Vocacional"
        )
        self.var_condicion = ctk.StringVar(value=cargo_unificado)
        self.var_grupos_caratula = ctk.StringVar(value="")

        def crear_campo_en_columna(row_i, base_col, texto, attr_name, valor="", var=None, state="normal"):
            ctk.CTkLabel(
                f2,
                text=texto,
                anchor="w",
                font=("Segoe UI", 12, "bold"),
            ).grid(row=row_i, column=base_col, sticky="w", padx=10, pady=5)

            if var is not None:
                entry = ctk.CTkEntry(f2, textvariable=var, state=state)
            else:
                entry = ctk.CTkEntry(f2, state=state)
                entry.insert(0, valor)

            entry.grid(row=row_i, column=base_col + 1, sticky="ew", padx=20, pady=5)
            setattr(self, attr_name, entry)

        left_fields = [
            ("Nombre Docente:", "entry_doc", data.get("docente_nombre", ""), None, "normal"),
            ("Cedula:", "entry_ced", data.get("docente_cedula", ""), None, "normal"),
            ("No. Seguro Social:", "entry_ss", data.get("seguro_social", ""), None, "normal"),
            ("No. Posicion:", "entry_pos", data.get("numero_posicion", ""), None, "normal"),
            ("Escuela:", "entry_esc", data.get("escuela_nombre", ""), None, "normal"),
            ("Region Escolar:", "entry_reg", data.get("escuela_region", ""), None, "normal"),
            ("Distrito:", "entry_dis", data.get("distrito", ""), None, "normal"),
            ("Corregimiento:", "entry_correg", data.get("corregimiento", ""), None, "normal"),
            ("Zona Escolar:", "entry_zon", data.get("zona_escolar", ""), None, "normal"),
            ("Ano Lectivo:", "entry_ano", data.get("ano_lectivo", "2026"), None, "normal"),
            ("Jornada:", "entry_jor", data.get("jornada", ""), None, "normal"),
            ("Fecha Trimestre 1:", "entry_t1", data.get("fecha_t1", ""), None, "normal"),
            ("Fecha Trimestre 2:", "entry_t2", data.get("fecha_t2", ""), None, "normal"),
            ("Fecha Trimestre 3:", "entry_t3", data.get("fecha_t3", ""), None, "normal"),
        ]

        right_fields = [
            ("Director(a):", "entry_dir", data.get("director_nombre", ""), None, "normal"),
            ("Subdirector(a):", "entry_sub", data.get("subdirector_nombre", ""), None, "normal"),
            ("Coordinador:", "entry_coo", data.get("coordinador_nombre", ""), None, "normal"),
            ("Telefono:", "entry_tel", data.get("telefono", ""), None, "normal"),
            ("Correo:", "entry_mail", data.get("correo", ""), None, "normal"),
            ("Condicion/Titulo (Portada y Caratula):", "entry_con", "", self.var_condicion, "normal"),
            ("Grupos (Caratula):", "entry_gru", "", self.var_grupos_caratula, "disabled"),
        ]

        row_start = 1
        max_rows = max(len(left_fields), len(right_fields))

        for idx in range(max_rows):
            row_i = row_start + idx
            if idx < len(left_fields):
                t, a, v, var, st = left_fields[idx]
                crear_campo_en_columna(row_i, 0, t, a, v, var, st)
            if idx < len(right_fields):
                t, a, v, var, st = right_fields[idx]
                crear_campo_en_columna(row_i, 2, t, a, v, var, st)


        self.actualizar_grupos_caratula()

        footer_row = row_start + max_rows

        # --- Configuración del Logo ---
        logo_path_guardado = ""
        if os.path.exists(CONFIG_FILE):
            try:
                with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                    cfg = json.load(f)
                    logo_path_guardado = cfg.get("logo_escuela_path", "")
            except: pass

        self.var_logo_path = ctk.StringVar(value=logo_path_guardado)

        ctk.CTkLabel(f2, text="Logo de Escuela (Word):", anchor="w", font=("Segoe UI", 12, "bold")).grid(row=footer_row, column=0, sticky="w", padx=10, pady=5)
        entry_logo = ctk.CTkEntry(f2, textvariable=self.var_logo_path, state="readonly")
        entry_logo.grid(row=footer_row, column=1, sticky="ew", padx=20, pady=5)
        btn_logo = ctk.CTkButton(f2, text="Seleccionar Logo", command=self.seleccionar_logo)
        btn_logo.grid(row=footer_row, column=2, sticky="w", padx=10, pady=5)

        footer_row += 1
        # --- Fin Logo ---

        ctk.CTkLabel(

            f2,
            text="Al sincronizar, el programa inyectara estos datos en Portadas, Horarios y Asistencias.",
            text_color="#F59E0B",
            font=("Segoe UI", 11),
        ).grid(row=footer_row, column=0, columnspan=4, pady=15)

        self.btn_sinc = ctk.CTkButton(
            f2,
            text="SINCRONIZAR Y SOBREESCRIBIR EXCEL",
            fg_color="#3B82F6",
            height=45,
            font=("Segoe UI", 14, "bold"),
            command=self.sincronizar_plantilla,
        )
        self.btn_sinc.grid(row=footer_row + 1, column=0, columnspan=4, pady=10)


    def seleccionar_logo(self):
        ruta = filedialog.askopenfilename(
            title="Seleccionar Logo de Escuela",
            filetypes=(("Imágenes", "*.png *.jpg *.jpeg"), ("Todos", "*.*"))
        )
        if ruta:
            self.var_logo_path.set(ruta)
            self._guardar_logo_path(ruta)

    def _guardar_logo_path(self, ruta):
        datos = {}
        if os.path.exists(CONFIG_FILE):
            with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                try: datos = json.load(f)
                except: pass
        datos["logo_escuela_path"] = ruta
        with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
            json.dump(datos, f, ensure_ascii=False, indent=4)

    def _crear_campo(self, parent, row, texto, valor):
        ctk.CTkLabel(parent, text=texto, anchor="w", font=("Segoe UI", 12, "bold")).grid(row=row, column=0, sticky="w", padx=10, pady=5)
        entry = ctk.CTkEntry(parent)
        entry.insert(0, valor)
        entry.grid(row=row, column=1, sticky="ew", padx=20, pady=5)
        return entry

    # ================= PESTAÑA 2: HORARIOS (GRID RESPONSIVO Y CENTRADO) =================
    def crear_panel_horarios(self):
        horario_data = self.engine.obtener_horario()
        
        # Calculadora
        f_calc = ctk.CTkFrame(self.tab_hor, fg_color="#1E2D42", corner_radius=10)
        f_calc.pack(fill="x", padx=10, pady=(10,0))
        ctk.CTkLabel(f_calc, text="⏱️ Calculadora de Tiempos", font=("Segoe UI", 16, "bold"), text_color="#10B981").grid(row=0, column=0, columnspan=4, sticky="w", padx=10, pady=(5,5))
        
        ctk.CTkLabel(f_calc, text="Hora Entrada:").grid(row=1, column=0, sticky="e", padx=5)
        self.calc_inicio = ctk.CTkEntry(f_calc, width=80); self.calc_inicio.insert(0, "07:00"); self.calc_inicio.grid(row=1, column=1, sticky="w", padx=5)
        ctk.CTkLabel(f_calc, text="Hora Salida:").grid(row=1, column=2, sticky="e", padx=5)
        self.calc_salida = ctk.CTkEntry(f_calc, width=80); self.calc_salida.insert(0, "12:15"); self.calc_salida.grid(row=1, column=3, sticky="w", padx=5)
        
        ctk.CTkLabel(f_calc, text="Mins Receso:").grid(row=2, column=0, sticky="e", padx=5, pady=10)
        self.calc_receso = ctk.CTkEntry(f_calc, width=60); self.calc_receso.insert(0, "20"); self.calc_receso.grid(row=2, column=1, sticky="w", padx=5, pady=10)
        ctk.CTkLabel(f_calc, text="Después del per.:").grid(row=2, column=2, sticky="e", padx=5, pady=10)
        self.calc_per_receso = ctk.CTkOptionMenu(f_calc, values=["1", "2", "3", "4", "5"], width=60); self.calc_per_receso.set("4"); self.calc_per_receso.grid(row=2, column=3, sticky="w", padx=5, pady=10)
        
        ctk.CTkButton(f_calc, text="⚡ Generar Horas", fg_color="#3B82F6", command=self.calcular_horas_automaticas).grid(row=1, column=4, rowspan=2, padx=20, sticky="ew")

        # Tabla de Horario
        f_table = ctk.CTkFrame(self.tab_hor, fg_color="#1A2638", corner_radius=10)
        f_table.pack(fill="both", expand=True, padx=10, pady=10)
        
        f_table.grid_columnconfigure((2,3,4,5,6), weight=1)
        f_table.grid_columnconfigure(1, weight=0, minsize=100) 
        f_table.grid_columnconfigure(0, weight=0, minsize=40)  
        
        headers = ["Per.", "Horas", "Lunes", "Martes", "Miércoles", "Jueves", "Viernes"]
        for col, text in enumerate(headers):
            ctk.CTkLabel(f_table, text=text, font=("Segoe UI", 12, "bold"), fg_color="#253650", corner_radius=5).grid(row=0, column=col, sticky="ew", padx=2, pady=5)

        self.entradas_horario = []
        self.receso_hora_var = ctk.StringVar(value="-- : --")
        
        current_row = 1

        for i, fila in enumerate(horario_data):
            if i == 4:
                # =========================================================================================
                # 🛠️ NUEVO DISEÑO DE RECESO POR CASILLAS INDEPENDIENTES Y ALINEADAS
                # =========================================================================================
                
                # Definimos una altura uniforme fija (height=30) para igualar a las casillas blancas normales
                altura_receso = 30
                
                # Reducimos los márgenes verticales (pady=1) para un look más fino
                margen_v = 1 

                # 🌟 1. CASILLA PERIODO (Columna 0): Estrella independiente sin fondo azul.
                # Se coloca directamente en f_table, usando el fondo gris predeterminado de la tabla.
                lbl_star = ctk.CTkLabel(f_table, text="★", font=("Segoe UI", 12, "bold")) 
                lbl_star.grid(row=current_row, column=0, sticky="nsew", pady=margen_v)

                # ⏱️ 2. CASILLA HORAS (Columna 1): Cuadro de hora independiente y enmarcado.
                # Usamos un Frame gris claro independiente para que parezca una "casilla" separada.
                f_rec_time = ctk.CTkFrame(f_table, fg_color="#2D75AD", corner_radius=5, height=altura_receso) 
                f_rec_time.grid(row=current_row, column=1, sticky="nsew", padx=2, pady=margen_v)
                
                # Evita que el contenido estire el frame
                f_rec_time.grid_propagate(False) 

                # Label interno CENTRADO ABSOLUTAMENTE en su propia casilla gris (como pediste)
                # sticky="w" (oeste/izquierda) con padx=20 para separarlo del borde interno.
                ctk.CTkLabel(f_rec_time, textvariable=self.receso_hora_var, font=("Segoe UI", 12, "bold"), text_color="#E7E9EB").place(relx=0.5, rely=0.5, anchor="center")


                # 📖 3. CASILLA TEXTO (Columnas 2-6): Bloque azul continuo, grande y centrado.
                # Abarca desde la columna 2 hasta la 6 (columnspan=5), creando la franja azul que dibujaste.
                f_rec_text = ctk.CTkFrame(f_table, fg_color="#2563EB", corner_radius=5, height=altura_receso) 
                f_rec_text.grid(row=current_row, column=2, columnspan=5, sticky="nsew", padx=2, pady=margen_v)
                
                # Evita que el contenido estire el frame
                f_rec_text.grid_propagate(False) 

                # Label central abarcando de la columna 2 a la 6 (Días de la semana).
                # Texto cambiado de "R E C E S O   A C A D É M I C O" a "RECESO ESCOLAR" como pediste.
                # Font Bold Prominente (como indica tu dibujo rojo grande)
                lbl_receso = ctk.CTkLabel(f_rec_text, text="RECESO ESCOLAR", font=("Segoe UI", 12, "bold"), text_color="white")
                
                # Centrado horizontal y verticalmente en el CENTRO ABSOLUTO de su propia casilla azul (como pediste)
                lbl_receso.place(relx=0.5, rely=0.5, anchor="center")
                
                # =========================================================================================
                
                current_row += 1

            # ====== ESTO ERA LO QUE FALTABA: EL CÓDIGO QUE DIBUJA LAS CASILLAS ======
            ctk.CTkLabel(f_table, text=f"{i+1}").grid(row=current_row, column=0, padx=2, pady=2)
            
            ent_horas = ctk.CTkEntry(f_table, justify="center"); ent_horas.insert(0, fila.get("horas", ""))
            ent_horas.grid(row=current_row, column=1, sticky="ew", padx=2, pady=2)
            
            ent_lun = ctk.CTkEntry(f_table, justify="center"); ent_lun.insert(0, fila.get("lunes", ""))
            ent_lun.grid(row=current_row, column=2, sticky="ew", padx=2, pady=2)
            
            ent_mar = ctk.CTkEntry(f_table, justify="center"); ent_mar.insert(0, fila.get("martes", ""))
            ent_mar.grid(row=current_row, column=3, sticky="ew", padx=2, pady=2)
            
            ent_mie = ctk.CTkEntry(f_table, justify="center"); ent_mie.insert(0, fila.get("miercoles", ""))
            ent_mie.grid(row=current_row, column=4, sticky="ew", padx=2, pady=2)
            
            ent_jue = ctk.CTkEntry(f_table, justify="center"); ent_jue.insert(0, fila.get("jueves", ""))
            ent_jue.grid(row=current_row, column=5, sticky="ew", padx=2, pady=2)
            
            ent_vie = ctk.CTkEntry(f_table, justify="center"); ent_vie.insert(0, fila.get("viernes", ""))
            ent_vie.grid(row=current_row, column=6, sticky="ew", padx=2, pady=2)
            
            self.entradas_horario.append({"horas": ent_horas, "lunes": ent_lun, "martes": ent_mar, "miercoles": ent_mie, "jueves": ent_jue, "viernes": ent_vie})
            current_row += 1

        # Botón de guardar seguro y expandible
        self.btn_guardar_h = ctk.CTkButton(self.tab_hor, text="💾 GUARDAR HORARIO EN EXCEL", fg_color="#F59E0B", hover_color="#D97706", height=45, font=("Segoe UI", 14, "bold"), command=self.guardar_horario)
        self.btn_guardar_h.pack(fill="x", padx=10, pady=10)

    # ================= FUNCIONES COMPARTIDAS =================
    def calcular_horas_automaticas(self):
        try:
            inicio_str = self.calc_inicio.get().strip()
            salida_str = self.calc_salida.get().strip()
            mins_receso = int(self.calc_receso.get().strip())
            periodo_receso = int(self.calc_per_receso.get())
            
            h_in, m_in = map(int, inicio_str.split(":"))
            h_out, m_out = map(int, salida_str.split(":"))
            t_in = datetime.datetime(2000, 1, 1, h_in, m_in)
            t_out = datetime.datetime(2000, 1, 1, h_out, m_out)
            
            total_mins = (t_out - t_in).total_seconds() / 60
            mins_clase = int((total_mins - mins_receso) / 8) 
            
            hora_actual = t_in
            for i, campos in enumerate(self.entradas_horario):
                if i == periodo_receso: 
                    hora_fin_receso = hora_actual + datetime.timedelta(minutes=mins_receso)
                    self.receso_hora_var.set(f"{hora_actual.strftime('%I:%M').lstrip('0')} - {hora_fin_receso.strftime('%I:%M').lstrip('0')}")
                    hora_actual = hora_fin_receso
                
                hora_fin = hora_actual + datetime.timedelta(minutes=mins_clase)
                campos["horas"].delete(0, 'end')
                campos["horas"].insert(0, f"{hora_actual.strftime('%I:%M').lstrip('0')} - {hora_fin.strftime('%I:%M').lstrip('0')}")
                hora_actual = hora_fin
        except Exception as e:
            messagebox.showerror("Error", "Use formato HH:MM (Ej: 07:00) para las horas.")

    def guardar_horario(self):
        datos_guardar = []
        for cols in self.entradas_horario:
            datos_guardar.append({
                "horas": cols["horas"].get().strip(), "lunes": cols["lunes"].get().strip(),
                "martes": cols["martes"].get().strip(), "miercoles": cols["miercoles"].get().strip(),
                "jueves": cols["jueves"].get().strip(), "viernes": cols["viernes"].get().strip()
            })
        self.btn_guardar_h.configure(text="Guardando...", state="disabled"); self.update()
        if self.engine.guardar_horario(datos_guardar): messagebox.showinfo("Éxito", "El Horario se actualizó.")
        else: messagebox.showerror("Error", "No se encontró la hoja Horario.")
        self.btn_guardar_h.configure(text="💾 GUARDAR HORARIO EN EXCEL", state="normal")

    def sincronizar_plantilla(self):
        cargo_unificado = self.entry_con.get().strip() or "Instructor Vocacional"
        grupos_auto = ", ".join(self.engine.obtener_grados_activos())
        datos = {
            "docente_nombre": self.entry_doc.get().strip(), "docente_cedula": self.entry_ced.get().strip(),
            "seguro_social": self.entry_ss.get().strip(), "numero_posicion": self.entry_pos.get().strip(),
            "condicion_nombramiento": cargo_unificado, "escuela_nombre": self.entry_esc.get().strip(),
            "escuela_region": self.entry_reg.get().strip(), "distrito": self.entry_dis.get().strip(),
            "corregimiento": self.entry_correg.get().strip(), "zona_escolar": self.entry_zon.get().strip(),
            "director_nombre": self.entry_dir.get().strip(), "subdirector_nombre": self.entry_sub.get().strip(),
            "coordinador_nombre": self.entry_coo.get().strip(), "telefono": self.entry_tel.get().strip(),
            "correo": self.entry_mail.get().strip(), "ano_lectivo": self.entry_ano.get().strip(),
            "jornada": self.entry_jor.get().strip(), "fecha_t1": self.entry_t1.get().strip(),
            "fecha_t2": self.entry_t2.get().strip(), "fecha_t3": self.entry_t3.get().strip(),
            "titulo_caratula": cargo_unificado,
            "grupos_caratula": grupos_auto,
            "materias_activas": [m for m, var in getattr(self, "vars_materias", {}).items() if var.get()] if getattr(self, 'vars_materias', None) else [],
        }
        with open(CONFIG_FILE, 'w', encoding='utf-8') as f: json.dump(datos, f, ensure_ascii=False, indent=4)
        self.btn_sinc.configure(text="⏳ Sincronizando...", state="disabled"); self.update()
        def tarea():
            self.engine.sincronizar_plantilla_maestra(datos)
            self.after(0, lambda: self.finalizar_sinc())
        threading.Thread(target=tarea, daemon=True).start()

    def finalizar_sinc(self):
        self.btn_sinc.configure(text="✨ SINCRONIZAR Y SOBREESCRIBIR EXCEL", state="normal")
        messagebox.showinfo("Éxito", "Libreta actualizada con éxito.")

    def cambiar_modalidad(self):
        nueva = self.var_modalidad.get().lower()
        if nueva == self.engine.modalidad: return

        # Advertencia de cambios sin guardar
        msg = ("Advertencia: Si tiene cambios sin guardar (como modificar un horario o datos), "
               "se perderán al cambiar de modalidad.\n\n"
               f"¿Desea descartar los cambios y cambiar a modo {nueva.capitalize()}?")

        if not messagebox.askyesno("Cambio de Modalidad", msg):
            # Restaurar el valor del radio button si cancelan
            self.var_modalidad.set(self.engine.modalidad.capitalize())
            return

        archivo_nuevo = "Registro_Primaria.xlsx" if nueva == "primaria" else "Registro_2026.xlsx"
        ruta_nueva = os.path.join(BASE_DIR, "..", archivo_nuevo)
        if not os.path.exists(ruta_nueva):
            self.var_modalidad.set(self.engine.modalidad.capitalize())
            return messagebox.showerror("Error", f"Falta el archivo: {archivo_nuevo}")

        # Reiniciar motor y reconstruir UI
        self.app_principal.reiniciar_motor(ruta_nueva, nueva)

        # Limpiar/resetear la vista de configuración actual re-creándola,
        # Ya que al reiniciar el motor la app principal recarga la interfaz,
        # aseguremos que los campos internos que guardan referencias se borren si esta vista sigue viva (aunque al llamar mostrar_dashboard en teoria se cierra la vista).
        # En RegistroDoc, la navegacion suele destruir el main_content_frame, pero si no,
        # app_principal.mostrar_dashboard() recarga el Dashboard. Por lo que esto deberia ser suficiente para dejar el entorno limpio.

    # ================= PESTAÑA 3: GRADOS =================
    def crear_panel_grados(self):
        f1 = ctk.CTkFrame(self.tab_gra, fg_color="#1E2D42", corner_radius=10)
        f1.pack(fill="x", padx=20, pady=10, ipadx=10, ipady=10)

        ctk.CTkLabel(f1, text="➕ Agregar Nuevo Grado o Grupo (Ej. 10° o 8°B)", font=("Segoe UI", 16, "bold"), text_color="#10B981").pack(anchor="w", padx=20, pady=10)
        row1 = ctk.CTkFrame(f1, fg_color="transparent"); row1.pack(fill="x", padx=20, pady=5)
        self.entry_nuevo_grado = ctk.CTkEntry(row1, placeholder_text="Nombre del grado", width=150); self.entry_nuevo_grado.pack(side="left", padx=5)
        self.entry_cons_grado = ctk.CTkEntry(row1, placeholder_text="Prof. Consejero", width=200); self.entry_cons_grado.pack(side="left", padx=5)
        self.btn_crear_grado = ctk.CTkButton(row1, text="Crear Grado", command=self.agregar_grado); self.btn_crear_grado.pack(side="left", padx=20)

        f2 = ctk.CTkFrame(self.tab_gra, fg_color="#1A2638", corner_radius=10)
        f2.pack(fill="both", expand=True, padx=20, pady=10, ipadx=10, ipady=10)
        
        ctk.CTkLabel(f2, text="🗑️ Eliminar Grado Completo", font=("Segoe UI", 16, "bold"), text_color="#EF4444").pack(anchor="w", padx=20, pady=10)
        row2 = ctk.CTkFrame(f2, fg_color="transparent"); row2.pack(fill="x", padx=20, pady=15)
        self.combo_grado_del = ctk.CTkOptionMenu(row2, values=["Cargando..."]); self.combo_grado_del.pack(side="left", padx=5)
        self.btn_del_grado = ctk.CTkButton(row2, text="Eliminar Grado", fg_color="#EF4444", hover_color="#B91C1C", command=self.eliminar_grado); self.btn_del_grado.pack(side="left", padx=20)

        ctk.CTkLabel(f2, text="──────────────────────────", text_color="#475569").pack(pady=10)
        ctk.CTkLabel(f2, text="🔄 Actualizar Profesor Consejero", font=("Segoe UI", 16, "bold"), text_color="#3B82F6").pack(anchor="w", padx=20, pady=5)
        row3 = ctk.CTkFrame(f2, fg_color="transparent"); row3.pack(fill="x", padx=20, pady=10)
        self.combo_grado_cons = ctk.CTkOptionMenu(row3, values=["Cargando..."], command=self.mostrar_consejero_actual); self.combo_grado_cons.pack(side="left", padx=5)
        self.entry_nuevo_cons = ctk.CTkEntry(row3, placeholder_text="Nuevo Nombre del Consejero", width=250); self.entry_nuevo_cons.pack(side="left", padx=15)
        self.btn_act_cons = ctk.CTkButton(row3, text="Actualizar Consejero", command=self.actualizar_consejero); self.btn_act_cons.pack(side="left", padx=10)

    def mostrar_consejero_actual(self, grado):
        consejero = self.engine.obtener_consejero_actual(grado)
        self.entry_nuevo_cons.delete(0, 'end')
        if consejero != "No asignado": self.entry_nuevo_cons.insert(0, consejero)

    def agregar_grado(self):
        g = self.entry_nuevo_grado.get().strip(); c = self.entry_cons_grado.get().strip()
        if not g: return messagebox.showwarning("Atención", "Escribe el nombre del grado.")
        self.btn_crear_grado.configure(text="Creando...", state="disabled"); self.update()
        exito, msj = self.engine.agregar_grado(g, c, "Matutina")
        self.btn_crear_grado.configure(text="Crear Grado", state="normal")
        if exito:
            messagebox.showinfo("Éxito", "Grado creado."); self.entry_nuevo_grado.delete(0, 'end'); self.entry_cons_grado.delete(0, 'end')
            self.actualizar_listas_ui(); self.app_principal.engine = self.engine 
        else: messagebox.showerror("Error", msj)

    def actualizar_consejero(self):
        g = self.combo_grado_cons.get(); c = self.entry_nuevo_cons.get().strip()
        if not g or not c: return messagebox.showwarning("Atención", "Escriba el nombre del nuevo consejero.")
        self.btn_act_cons.configure(text="Actualizando...", state="disabled"); self.update()
        if self.engine.actualizar_consejero(g, c): messagebox.showinfo("Éxito", "Consejero actualizado.")
        else: messagebox.showerror("Error", "No se pudo actualizar.")
        self.btn_act_cons.configure(text="Actualizar Consejero", state="normal")

    def eliminar_grado(self):
        g = self.combo_grado_del.get()
        if not g or g == "Cargando...": return
        if messagebox.askyesno("Peligro", f"¿Eliminar DEFINITIVAMENTE {g}?"):
            self.btn_del_grado.configure(text="Borrando...", state="disabled"); self.update()
            self.engine.eliminar_grado(g); self.btn_del_grado.configure(text="Eliminar Grado", state="normal")
            messagebox.showinfo("Éxito", "Grado eliminado."); self.actualizar_listas_ui()

    # ================= PESTAÑA 4: MATERIAS =================
    def crear_panel_materias(self):
        f1 = ctk.CTkFrame(self.tab_mat, fg_color="#1E2D42", corner_radius=10)
        f1.pack(fill="x", padx=20, pady=10, ipadx=10, ipady=10)

        ctk.CTkLabel(f1, text="➕ Agregar Materia a un Grado", font=("Segoe UI", 16, "bold"), text_color="#10B981").pack(anchor="w", padx=20, pady=10)
        row1 = ctk.CTkFrame(f1, fg_color="transparent"); row1.pack(fill="x", padx=20, pady=5)
        ctk.CTkLabel(row1, text="Grado:").pack(side="left")
        self.combo_grado_mat = ctk.CTkOptionMenu(row1, values=["Cargando..."], command=self.cargar_materias_actuales, width=80); self.combo_grado_mat.pack(side="left", padx=10)
        ctk.CTkLabel(row1, text="Clonar formato de:").pack(side="left", padx=(10,0))
        self.combo_mat_base = ctk.CTkOptionMenu(row1, values=["Cargando..."], width=150); self.combo_mat_base.pack(side="left", padx=10)

        row2 = ctk.CTkFrame(f1, fg_color="transparent"); row2.pack(fill="x", padx=20, pady=10)
        self.entry_nueva_mat = ctk.CTkEntry(row2, placeholder_text="Nombre de Nueva Materia", width=200); self.entry_nueva_mat.pack(side="left", padx=5)
        self.combo_jornada_mat = ctk.CTkOptionMenu(row2, values=["Matutina", "Vespertina", "Nocturna"], width=120); self.combo_jornada_mat.pack(side="left", padx=5)
        self.btn_crear_mat = ctk.CTkButton(row2, text="Crear Materia", command=self.clonar_materia); self.btn_crear_mat.pack(side="left", padx=20)

        f2 = ctk.CTkFrame(self.tab_mat, fg_color="#1A2638", corner_radius=10)
        f2.pack(fill="both", expand=True, padx=20, pady=10, ipadx=10, ipady=10)

        ctk.CTkLabel(f2, text="🗑️ Eliminar Materia", font=("Segoe UI", 16, "bold"), text_color="#EF4444").pack(anchor="w", padx=20, pady=10)
        row3 = ctk.CTkFrame(f2, fg_color="transparent"); row3.pack(fill="x", padx=20, pady=5)
        self.combo_mat_del = ctk.CTkOptionMenu(row3, values=["Seleccione arriba"]); self.combo_mat_del.pack(side="left", padx=5)
        self.btn_del_mat = ctk.CTkButton(row3, text="Eliminar Materia", fg_color="#EF4444", hover_color="#B91C1C", command=self.eliminar_materia); self.btn_del_mat.pack(side="left", padx=20)

    def actualizar_grupos_caratula(self):
        if not hasattr(self, "var_grupos_caratula"):
            return
        grupos = ", ".join(self.engine.obtener_grados_activos())
        self.var_grupos_caratula.set(grupos)

    def actualizar_listas_ui(self):
        grados = self.engine.obtener_grados_activos()
        self.combo_grado_del.configure(values=grados); self.combo_grado_del.set(grados[0] if grados else "")
        self.combo_grado_cons.configure(values=grados)
        if grados: self.combo_grado_cons.set(grados[0]); self.mostrar_consejero_actual(grados[0]) 
        self.combo_grado_mat.configure(values=grados); self.combo_grado_mat.set(grados[0] if grados else "")
        self.cargar_materias_actuales(grados[0] if grados else "")
        self.actualizar_grupos_caratula()

    def cargar_materias_actuales(self, grado):
        if not grado: return
        materias = self.engine.obtener_materias_por_grado(grado)
        if materias and materias[0] != "Sin materias registradas":
            self.combo_mat_base.configure(values=materias); self.combo_mat_base.set(materias[0])
            self.combo_mat_del.configure(values=materias); self.combo_mat_del.set(materias[0])
        else:
            self.combo_mat_base.configure(values=["Ninguna"]); self.combo_mat_base.set("Ninguna")
            self.combo_mat_del.configure(values=["Ninguna"]); self.combo_mat_del.set("Ninguna")

    def clonar_materia(self):
        g = self.combo_grado_mat.get(); base = self.combo_mat_base.get(); n_mat = self.entry_nueva_mat.get().strip(); jor = self.combo_jornada_mat.get()
        if base == "Ninguna" or not n_mat: return messagebox.showwarning("Atención", "Escriba el nombre de la nueva materia.")
        self.btn_crear_mat.configure(text="Clonando...", state="disabled"); self.update()
        exito, msj = self.engine.clonar_materia(g, base, n_mat, jor)
        self.btn_crear_mat.configure(text="Crear Materia", state="normal")
        if exito: messagebox.showinfo("Éxito", msj); self.entry_nueva_mat.delete(0, 'end'); self.cargar_materias_actuales(g)
        else: messagebox.showerror("Error", msj)

    def eliminar_materia(self):
        g = self.combo_grado_mat.get(); m = self.combo_mat_del.get()
        if m == "Ninguna": return
        if messagebox.askyesno("Confirmar", f"¿Eliminar {m} de {g}?"):
            self.engine.eliminar_materia(g, m); messagebox.showinfo("Éxito", "Materia eliminada."); self.cargar_materias_actuales(g)
