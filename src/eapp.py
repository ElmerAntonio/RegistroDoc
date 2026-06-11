import customtkinter as ctk
from tkinter import messagebox
import datetime
import threading
from rdsecurity import validar_nota_meduca


class NotasFrame(ctk.CTkFrame):
    def __init__(self, master, engine, **kwargs):
        super().__init__(master, fg_color="transparent", **kwargs)
        self.engine = engine
        self.entradas_notas = {}
        self.pills_notas = {}
        self.estudiantes_notas = {}
        self.col_a_modificar = None

        self.grid_columnconfigure(0, weight=8)
        self.grid_columnconfigure(1, weight=2)
        self.grid_rowconfigure(0, weight=1)

        self.crear_panel_izquierdo()
        self.crear_panel_derecho()

        self.al_cambiar_grado(self.combo_grado.get())

    def crear_panel_izquierdo(self):
        frame_izq = ctk.CTkFrame(self, fg_color="#1A2638", corner_radius=10)
        frame_izq.grid(row=0, column=0, sticky="nsew", padx=(0, 10))

        top = ctk.CTkFrame(frame_izq, fg_color="transparent")
        top.pack(fill="x", padx=20, pady=15)

        lbl_texto = "Grado/Grupo:"
        opciones = self.engine.obtener_grados_activos()
        if not opciones:
            opciones = ["7°", "8°", "9°"] if self.engine.modalidad == "premedia" else ["1°"]

        ctk.CTkLabel(top, text=lbl_texto, font=(
            "Segoe UI", 16, "bold")).pack(side="left")

        self.combo_grado = ctk.CTkOptionMenu(
            top, values=opciones, command=self.al_cambiar_grado)
        self.combo_grado.pack(side="left", padx=10)

        # Switch/Checkbox for points scale next to the grade selection
        self.var_usar_puntos = ctk.BooleanVar(value=False)
        self.switch_puntos = ctk.CTkSwitch(
            top, text="Escala Puntos",
            variable=self.var_usar_puntos,
            command=self.al_cambiar_modo_puntos,
            font=("Segoe UI", 11, "bold")
        )
        self.switch_puntos.pack(side="left", padx=(15, 8))

        self.lbl_pts_max = ctk.CTkLabel(top, text="Máx:", font=("Segoe UI", 11, "bold"))
        self.lbl_pts_max.pack(side="left", padx=(5, 2))
        self.entry_pts_max = ctk.CTkEntry(top, width=45, height=24, justify="center", font=("Segoe UI", 11))
        self.entry_pts_max.insert(0, "40")
        self.entry_pts_max.pack(side="left", padx=2)
        self.entry_pts_max.bind("<KeyRelease>", lambda e: self.recalcular_escala_visual())
        self.entry_pts_max.configure(state="disabled")

        self.lbl_exigencia = ctk.CTkLabel(top, text="Exig:", font=("Segoe UI", 11, "bold"))
        self.lbl_exigencia.pack(side="left", padx=(5, 2))
        self.combo_exigencia = ctk.CTkOptionMenu(
            top, values=["50%", "60%", "70%"], width=75, height=24,
            command=lambda _: self.recalcular_escala_visual(),
            font=("Segoe UI", 11)
        )
        self.combo_exigencia.set("60%")
        self.combo_exigencia.pack(side="left", padx=2)
        self.combo_exigencia.configure(state="disabled")

        self.btn_ver_tabla_puntos = ctk.CTkButton(
            top, text="📋 Tabla", width=65, height=24,
            fg_color="#8B5CF6", hover_color="#7C3AED",
            font=("Segoe UI", 11, "bold"),
            command=self.abrir_tabla_puntos_modal
        )
        self.btn_ver_tabla_puntos.pack(side="left", padx=(8, 2))
        self.btn_ver_tabla_puntos.configure(state="disabled")

        self.scroll_estudiantes = ctk.CTkScrollableFrame(
            frame_izq, fg_color="transparent")
        self.scroll_estudiantes.pack(fill="both", expand=True,
                                     padx=15, pady=10)

    def crear_panel_derecho(self):
        frame_der = ctk.CTkFrame(self, fg_color="#1E2D42", corner_radius=10)
        frame_der.grid(row=0, column=1, sticky="nsew")

        ctk.CTkLabel(frame_der, text="Gestión de Notas", font=(
            "Segoe UI", 18, "bold"), text_color="#3B82F6").pack(pady=(20, 10))

        ctk.CTkLabel(frame_der, text="Materia a calificar:",
                     font=("Segoe UI", 13)).pack(anchor="w", padx=20)
        self.combo_materia = ctk.CTkOptionMenu(
            frame_der, values=["Cargando..."],
            command=self.cargar_descripciones)
        self.combo_materia.pack(fill="x", padx=20, pady=5)

        self.tabs = ctk.CTkTabview(frame_der, command=self.al_cambiar_pestana)
        self.tabs.pack(fill="both", expand=True, padx=10, pady=10)

        tab_nueva = self.tabs.add("Nueva Nota")
        tab_mod = self.tabs.add("Modificar")
        tab_tareas = self.tabs.add("Tareas")

        self.scroll_tareas_tab = ctk.CTkScrollableFrame(tab_tareas, fg_color="transparent")
        self.scroll_tareas_tab.pack(fill="both", expand=True, padx=5, pady=5)

        self.btn_nueva_t = ctk.CTkButton(
            tab_tareas,
            text="➕ Programar Tarea",
            fg_color="#10B981",
            hover_color="#059669",
            font=("Segoe UI", 12, "bold"),
            command=self.abrir_modal_programar_tarea
        )
        self.btn_nueva_t.pack(fill="x", padx=10, pady=(5, 10))

        # ====== TAB: NUEVA NOTA ======
        self.combo_trimestre = ctk.CTkOptionMenu(
            tab_nueva, values=["Trimestre 1", "Trimestre 2", "Trimestre 3"])
        self.combo_trimestre.pack(fill="x", padx=10, pady=10)
        try:
            from utils.date_helpers import obtener_trimestre_actual
            self.combo_trimestre.set(obtener_trimestre_actual())
            self.combo_trimestre.configure(state="disabled")
        except Exception:
            pass

        self.combo_tipo = ctk.CTkOptionMenu(
            tab_nueva, values=["Diaria / Parcial", "Apreciación", "Examen"],
            command=self.al_cambiar_tipo_nota)
        self.combo_tipo.pack(fill="x", padx=10, pady=10)

        ctk.CTkLabel(tab_nueva, text="Fecha (MM-DD):", font=("Segoe UI", 12)
                     ).pack(anchor="w", padx=10, pady=(5, 0))
        self.entry_fecha = ctk.CTkEntry(tab_nueva)
        self.entry_fecha.insert(0, datetime.datetime.now().strftime("%m-%d"))
        self.entry_fecha.pack(fill="x", padx=10, pady=5)

        self.lbl_desc_counter = ctk.CTkLabel(
            tab_nueva, text="Descripción (Ej. Charla): (0/25)",
            font=("Segoe UI", 12))
        self.lbl_desc_counter.pack(anchor="w", padx=10, pady=(5, 0))

        self.var_desc = ctk.StringVar()
        self.var_desc.trace_add("write", self.actualizar_contador_desc)
        self.entry_desc = ctk.CTkEntry(tab_nueva, textvariable=self.var_desc, font=("Calibri", 11), justify="center")
        self.entry_desc.pack(fill="x", padx=10, pady=5)

        # Relleno Rápido de Notas
        relleno_frame = ctk.CTkFrame(tab_nueva, fg_color="transparent")
        relleno_frame.pack(fill="x", padx=10, pady=(5, 5))
        ctk.CTkLabel(relleno_frame, text="Relleno rápido:", font=("Segoe UI", 11)).pack(side="left")
        self.entry_relleno = ctk.CTkEntry(relleno_frame, width=50, placeholder_text="5.0", justify="center")
        self.entry_relleno.insert(0, "5.0")
        self.entry_relleno.pack(side="left", padx=5)
        btn_rellenar = ctk.CTkButton(
            relleno_frame, text="Rellenar", width=70, height=26,
            fg_color="#3B82F6", hover_color="#2563EB",
            font=("Segoe UI", 11, "bold"),
            command=self.rellenar_notas_grupo
        )
        btn_rellenar.pack(side="left", padx=5)

        self.btn_guardar_nueva = ctk.CTkButton(
            tab_nueva,
            text="💾 GUARDAR NUEVA",
            fg_color="#10B981",
            hover_color="#059669",
            font=(
                "Segoe UI",
                14,
                "bold"),
            height=40,
            command=self.guardar_notas)
        self.btn_guardar_nueva.pack(pady=10, padx=10, fill="x")

        self.btn_programar_tarea = ctk.CTkButton(
            tab_nueva,
            text="📅 PROGRAMAR COMO TAREA",
            fg_color="#6366F1",
            hover_color="#4F46E5",
            font=("Segoe UI", 12, "bold"),
            height=32,
            command=self.abrir_modal_programar_tarea
        )
        self.btn_programar_tarea.pack(pady=(5, 10), padx=10, fill="x")

        # ====== TAB: MODIFICAR NOTA ======
        self.combo_trimestre_mod = ctk.CTkOptionMenu(
            tab_mod, values=["Trimestre 1", "Trimestre 2", "Trimestre 3"],
            command=self.cargar_descripciones)
        self.combo_trimestre_mod.pack(fill="x", padx=10, pady=5)

        self.combo_tipo_mod = ctk.CTkOptionMenu(
            tab_mod, values=["Diaria / Parcial", "Apreciación", "Examen"],
            command=self.cargar_descripciones)
        self.combo_tipo_mod.pack(fill="x", padx=10, pady=5)

        ctk.CTkLabel(tab_mod, text="Seleccione la nota a editar:", font=(
            "Segoe UI", 12, "bold")).pack(anchor="w", padx=10, pady=(15, 0))
        self.combo_desc_mod = ctk.CTkOptionMenu(
            tab_mod, values=["Seleccione parámetros arriba..."])
        self.combo_desc_mod.pack(fill="x", padx=10, pady=5)

        ctk.CTkButton(
            tab_mod,
            text="🔍 CARGAR A LA LISTA",
            fg_color="#F59E0B",
            hover_color="#D97706",
            font=(
                "Segoe UI",
                12,
                "bold"),
            command=self.buscar_modificar).pack(
            pady=10,
            padx=10,
            fill="x")

        self.btn_actualizar = ctk.CTkButton(
            tab_mod,
            text="🔄 ACTUALIZAR EXCEL",
            fg_color="#3B82F6",
            hover_color="#2563EB",
            font=(
                "Segoe UI",
                14,
                "bold"),
            height=40,
            command=self.actualizar_notas)
        self.btn_actualizar.pack(pady=5, padx=10, fill="x")

    def al_cambiar_pestana(self):
        self.limpiar_campos_notas()

    def limpiar_campos_notas(self):
        for entries_list in self.entradas_notas.values():
            for entry in entries_list:
                if entry.winfo_exists():
                    entry.delete(0, 'end')
        if hasattr(self, 'entry_desc') and self.entry_desc.winfo_exists():
            try:
                old_state = self.entry_desc.cget("state")
                self.entry_desc.configure(state="normal")
                self.entry_desc.delete(0, 'end')
                self.entry_desc.configure(state=old_state)
            except Exception:
                pass

    def rellenar_notas_grupo(self):
        val = self.entry_relleno.get().strip()
        if not val:
            val = "40" if self.var_usar_puntos.get() else "5.0"
        
        try:
            val_float = float(val.replace(",", "."))
            if self.var_usar_puntos.get():
                try:
                    pts_max = float(self.entry_pts_max.get())
                except Exception:
                    pts_max = 40.0
                if val_float < 0 or val_float > pts_max:
                    messagebox.showwarning("Atención", f"Los puntos deben estar entre 0 y {pts_max}.")
                    return
                val = f"{val_float:.0f}" if val_float.is_integer() else f"{val_float:.1f}"
            else:
                if val_float < 1.0 or val_float > 5.0:
                    messagebox.showwarning("Atención", "La nota debe estar entre 1.0 y 5.0.")
                    return
                val = f"{val_float:.1f}"
        except ValueError:
            msg = "Ingrese un valor numérico válido de puntos." if self.var_usar_puntos.get() else "Ingrese un valor numérico de nota válido (Ej: 5.0)."
            messagebox.showwarning("Atención", msg)
            return

        for eid, entries in self.entradas_notas.items():
            if len(entries) > 1:
                entry_pts = entries[1]
                entry_pts.delete(0, 'end')
                entry_pts.insert(0, val)
                self.al_cambiar_puntos_estudiante(eid)
            else:
                entry = entries[0]
                entry.delete(0, 'end')
                entry.insert(0, val)
                pass

    def actualizar_contador_desc(self, *args):
        texto = self.var_desc.get()
        if len(texto) > 25:
            # We must use self.entry_desc to prevent recursive trace issues or lag in UI update
            # sometimes set() causes cursor jump
            self.var_desc.set(texto[:25])
            # Move cursor to end
            self.entry_desc.icursor("end")
            texto = texto[:25]
        if hasattr(self, 'lbl_desc_counter'):
            self.lbl_desc_counter.configure(
                text=f"Descripción (Ej. Charla): ({len(texto)}/25)")

    def al_cambiar_tipo_nota(self, valor):
        self.entry_desc.configure(state="normal")
        self.entry_desc.delete(0, 'end')

        if valor == "Examen":
            self.entry_desc.insert(0, "Automático (Examen + Fecha)")
            self.entry_desc.configure(state="disabled")
        else:
            self.entry_desc.configure(
                placeholder_text="Ej. Charla, Taller, etc.")

    def al_cambiar_grado(self, grado):
        self.cargar_estudiantes(grado)
        materias = self.engine.obtener_materias_por_grado(grado)
        if materias:
            self.combo_materia.configure(values=materias)
            self.combo_materia.set(materias[0])
        else:
            self.combo_materia.configure(values=["No hay materias"])
            self.combo_materia.set("No hay materias")
        self.cargar_descripciones()

    def cargar_estudiantes(self, grado=None):
        if grado is None:
            grado = self.combo_grado.get()
        for w in self.scroll_estudiantes.winfo_children():
            w.destroy()
        self.entradas_notas.clear()
        self.pills_notas.clear()
        self.estudiantes_notas.clear()
        self.col_a_modificar = None

        # Header Row
        header_row = ctk.CTkFrame(self.scroll_estudiantes, fg_color="transparent")
        header_row.pack(fill="x", pady=(0, 10))
        ctk.CTkLabel(header_row, text="N°", width=30, font=("Segoe UI", 12, "bold")).pack(side="left")
        ctk.CTkLabel(header_row, text="Nombre", anchor="w", font=("Segoe UI", 12, "bold")).pack(side="left", fill="x", expand=True, padx=(10, 10))
        
        ctk.CTkLabel(header_row, text="Nota", width=80, anchor="center", font=("Segoe UI", 12, "bold")).pack(side="right", padx=10)
        if self.var_usar_puntos.get():
            ctk.CTkLabel(header_row, text="Puntos", width=60, anchor="center", font=("Segoe UI", 12, "bold")).pack(side="right", padx=10)

        ests = self.engine.obtener_estudiantes_completos(grado)
        for est in ests:
            row = ctk.CTkFrame(self.scroll_estudiantes, fg_color="transparent")
            row.pack(fill="x", pady=2)
            ctk.CTkLabel(row, text=f"{est['id']}.", width=30).pack(side="left")
            ctk.CTkLabel(row, text=est['nombre'], anchor="w").pack(side="left", fill="x", expand=True, padx=(10, 10))
            
            # Botón de informe individual (boletín)
            btn_inf = ctk.CTkButton(
                row, text="📄 Boletín", width=78, height=24,
                fg_color="transparent", hover_color="#2A3B50",
                text_color="#10B981", font=("Segoe UI", 11, "bold"),
                command=lambda e=est: self.generar_informe_estudiante(e)
            )
            btn_inf.pack(side="right", padx=5)

            entry = ctk.CTkEntry(row, width=80, height=25, justify="center", placeholder_text="-", font=("Segoe UI", 11))
            entry.pack(side="right", padx=10)
            
            if self.var_usar_puntos.get():
                entry_pts = ctk.CTkEntry(row, width=60, height=25, justify="center", placeholder_text="Pts", font=("Segoe UI", 11))
                entry_pts.pack(side="right", padx=10)
                
                # Evento key release para calcular la nota por puntos en tiempo real
                entry_pts.bind("<KeyRelease>", lambda e, eid=est['id']: self.al_cambiar_puntos_estudiante(eid))
                
                # Keyboard navigation
                entry_pts.bind("<Return>", lambda e, idx=est['id']: self.al_presionar_enter(idx))
                entry_pts.bind("<Down>", lambda e, idx=est['id']: self.al_presionar_abajo(idx))
                entry_pts.bind("<Up>", lambda e, idx=est['id']: self.al_presionar_arriba(idx))
                
                self.entradas_notas[est['id']] = [entry, entry_pts]
            else:
                # Keyboard navigation
                entry.bind("<Return>", lambda e, idx=est['id']: self.al_presionar_enter(idx))
                entry.bind("<Down>", lambda e, idx=est['id']: self.al_presionar_abajo(idx))
                entry.bind("<Up>", lambda e, idx=est['id']: self.al_presionar_arriba(idx))
                
                self.entradas_notas[est['id']] = [entry]
                
            self.pills_notas[est['id']] = None
            self.estudiantes_notas[est['id']] = est

    def cargar_descripciones(self, *args):
        grado = self.combo_grado.get()
        materia = self.combo_materia.get()
        trimestre = self.combo_trimestre_mod.get()
        tipo = self.combo_tipo_mod.get()

        descripciones = self.engine.obtener_descripciones_notas(
            grado, materia, trimestre, tipo)
        if descripciones:
            self.combo_desc_mod.configure(values=descripciones)
            self.combo_desc_mod.set(descripciones[0])
        else:
            vacio = ["Sin notas registradas"]
            self.combo_desc_mod.configure(values=vacio)
            self.combo_desc_mod.set(vacio[0])

        if hasattr(self, 'scroll_tareas_tab'):
            self.actualizar_tab_tareas()

    def actualizar_vista(self):
        self.al_cambiar_grado(self.combo_grado.get())

    def actualizar_tab_tareas(self):
        for w in self.scroll_tareas_tab.winfo_children():
            w.destroy()
            
        grado = self.combo_grado.get()
        materia = self.combo_materia.get()
        
        if not grado or not materia:
            return
            
        from tareas import cargar_tareas
        all_tareas = cargar_tareas()
        
        filtered_tareas = [
            t for t in all_tareas
            if t.get("grado") == grado and (t.get("materia") == materia or t.get("materia") == "General")
        ]
        
        if not filtered_tareas:
            ctk.CTkLabel(
                self.scroll_tareas_tab,
                text="No hay tareas programadas\npara esta materia.",
                font=("Segoe UI", 12),
                text_color="#94A3B8"
            ).pack(pady=30)
            return
            
        pendientes = [t for t in filtered_tareas if not t.get("completada")]
        completadas = [t for t in filtered_tareas if t.get("completada")]
        
        import datetime
        hoy = datetime.date.today()
        
        def calculate_urgencia(fecha_str):
            try:
                partes = fecha_str.split("-")
                fecha = datetime.date(int(partes[2]), int(partes[1]), int(partes[0]))
                dias = (fecha - hoy).days
                if dias < 0:
                    return "vencida", dias
                elif dias == 0:
                    return "hoy", 0
                elif dias <= 2:
                    return "urgente", dias
                else:
                    return "normal", dias
            except Exception:
                return "normal", 999
                
        if pendientes:
            ctk.CTkLabel(
                self.scroll_tareas_tab,
                text="⏳ PENDIENTES",
                font=("Segoe UI", 11, "bold"),
                text_color="#F59E0B"
            ).pack(anchor="w", padx=5, pady=(5, 5))
            
            for t in pendientes:
                urg, dias = calculate_urgencia(t["fecha_limite"])
                colores_urg = {
                    "vencida": "#EF4444", "hoy": "#F59E0B",
                    "urgente": "#FB923C", "normal": "#475569"
                }
                color_b = colores_urg.get(urg, "#475569")
                
                card = ctk.CTkFrame(
                    self.scroll_tareas_tab,
                    fg_color="#1E293B",
                    border_width=1,
                    border_color=color_b,
                    corner_radius=8
                )
                card.pack(fill="x", pady=4, padx=2)
                
                ctk.CTkLabel(
                    card,
                    text=t["titulo"],
                    font=("Segoe UI", 12, "bold"),
                    text_color="white",
                    anchor="w",
                    justify="left"
                ).pack(anchor="w", padx=10, pady=(6, 0))
                
                dias_txt = ""
                if urg == "vencida":
                    dias_txt = f" (Vencida hace {-dias}d)"
                elif urg == "hoy":
                    dias_txt = " (Hoy!)"
                elif urg == "urgente":
                    dias_txt = f" (En {dias}d)"
                
                ctk.CTkLabel(
                    card,
                    text=f"{t['tipo']} | Lim: {t['fecha_limite']}{dias_txt}",
                    font=("Segoe UI", 10),
                    text_color=color_b,
                    anchor="w"
                ).pack(anchor="w", padx=10, pady=(0, 4))
                
                btn_frame = ctk.CTkFrame(card, fg_color="transparent")
                btn_frame.pack(fill="x", padx=10, pady=(0, 6))
                
                def marcar_comp(tid=t["id"]):
                    from tareas import marcar_completada
                    marcar_completada(tid)
                    self.actualizar_tab_tareas()
                    root = self.winfo_toplevel()
                    if hasattr(root, "mostrar_toast"):
                        root.mostrar_toast("✓ Tarea completada")
                    if hasattr(root, "main_app") and hasattr(root.main_app, "_frames"):
                        dashboard = root.main_app._frames.get("DashboardFrame")
                        if dashboard and hasattr(dashboard, "actualizar_vista"):
                            dashboard.actualizar_vista()
                
                ctk.CTkButton(
                    btn_frame,
                    text="✓",
                    width=28,
                    height=24,
                    fg_color="#10B981",
                    hover_color="#059669",
                    command=marcar_comp
                ).pack(side="left", padx=2)
                
                def calificar_tarea(tit=t["titulo"], ttype=t["tipo"], date=t["fecha_limite"]):
                    self.tabs.set("Nueva Nota")
                    
                    mapped_type = "Diaria / Parcial"
                    if ttype in ["Apreciación", "Asistencia", "Hábitos"]:
                        mapped_type = "Apreciación"
                    elif ttype == "Examen":
                        mapped_type = "Examen"
                    
                    self.combo_tipo.set(mapped_type)
                    self.al_cambiar_tipo_nota(mapped_type)
                    
                    if mapped_type != "Examen":
                        self.entry_desc.configure(state="normal")
                        self.entry_desc.delete(0, "end")
                        self.entry_desc.insert(0, tit[:20])
                    
                    try:
                        partes = date.split("-")
                        fecha_mm_dd = f"{partes[1]}-{partes[0]}"
                        self.entry_fecha.delete(0, "end")
                        self.entry_fecha.insert(0, fecha_mm_dd)
                    except Exception:
                        pass
                        
                ctk.CTkButton(
                    btn_frame,
                    text="📝 Calificar",
                    width=75,
                    height=24,
                    fg_color="#3B82F6",
                    hover_color="#2563EB",
                    font=("Segoe UI", 10, "bold"),
                    command=calificar_tarea
                ).pack(side="left", padx=2)
                
                def del_t(tid=t["id"]):
                    if messagebox.askyesno("Confirmar", "¿Desea eliminar esta tarea programada?"):
                        from tareas import eliminar_tarea
                        eliminar_tarea(tid)
                        self.actualizar_tab_tareas()
                        root = self.winfo_toplevel()
                        if hasattr(root, "mostrar_toast"):
                            root.mostrar_toast("Tarea eliminada")
                        if hasattr(root, "main_app") and hasattr(root.main_app, "_frames"):
                            dashboard = root.main_app._frames.get("DashboardFrame")
                            if dashboard and hasattr(dashboard, "actualizar_vista"):
                                dashboard.actualizar_vista()
                                
                ctk.CTkButton(
                    btn_frame,
                    text="🗑️",
                    width=28,
                    height=24,
                    fg_color="#EF4444",
                    hover_color="#DC2626",
                    command=del_t
                ).pack(side="right", padx=2)
                
        if completadas:
            ctk.CTkLabel(
                self.scroll_tareas_tab,
                text="✅ COMPLETADAS",
                font=("Segoe UI", 11, "bold"),
                text_color="#10B981"
            ).pack(anchor="w", padx=5, pady=(12, 5))
            
            for t in completadas[-5:]:
                card = ctk.CTkFrame(
                    self.scroll_tareas_tab,
                    fg_color="#0F172A",
                    border_width=1,
                    border_color="#334155",
                    corner_radius=8
                )
                card.pack(fill="x", pady=4, padx=2)
                
                ctk.CTkLabel(
                    card,
                    text=t["titulo"],
                    font=("Segoe UI", 11, "bold", "overstrike"),
                    text_color="#64748B",
                    anchor="w",
                    justify="left"
                ).pack(anchor="w", padx=10, pady=(6, 0))
                
                ctk.CTkLabel(
                    card,
                    text=f"{t['tipo']} | Lim: {t['fecha_limite']}",
                    font=("Segoe UI", 9),
                    text_color="#475569",
                    anchor="w"
                ).pack(anchor="w", padx=10, pady=(0, 6))
                
                btn_frame = ctk.CTkFrame(card, fg_color="transparent")
                btn_frame.pack(fill="x", padx=10, pady=(0, 6))
                
                def del_t(tid=t["id"]):
                    if messagebox.askyesno("Confirmar", "¿Desea eliminar esta tarea programada?"):
                        from tareas import eliminar_tarea
                        eliminar_tarea(tid)
                        self.actualizar_tab_tareas()
                        root = self.winfo_toplevel()
                        if hasattr(root, "mostrar_toast"):
                            root.mostrar_toast("Tarea eliminada")
                        if hasattr(root, "main_app") and hasattr(root.main_app, "_frames"):
                            dashboard = root.main_app._frames.get("DashboardFrame")
                            if dashboard and hasattr(dashboard, "actualizar_vista"):
                                dashboard.actualizar_vista()
                                
                ctk.CTkButton(
                    btn_frame,
                    text="🗑️",
                    width=28,
                    height=24,
                    fg_color="#EF4444",
                    hover_color="#DC2626",
                    command=del_t
                ).pack(side="right", padx=2)

    def _recopilar_notas_validadas(self):
        notas_guardar = {}
        for id_est, entries_list in self.entradas_notas.items():
            val = entries_list[0].get().strip()
            if val:
                valido, nota, msg = validar_nota_meduca(val)
                if not valido:
                    messagebox.showerror("Error", f"Error en la nota para {id_est}: {msg}")
                    return None
                notas_guardar[id_est] = nota
        return notas_guardar

    # ==============================================================
    # NUEVA MAGIA: GUARDADO EN SEGUNDO PLANO (THREADING)
    # ==============================================================
    def guardar_notas(self):
        materia = self.combo_materia.get()
        grado = self.combo_grado.get()
        trimestre = self.combo_trimestre.get()
        tipo = self.combo_tipo.get()
        fecha = self.entry_fecha.get()
        desc = self.entry_desc.get().strip()

        if tipo != "Examen" and not desc:
            messagebox.showwarning("Atención", "Escribe una descripción.")
            return

        if tipo != "Examen" and len(desc) > 20:
            if not messagebox.askyesno(
                "Advertencia de Formato",
                f"La descripción '{desc}' tiene {len(desc)} caracteres (máx. recomendado: 20).\n"
                "Las descripciones largas podrían recortarse en la impresión de la planilla.\n\n"
                "¿Desea continuar de todos modos?"
            ):
                return

        notas_guardar = self._recopilar_notas_validadas()
        if notas_guardar is None or not notas_guardar:
            if notas_guardar == {}:
                messagebox.showinfo("Aviso", "No hay notas para guardar.")
            return

        # 1. Limpiamos la pantalla inmediatamente para fluidez
        for entries_list in self.entradas_notas.values():
            for entry in entries_list:
                entry.delete(0, 'end')
        if tipo != "Examen":
            self.entry_desc.delete(0, 'end')

        self.btn_guardar_nueva.configure(
            text="Procesando en fondo...", state="disabled")

        # 2. Creamos la tarea que trabajará "invisible"
        def tarea_fondo():
            exito, msj = self.engine.guardar_columna_notas(
                grado, materia, trimestre, tipo, fecha, desc, notas_guardar)
            # Volvemos a la pantalla principal para devolver
            # el botón a la normalidad
            self.after(0, lambda: self.finalizar_guardado(exito, msj))

        # 3. Soltamos al trabajador invisible
        threading.Thread(target=tarea_fondo, daemon=True).start()

    def finalizar_guardado(self, exito, msj):
        self.btn_guardar_nueva.configure(text="💾 GUARDAR NUEVA",
                                         state="normal")
        if exito:
            self.cargar_descripciones()
            root = self.winfo_toplevel()
            if hasattr(root, "mostrar_toast"):
                root.mostrar_toast("✓ Calificaciones guardadas en Excel", color="#10B981")
        else:
            messagebox.showerror("Error de guardado", msj)

    def buscar_modificar(self):
        grado = self.combo_grado.get()
        materia = self.combo_materia.get()
        trimestre = self.combo_trimestre_mod.get()
        tipo = self.combo_tipo_mod.get()
        desc = self.combo_desc_mod.get()

        if desc == "Sin notas registradas" or \
                desc == "Seleccione parámetros arriba...":
            messagebox.showwarning(
                "Atención", "No hay una nota válida seleccionada.")
            return

        resultado = self.engine.buscar_notas_por_descripcion_exacta(
            grado, materia, trimestre, tipo, desc)

        if not resultado:
            messagebox.showerror(
                "Error", "No se encontraron las notas en el archivo.")
            return

        self.col_a_modificar = resultado["columna"]
        notas_existentes = resultado["notas"]

        for id_est, entries_list in self.entradas_notas.items():
            for entry in entries_list:
                entry.delete(0, 'end')
            if id_est in notas_existentes:
                entries_list[0].insert(0, str(notas_existentes[id_est]))

    def actualizar_notas(self):
        if not self.col_a_modificar:
            return

        grado = self.combo_grado.get()
        materia = self.combo_materia.get()
        notas_guardar = self._recopilar_notas_validadas()
        if notas_guardar is None:
            return

        for entries_list in self.entradas_notas.values():
            for entry in entries_list:
                entry.delete(0, 'end')
        self.btn_actualizar.configure(
            text="Actualizando en fondo...", state="disabled")

        def tarea_fondo_act():
            exito = self.engine.actualizar_notas_existentes(
                grado, materia, self.col_a_modificar, notas_guardar)
            self.after(0, lambda: self.finalizar_actualizacion(exito))

        threading.Thread(target=tarea_fondo_act, daemon=True).start()

    def finalizar_actualizacion(self, exito):
        self.btn_actualizar.configure(text="🔄 ACTUALIZAR EXCEL",
                                      state="normal")
        self.col_a_modificar = None
        if exito:
            root = self.winfo_toplevel()
            if hasattr(root, "mostrar_toast"):
                root.mostrar_toast("✓ Calificaciones actualizadas en Excel", color="#10B981")
        else:
            messagebox.showerror(
                "Error", "No se pudieron actualizar las notas.")

    def abrir_modal_programar_tarea(self):
        from rdsecurity import cargar_config_segura
        import tareas
        import datetime
        import os
        
        grado = self.combo_grado.get()
        materia = self.combo_materia.get()
        tipo_nota = self.combo_tipo.get()
        desc = self.entry_desc.get().strip()
        if desc.startswith("Automático"):
            desc = ""
        fecha_nota = self.entry_fecha.get().strip()
        
        tipo_map = {
            "Diaria / Parcial": "Parcial",
            "Apreciación": "Apreciación",
            "Examen": "Examen"
        }
        tipo_tarea = tipo_map.get(tipo_nota, "Otro")
        
        cfg = cargar_config_segura({})
        ano = cfg.get("ano_lectivo", str(datetime.datetime.now().year))
        
        fecha_limite = ""
        if fecha_nota and "-" in fecha_nota:
            partes = fecha_nota.split("-")
            if len(partes) == 2:
                fecha_limite = f"{partes[1]}-{partes[0]}-{ano}"
        
        if not fecha_limite:
            fecha_limite = datetime.datetime.now().strftime(f"%d-%m-%Y")
            
        modal = ctk.CTkToplevel(self)
        modal.title("Programar Nueva Tarea")
        modal.geometry("720x400")
        modal.resizable(False, False)
        modal.transient(self)
        modal.grab_set()
        
        from config import establecer_icono_ventana
        establecer_icono_ventana(modal)
        
        modal.update_idletasks()
        w = 720
        h = 400
        x = self.winfo_toplevel().winfo_x() + (self.winfo_toplevel().winfo_width() // 2) - (w // 2)
        y = self.winfo_toplevel().winfo_y() + (self.winfo_toplevel().winfo_height() // 2) - (h // 2)
        modal.geometry(f"{w}x{h}+{x}+{y}")
        
        # Main containers
        container = ctk.CTkFrame(modal, fg_color="transparent")
        container.pack(fill="both", expand=True, padx=20, pady=20)
        
        # Título
        ctk.CTkLabel(container, text="Programar Nueva Tarea", font=("Segoe UI", 16, "bold"), text_color="#00DDEB").pack(pady=(0, 10))
        
        left_col = ctk.CTkFrame(container, fg_color="transparent")
        left_col.pack(side="left", fill="both", expand=True, padx=(0, 15))
        
        right_col = ctk.CTkFrame(container, fg_color="transparent")
        right_col.pack(side="right", fill="both", expand=True, padx=(15, 0))
        
        ctk.CTkLabel(left_col, text="Grados/Grupos:", font=("Segoe UI", 12, "bold")).pack(anchor="w", pady=(0, 5))
        
        var_todos = ctk.BooleanVar(value=False)
        def toggle_todos():
            val = var_todos.get()
            for chk in grado_checkboxes.values():
                if val:
                    chk.select()
                else:
                    chk.deselect()
                    
        chk_todos = ctk.CTkCheckBox(left_col, text="Seleccionar todos", variable=var_todos, font=("Segoe UI", 11))
        chk_todos.configure(command=toggle_todos)
        chk_todos.pack(anchor="w", pady=2)
        
        grados_activos = self.engine.obtener_grados_activos() or ["Sin datos"]
        grados_frame = ctk.CTkScrollableFrame(left_col, height=220, fg_color="#1E2D42", border_color="#2A3B50", border_width=1)
        grados_frame.pack(fill="both", expand=True, pady=5)
        
        grado_checkboxes = {}
        for g in grados_activos:
            chk = ctk.CTkCheckBox(grados_frame, text=g, font=("Segoe UI", 11))
            chk.pack(anchor="w", padx=5, pady=4)
            if g == grado:
                chk.select()
            grado_checkboxes[g] = chk
            
        def guardar():
            tit = entry_t.get().strip()
            fl = entry_f.get().strip()
            if not tit:
                messagebox.showwarning("Atención", "Escriba un título para la tarea.")
                return
            
            try:
                partes = fl.split("-")
                if len(partes) != 3:
                    raise ValueError("Formato de fecha inválido")
                dt_limite = datetime.date(int(partes[2]), int(partes[1]), int(partes[0]))
            except Exception:
                messagebox.showwarning("Atención", "La fecha debe estar en formato DD-MM-YYYY (Ej: 15-06-2026).")
                return
                
            from utils.date_helpers import obtener_rango_fechas_trimestre
            try:
                t1_start, _ = obtener_rango_fechas_trimestre("Trimestre 1")
                _, t3_end = obtener_rango_fechas_trimestre("Trimestre 3")
                if not (t1_start <= dt_limite <= t3_end):
                    messagebox.showwarning(
                        "Fecha fuera de rango", 
                        f"La fecha límite ({fl}) está fuera del período escolar.\n\n"
                        f"Rango escolar: {t1_start.strftime('%d-%m-%Y')} a {t3_end.strftime('%d-%m-%Y')}.\n"
                        "Por favor ingrese una fecha dentro de este rango."
                    )
                    return
            except Exception:
                pass
            
            grados_seleccionados = [g for g, chk in grado_checkboxes.items() if chk.get()]
            if not grados_seleccionados:
                messagebox.showwarning("Atención", "Seleccione al menos un Grado/Grupo.")
                return
                
            for g in grados_seleccionados:
                tareas.agregar_tarea(tit, g, materia, combo_t.get(), fl, f"Programada desde Notas. Tipo: {tipo_nota}")
                
            modal.destroy()
            
            if hasattr(self, 'scroll_tareas_tab'):
                self.actualizar_tab_tareas()
                
            root = self.winfo_toplevel()
            if hasattr(root, "mostrar_toast"):
                root.mostrar_toast("✓ Tarea programada exitosamente", color="#10B981")
            if hasattr(root, "main_app") and hasattr(root.main_app, "_frames"):
                dashboard = root.main_app._frames.get("DashboardFrame")
                if dashboard and hasattr(dashboard, "actualizar_vista"):
                    dashboard.actualizar_vista()

        ctk.CTkLabel(right_col, text="Materia:", font=("Segoe UI", 12, "bold")).pack(anchor="w", pady=(0, 2))
        entry_m = ctk.CTkEntry(right_col, height=30)
        entry_m.insert(0, materia)
        entry_m.configure(state="disabled")
        entry_m.pack(fill="x", pady=(0, 8))
        
        ctk.CTkLabel(right_col, text="Título/Tema de la Tarea:", font=("Segoe UI", 12, "bold")).pack(anchor="w", pady=(0, 2))
        entry_t = ctk.CTkEntry(right_col, height=30)
        entry_t.insert(0, desc)
        entry_t.pack(fill="x", pady=(0, 8))
        
        ctk.CTkLabel(right_col, text="Tipo:", font=("Segoe UI", 12, "bold")).pack(anchor="w", pady=(0, 2))
        combo_t = ctk.CTkOptionMenu(right_col, values=["Parcial", "Apreciación", "Examen", "Otro"], height=30)
        combo_t.set(tipo_tarea)
        combo_t.pack(fill="x", pady=(0, 8))
        
        ctk.CTkLabel(right_col, text="Fecha Límite (DD-MM-YYYY):", font=("Segoe UI", 12, "bold")).pack(anchor="w", pady=(0, 2))
        entry_f = ctk.CTkEntry(right_col, height=30)
        entry_f.insert(0, fecha_limite)
        entry_f.pack(fill="x", pady=(0, 15))
        
        btn_save = ctk.CTkButton(right_col, text="Agendar Tarea", fg_color="#10B981", hover_color="#059669", height=32, command=guardar)
        btn_save.pack(fill="x", pady=(5, 0))

    def generar_informe_estudiante(self, est):
        from rdsecurity import cargar_config_segura
        from documentos_maestro import generar_informe_calificaciones
        from utils.footer_utils import abrir_documento
        import datetime
        
        grado = self.combo_grado.get()
        # Recopilar materias con al menos una nota registrada en T1, T2 o T3
        materias_con_nota = set()
        notas_por_trimestre = {}
        for t_num in [1, 2, 3]:
            notas_t = self.engine.obtener_notas_estudiante(est["nombre"], grado, trimestre=t_num)
            notas_por_trimestre[t_num] = notas_t
            for mat_nombre, nota_val in notas_t.items():
                try:
                    val = float(nota_val)
                    if val >= 1.0:
                        materias_con_nota.add(mat_nombre)
                except (TypeError, ValueError):
                    pass

        # Fallback: si el alumno no tiene ninguna nota en todo el año, usar todas las del grado
        if not materias_con_nota:
            all_mats = self.engine.obtener_materias_por_grado(grado)
            materias_con_nota = set([m for m in all_mats if m not in ["Sin materias registradas", "Sin materias", "No hay materias", "General"]])

        cfg = cargar_config_segura({})
        trimestres = {}
        for t_key, t_num in [("T1", 1), ("T2", 2), ("T3", 3)]:
            notas_t = notas_por_trimestre[t_num]
            materias = []
            for mat_nombre in sorted(list(materias_con_nota)):
                nota_val = notas_t.get(mat_nombre, None)
                if nota_val is not None:
                    try:
                        val = float(nota_val)
                    except (TypeError, ValueError):
                        val = 0.0
                else:
                    val = 0.0
                
                materias.append({
                    "nombre": mat_nombre,
                    "nota": val,
                    "estado": "APROBADO" if val >= 3.0 else "REPROBADO"
                })
            prom = sum(m["nota"] for m in materias if m["nota"] > 0) / len([m for m in materias if m["nota"] > 0]) if any(m["nota"] > 0 for m in materias) else 0
            trimestres[t_key] = {"materias": materias, "promedio": round(prom, 1)}
        
        datos = {
            "alumno":          est["nombre"],
            "cedula":          est.get("cedula", ""),
            "grado":           grado,
            "docente_nombre":  cfg.get("docente_nombre", ""),
            "escuela_nombre":  cfg.get("escuela_nombre", ""),
            "ano_lectivo":     cfg.get("ano_lectivo", "2026"),
            "fecha":           datetime.datetime.now().strftime("%d-%m-%Y"),
            "trimestres":      trimestres,
            "observacion_general": "",
            "estado_final":    "EN PROCESO"
        }
        proms = [trimestres[tk]["promedio"] for tk in trimestres if trimestres[tk].get("promedio")]
        if proms:
            prom_global = sum(proms) / len(proms)
            datos["estado_final"] = "APROBADO" if prom_global >= 3.0 else "REPROBADO"
            
        ruta = generar_informe_calificaciones(datos)
        if ruta:
            abrir_documento(ruta)
            messagebox.showinfo("✓ Informe Generado", f"Informe de calificaciones individual generado:\n{ruta}")

    def abrir_tabla_puntos_modal(self):
        try:
            pts_max = int(float(self.entry_pts_max.get()))
        except Exception:
            pts_max = 40
            
        pct_str = self.combo_exigencia.get().replace("%", "")
        try:
            exigencia = float(pct_str) / 100.0
        except Exception:
            exigencia = 0.60
            
        modal = ctk.CTkToplevel(self)
        modal.title(f"⚖️ Tabla de Escala ({pts_max} pts, {pct_str}%)")
        modal.geometry("300x450")
        modal.resizable(False, False)
        modal.transient(self)
        
        from config import establecer_icono_ventana
        establecer_icono_ventana(modal)
        
        modal.update_idletasks()
        w = modal.winfo_width()
        h = modal.winfo_height()
        x = self.winfo_toplevel().winfo_x() + (self.winfo_toplevel().winfo_width() // 2) - (w // 2)
        y = self.winfo_toplevel().winfo_y() + (self.winfo_toplevel().winfo_height() // 2) - (h // 2)
        modal.geometry(f"+{x}+{y}")
        
        ctk.CTkLabel(modal, text=f"Conversión de Puntos a Nota", font=("Segoe UI", 13, "bold")).pack(pady=10)
        
        scroll = ctk.CTkScrollableFrame(modal, fg_color="transparent")
        scroll.pack(fill="both", expand=True, padx=15, pady=5)
        
        # Header of table
        hdr = ctk.CTkFrame(scroll, fg_color="#253650")
        hdr.pack(fill="x", pady=2)
        ctk.CTkLabel(hdr, text="Puntos Obtenidos", width=120, font=("Segoe UI", 11, "bold")).pack(side="left", padx=10)
        ctk.CTkLabel(hdr, text="Nota MEDUCA", width=100, font=("Segoe UI", 11, "bold")).pack(side="right", padx=10)
        
        for p in range(pts_max, -1, -1):
            nota = self.calcular_nota_meduca_puntos(p, pts_max, exigencia)
            row = ctk.CTkFrame(scroll, fg_color="#1A2638" if p % 2 == 0 else "transparent")
            row.pack(fill="x", pady=1)
            
            ctk.CTkLabel(row, text=f"{p} pts", width=120).pack(side="left", padx=10)
            color_nota = "#10B981" if nota >= 3.0 else "#EF4444"
            ctk.CTkLabel(row, text=f"{nota:.1f}", width=100, text_color=color_nota, font=("Segoe UI", 11, "bold")).pack(side="right", padx=10)

    def al_cambiar_puntos_estudiante(self, id_est):
        if not self.var_usar_puntos.get():
            return
            
        entries = self.entradas_notas[id_est]
        if len(entries) < 2:
            return
            
        entry_nota = entries[0]
        entry_pts = entries[1]
        
        val = entry_pts.get().strip()
        if not val:
            entry_nota.delete(0, 'end')
            return
            
        try:
            pts_max = float(self.entry_pts_max.get())
        except (ValueError, TypeError):
            pts_max = 40.0
            
        pct_str = self.combo_exigencia.get().replace("%", "")
        try:
            exigencia = float(pct_str) / 100.0
        except (ValueError, TypeError):
            exigencia = 0.60
            
        try:
            pts_obtenidos = float(val.replace(",", "."))
            nota = self.calcular_nota_meduca_puntos(pts_obtenidos, pts_max, exigencia)
            entry_nota.delete(0, 'end')
            entry_nota.insert(0, f"{nota:.1f}")
        except Exception:
            pass

    def calcular_nota_meduca_puntos(self, puntos, max_puntos, exigencia=0.6):
        if max_puntos <= 0:
            return 1.0
        try:
            pts = float(puntos)
        except (ValueError, TypeError):
            return 1.0
        
        if pts < 0:
            pts = 0
        if pts > max_puntos:
            pts = max_puntos
            
        p_aprob = exigencia * max_puntos
        
        if pts < p_aprob:
            if p_aprob == 0:
                nota = 3.0
            else:
                nota = 1.0 + 2.0 * (pts / p_aprob)
        else:
            den = max_puntos - p_aprob
            if den == 0:
                nota = 5.0
            else:
                nota = 3.0 + 2.0 * ((pts - p_aprob) / den)
                
        nota = round(nota, 1)
        if nota < 1.0:
            nota = 1.0
        if nota > 5.0:
            nota = 5.0
        return nota

    def recalcular_escala_visual(self):
        for id_est in self.entradas_notas.keys():
            self.al_cambiar_puntos_estudiante(id_est)

    def al_cambiar_modo_puntos(self):
        state = "normal" if self.var_usar_puntos.get() else "disabled"
        self.entry_pts_max.configure(state=state)
        self.combo_exigencia.configure(state=state)
        self.btn_ver_tabla_puntos.configure(state=state)
        self.cargar_estudiantes()

    def al_presionar_enter(self, id_est):
        self.al_presionar_abajo(id_est)
        return "break"

    def asegurar_visible(self, row_widget):
        try:
            self.scroll_estudiantes.update_idletasks()
            y = row_widget.winfo_y()
            h = row_widget.winfo_height()
            
            canvas = self.scroll_estudiantes._parent_canvas
            bbox = canvas.bbox("all")
            inner_height = bbox[3] if bbox else canvas.winfo_reqheight()
            canvas_height = canvas.winfo_height()
            
            current_top, current_bottom = canvas.yview()
            
            target_top_frac = y / (inner_height or 1)
            target_bottom_frac = (y + h) / (inner_height or 1)
            
            if target_top_frac < current_top:
                canvas.yview_moveto(target_top_frac)
            elif target_bottom_frac > current_bottom:
                view_size = current_bottom - current_top
                new_top = max(0.0, target_bottom_frac - view_size)
                canvas.yview_moveto(new_top)
        except Exception as e:
            print(f"Error en asegurar_visible: {e}")

    def al_presionar_abajo(self, id_est):
        ids = sorted(list(self.entradas_notas.keys()))
        try:
            curr_idx = ids.index(id_est)
            if curr_idx + 1 < len(ids):
                next_id = ids[curr_idx + 1]
                entries = self.entradas_notas[next_id]
                next_entry = entries[1] if (len(entries) > 1 and self.var_usar_puntos.get()) else entries[0]
                next_entry.focus_set()
                next_entry.select_range(0, 'end')
                self.asegurar_visible(next_entry.master)
        except ValueError:
            pass

    def al_presionar_arriba(self, id_est):
        ids = sorted(list(self.entradas_notas.keys()))
        try:
            curr_idx = ids.index(id_est)
            if curr_idx - 1 >= 0:
                prev_id = ids[curr_idx - 1]
                entries = self.entradas_notas[prev_id]
                prev_entry = entries[1] if (len(entries) > 1 and self.var_usar_puntos.get()) else entries[0]
                prev_entry.focus_set()
                prev_entry.select_range(0, 'end')
                self.asegurar_visible(prev_entry.master)
        except ValueError:
            pass

    def actualizar_vista(self):
        """Recarga los grados activos y actualiza los estudiantes."""
        opciones = self.engine.obtener_grados_activos()
        if not opciones:
            opciones = ["7°", "8°", "9°"] if self.engine.modalidad == "premedia" else ["1°"]
        old_sel = self.combo_grado.get()
        self.combo_grado.configure(values=opciones)
        if old_sel in opciones:
            self.combo_grado.set(old_sel)
        else:
            self.combo_grado.set(opciones[0])
        self.al_cambiar_grado(self.combo_grado.get())

