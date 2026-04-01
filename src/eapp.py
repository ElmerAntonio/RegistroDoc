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

        if self.engine.modalidad == "premedia":
            lbl_texto = "Grado/Grupo:"
            opciones = self.engine.obtener_grados_activos()
            if not opciones:
                opciones = ["7°", "8°", "9°"]
        else:
            lbl_texto = "Materia/Grado:"
            # En primaria solo dan a 1 grado, las opciones son las materias
            g_ref = self.engine.obtener_grados_activos()
            g_ref = g_ref[0] if g_ref else ""
            opciones = self.engine.obtener_materias_por_grado(g_ref)
            if not opciones or opciones[0] == "Sin materias registradas":
                opciones = ["Español", "Matemáticas"]

        ctk.CTkLabel(top, text=lbl_texto, font=(
            "Segoe UI", 16, "bold")).pack(side="left")

        self.combo_grado = ctk.CTkOptionMenu(
            top, values=opciones, command=self.al_cambiar_grado)
        self.combo_grado.pack(side="left", padx=10)

        self.scroll_estudiantes = ctk.CTkScrollableFrame(
            frame_izq, fg_color="transparent", orientation="horizontal")
        self.scroll_estudiantes_inner = ctk.CTkScrollableFrame(
            self.scroll_estudiantes, fg_color="transparent")
        self.scroll_estudiantes_inner.pack(fill="both", expand=True)
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

        self.tabs = ctk.CTkTabview(frame_der)
        self.tabs.pack(fill="both", expand=True, padx=10, pady=10)

        tab_nueva = self.tabs.add("Nueva Nota")
        tab_mod = self.tabs.add("Modificar")

        # ====== TAB: NUEVA NOTA ======
        self.combo_trimestre = ctk.CTkOptionMenu(
            tab_nueva, values=["Trimestre 1", "Trimestre 2", "Trimestre 3"])
        self.combo_trimestre.pack(fill="x", padx=10, pady=10)

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
        self.entry_desc = ctk.CTkEntry(tab_nueva, textvariable=self.var_desc)
        self.entry_desc.pack(fill="x", padx=10, pady=5)

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
        self.btn_guardar_nueva.pack(pady=20, padx=10, fill="x")

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

    def actualizar_contador_desc(self, *args):
        texto = self.var_desc.get()
        if len(texto) > 25:
            self.var_desc.set(texto[:25])
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
        for w in self.scroll_estudiantes_inner.winfo_children():
            w.destroy()
        self.entradas_notas.clear()
        self.col_a_modificar = None

        # Header Row
        header_row = ctk.CTkFrame(
            self.scroll_estudiantes_inner, fg_color="transparent")
        header_row.pack(fill="x", pady=(0, 10))
        ctk.CTkLabel(header_row, text="N°", width=30, font=(
            "Segoe UI", 12, "bold")).pack(side="left")
        ctk.CTkLabel(header_row, text="Nombre", width=200, anchor="w",
                     font=("Segoe UI", 12, "bold")).pack(side="left")

        # Casillas de notas dinámicas
        casillas = ctk.CTkFrame(header_row, fg_color="transparent")
        casillas.pack(side="left", padx=10, fill="x", expand=True)

        if self.engine.modalidad == "premedia":
            # 12 Parciales + 9 Apreciacion + 2 Examen
            f_parciales = ctk.CTkFrame(casillas, fg_color="#1A3352")
            f_parciales.pack(side="left", padx=2)
            ctk.CTkLabel(f_parciales, text="Diaria/Parcial (12)",
                         font=("Segoe UI", 10)).pack()
            row_p = ctk.CTkFrame(f_parciales, fg_color="transparent")
            row_p.pack()
            for i in range(12):
                ctk.CTkLabel(row_p, text=str(i+1),
                             width=25).pack(side="left", padx=1)

            f_apreciacion = ctk.CTkFrame(casillas, fg_color="#1E3A5F")
            f_apreciacion.pack(side="left", padx=2)
            ctk.CTkLabel(f_apreciacion, text="Apreciación (9)",
                         font=("Segoe UI", 10)).pack()
            row_a = ctk.CTkFrame(f_apreciacion, fg_color="transparent")
            row_a.pack()
            for i in range(9):
                ctk.CTkLabel(row_a, text=str(i+1),
                             width=25).pack(side="left", padx=1)

            f_examen = ctk.CTkFrame(casillas, fg_color="#0F2744")
            f_examen.pack(side="left", padx=2)
            ctk.CTkLabel(f_examen, text="Examen (2)",
                         font=("Segoe UI", 10)).pack()
            row_e = ctk.CTkFrame(f_examen, fg_color="transparent")
            row_e.pack()
            for i in range(2):
                ctk.CTkLabel(row_e, text=str(i+1),
                             width=25).pack(side="left", padx=1)
        else:
            # Primaria: 28 planas
            f_primaria = ctk.CTkFrame(casillas, fg_color="#1A3352")
            f_primaria.pack(side="left", fill="x", expand=True)
            ctk.CTkLabel(f_primaria, text="Notas Trimestrales (28)",
                         font=("Segoe UI", 10)).pack()
            row_pr = ctk.CTkFrame(f_primaria, fg_color="transparent")
            row_pr.pack()
            for i in range(28):
                ctk.CTkLabel(row_pr, text=str(i+1),
                             width=25).pack(side="left", padx=1)

        ests = self.engine.obtener_estudiantes_completos(grado)
        for est in ests:
            row = ctk.CTkFrame(self.scroll_estudiantes_inner,
                               fg_color="transparent")
            row.pack(fill="x", pady=2)
            ctk.CTkLabel(row, text=f"{est['id']}.", width=30).pack(side="left")
            ctk.CTkLabel(row, text=est['nombre'],
                         width=200, anchor="w").pack(side="left")

            casillas_row = ctk.CTkFrame(row, fg_color="transparent")
            casillas_row.pack(side="left", padx=10, fill="x", expand=True)

            self.entradas_notas[est['id']] = []
            num_boxes = 23 if self.engine.modalidad == "premedia" else 28

            for i in range(num_boxes):
                entry = ctk.CTkEntry(
                    casillas_row,
                    width=25,
                    height=25,
                    justify="center",
                    placeholder_text="-",
                    font=(
                        "Segoe UI",
                        11))
                entry.pack(side="left", padx=1)
                self.entradas_notas[est['id']].append(entry)

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

    def _recopilar_notas_validadas(self):
        notas_guardar = {}
        for id_est, entries_list in self.entradas_notas.items():
            # Only read the first one for backwards compatibility
            # or the last active if requested
            # Since the original code had 1 entry, and the prompt implies
            # a grid of many, we adapt it to pick up whatever the user typed.
            # For simplicity, we get the first non-empty.
            val = ""
            for entry in entries_list:
                if entry.get().strip():
                    val = entry.get().strip()
                    break

            if val:
                valido, nota, msj = validar_nota_meduca(val)
                if valido:
                    notas_guardar[id_est] = nota
                else:
                    messagebox.showerror("Error", f"Error en la nota: {msj}")
                    return None

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
        if not exito:
            messagebox.showerror(
                "Error", "No se pudieron actualizar las notas.")
