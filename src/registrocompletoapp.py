import os
import threading
import customtkinter as ctk
from theme import C, FONT_TITLE, FONT_BODY
from rdsecurity import validar_nota_meduca

# DICCIONARIOS DE TRADUCCIÓN PANTALLA <-> EXCEL
EXCEL_A_UI = {".": "P", "-": "A", "T": "T", "E": "E"}

class RegistroCompletoFrame(ctk.CTkFrame):
    def __init__(self, master, engine, **kwargs):
        super().__init__(master, fg_color="transparent", **kwargs)
        self.engine = engine
        self.loading = False

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)

        self.crear_filtros()
        self.crear_tabs()
        
        # Cargar datos iniciales
        self.actualizar_vista()

    def crear_filtros(self):
        self.f_top = ctk.CTkFrame(self, fg_color=C["card"], corner_radius=10)
        self.f_top.grid(row=0, column=0, sticky="ew", padx=15, pady=(15, 10))

        # Título
        lbl_title = ctk.CTkLabel(self.f_top, text="📋 Registro General Consolidado",
                                 font=(FONT_TITLE, 18, "bold"), text_color=C["cian"])
        lbl_title.pack(side="left", padx=20, pady=15)

        # Combo de Grado
        grados = self.engine.obtener_grados_activos() or ["Sin datos"]
        self.combo_grado = ctk.CTkOptionMenu(self.f_top, values=grados, command=self.al_cambiar_grado)
        self.combo_grado.pack(side="right", padx=(5, 20), pady=15)
        self.combo_grado.set(grados[0])

        # Combo de Trimestre
        self.combo_trimestre = ctk.CTkOptionMenu(self.f_top, 
                                                 values=["Trimestre 1", "Trimestre 2", "Trimestre 3"],
                                                 command=lambda _: self.cargar_datos_asincrono())
        self.combo_trimestre.pack(side="right", padx=5, pady=15)

        # Combo de Materia
        materias = self.engine.obtener_materias_por_grado(grados[0]) if grados[0] != "Sin datos" else ["No hay materias"]
        self.combo_materia = ctk.CTkOptionMenu(self.f_top, values=materias, command=lambda _: self.cargar_datos_asincrono())
        self.combo_materia.pack(side="right", padx=5, pady=15)
        if materias:
            self.combo_materia.set(materias[0])

    def al_cambiar_grado(self, grado):
        materias = self.engine.obtener_materias_por_grado(grado)
        if materias:
            self.combo_materia.configure(values=materias)
            self.combo_materia.set(materias[0])
        else:
            self.combo_materia.configure(values=["No hay materias"])
            self.combo_materia.set("No hay materias")
        self.cargar_datos_asincrono()

    def crear_tabs(self):
        self.tabs = ctk.CTkTabview(self)
        self.tabs.grid(row=1, column=0, sticky="nsew", padx=15, pady=(0, 15))

        self.tab_notas = self.tabs.add("📝 Calificaciones del Grupo")
        self.tab_asis = self.tabs.add("📅 Asistencia del Grupo")

        # Configurar pestañas
        self.tab_notas.grid_columnconfigure(0, weight=1)
        self.tab_notas.grid_rowconfigure(1, weight=1)

        self.tab_asis.grid_columnconfigure(0, weight=1)
        self.tab_asis.grid_rowconfigure(1, weight=1)

        # Sub-frames para Notas
        self.stats_notas_frame = ctk.CTkFrame(self.tab_notas, fg_color="transparent")
        self.stats_notas_frame.grid(row=0, column=0, sticky="ew", padx=10, pady=10)

        self.scroll_notas = ctk.CTkScrollableFrame(self.tab_notas, fg_color="#0D1F35")
        self.scroll_notas.grid(row=1, column=0, sticky="nsew", padx=10, pady=(0, 10))

        # Sub-frames para Asistencia
        self.stats_asis_frame = ctk.CTkFrame(self.tab_asis, fg_color="transparent")
        self.stats_asis_frame.grid(row=0, column=0, sticky="ew", padx=10, pady=10)

        self.scroll_asis = ctk.CTkScrollableFrame(self.tab_asis, fg_color="#0D1F35")
        self.scroll_asis.grid(row=1, column=0, sticky="nsew", padx=10, pady=(0, 10))

    def actualizar_vista(self):
        self.cargar_datos_asincrono()

    def cargar_datos_asincrono(self):
        if self.loading:
            return
        self.loading = True

        # Limpiar e insertar cargadores
        self._limpiar_scroll(self.scroll_notas)
        self._limpiar_scroll(self.scroll_asis)

        lbl_cargando_n = ctk.CTkLabel(self.scroll_notas, text="🔄 Cargando registros de notas desde Excel...", 
                                      font=(FONT_BODY, 13, "bold"), text_color=C["cian"])
        lbl_cargando_n.pack(pady=40)

        lbl_cargando_a = ctk.CTkLabel(self.scroll_asis, text="🔄 Cargando asistencia desde Excel...", 
                                      font=(FONT_BODY, 13, "bold"), text_color=C["cian"])
        lbl_cargando_a.pack(pady=40)

        # Parámetros activos
        grado = self.combo_grado.get()
        trimestre = self.combo_trimestre.get()
        materia = self.combo_materia.get()

        def background_thread():
            try:
                data = self._obtener_datos_de_excel(grado, trimestre, materia)
                self.after(0, lambda: self._renderizar_datos(data))
            except Exception as e:
                print(f"[!] Error cargando datos consolidados: {e}")
                self.after(0, lambda: self._renderizar_error())

        t = threading.Thread(target=background_thread, daemon=True)
        t.start()

    def _obtener_datos_de_excel(self, grado, trimestre, materia):
        wb = self.engine._wb_cache
        if not wb:
            return None

        # Estudiantes
        estudiantes = self.engine.obtener_estudiantes_completos(grado, wb=wb)

        # 1. NOTAS
        headers_notas = []
        rows_notas = []
        stats_notas = {"actividades": 0, "promedio_grupal": 0.0, "tasa_aprobacion": 0.0, "max_nota": 0.0}

        nombre_hoja_prom = self.engine._encontrar_hoja_prom(wb, grado, materia)
        if nombre_hoja_prom:
            ws_prom = wb[nombre_hoja_prom]
            types = ["Diaria / Parcial", "Apreciación", "Examen"]
            columns_data = []
            
            for t in types:
                col_inicio, col_fin = self.engine._obtener_rango_columnas(ws_prom, trimestre, t)
                if col_inicio is not None and col_fin is not None:
                    for c in range(col_inicio, col_fin + 1):
                        desc_val = ws_prom.cell(row=self.engine.fila_desc, column=c).value
                        desc = str(desc_val).replace('\n', ' ').strip() if desc_val else f"Act. {c - col_inicio + 1}"
                        columns_data.append((c, desc, t))

            promedios_finales = self.engine.obtener_promedios_reales(grado, materia, trimestre, wb=wb)
            all_valid_notes = []

            for est in estudiantes:
                row_idx = est["id"] + 4
                student_grades = []
                for col_idx, desc, t in columns_data:
                    val = ws_prom.cell(row=row_idx, column=col_idx).value
                    valido, nota, _ = validar_nota_meduca(val)
                    student_grades.append(nota if valido else "")
                    if valido:
                        all_valid_notes.append(nota)

                prom_final = promedios_finales.get(est["nombre"], "")
                rows_notas.append({
                    "id": est["id"],
                    "nombre": est["nombre"],
                    "notas": student_grades,
                    "promedio": prom_final
                })

            headers_notas = [(desc, t) for c, desc, t in columns_data]
            stats_notas["actividades"] = len(headers_notas)
            if all_valid_notes:
                stats_notas["promedio_grupal"] = round(sum(all_valid_notes) / len(all_valid_notes), 2)
                stats_notas["max_nota"] = round(max(all_valid_notes), 2)

            promedios_validos = [p for p in promedios_finales.values() if isinstance(p, (int, float))]
            if promedios_validos:
                aprobados = sum(1 for p in promedios_validos if p >= 3.0)
                stats_notas["tasa_aprobacion"] = round((aprobados / len(promedios_validos)) * 100, 1)

        # 2. ASISTENCIA
        headers_asis = []
        rows_asis = []
        stats_asis = {"total_dias": 0, "asistencia_promedio": 0.0, "asistencia_perfecta": 0, "alerta_count": 0}

        fechas = self.engine.obtener_fechas_asistencia(grado, trimestre, wb=wb)
        if fechas:
            hoja_asis = self.engine._encontrar_hoja_asistencia(wb, grado)
            if hoja_asis:
                ws_asis = wb[hoja_asis]
                mapa_trimestres = {"Trimestre 1": 2, "Trimestre 2": 45, "Trimestre 3": 88}
                fila_fechas = mapa_trimestres.get(trimestre, 2)

                cols_fechas = []
                for f in fechas:
                    col_found = None
                    for c in range(3, 61):
                        val = ws_asis.cell(row=fila_fechas, column=c).value
                        if val and str(val).strip() == f.strip():
                            col_found = c
                            break
                    if col_found:
                        cols_fechas.append((col_found, f))

                total_asis_pcts = []
                for est in estudiantes:
                    row_idx = fila_fechas + est["id"]
                    student_status = []
                    ausencias = 0
                    total_dias = 0
                    for col_idx, f in cols_fechas:
                        val = ws_asis.cell(row=row_idx, column=col_idx).value
                        ui_val = EXCEL_A_UI.get(val, "P") if val is not None else "P"
                        student_status.append(ui_val)
                        if val is not None and str(val).strip():
                            total_dias += 1
                            if val == "-":
                                ausencias += 1

                    pct = 100.0
                    if total_dias > 0:
                        pct = round(((total_dias - ausencias) / total_dias) * 100, 1)

                    total_asis_pcts.append(pct)
                    if pct == 100.0 and total_dias > 0:
                        stats_asis["asistencia_perfecta"] += 1
                    if pct < 90.0 and total_dias > 0:
                        stats_asis["alerta_count"] += 1

                    rows_asis.append({
                        "id": est["id"],
                        "nombre": est["nombre"],
                        "asistencia": student_status,
                        "porcentaje": pct,
                        "dias": total_dias
                    })

                headers_asis = [f for col_idx, f in cols_fechas]
                stats_asis["total_dias"] = len(headers_asis)
                if total_asis_pcts:
                    stats_asis["asistencia_promedio"] = round(sum(total_asis_pcts) / len(total_asis_pcts), 1)

        return {
            "headers_notas": headers_notas,
            "rows_notas": rows_notas,
            "stats_notas": stats_notas,
            "headers_asis": headers_asis,
            "rows_asis": rows_asis,
            "stats_asis": stats_asis
        }

    def _renderizar_datos(self, data):
        self.loading = False
        self._limpiar_scroll(self.scroll_notas)
        self._limpiar_scroll(self.scroll_asis)

        if not data:
            self._renderizar_error()
            return

        # ─── 1. NOTAS ───
        self._renderizar_stats_notas(data["stats_notas"])
        h_notas = data["headers_notas"]
        r_notas = data["rows_notas"]

        if not h_notas:
            ctk.CTkLabel(self.scroll_notas, text="No hay calificaciones registradas para esta materia y trimestre.",
                         font=(FONT_BODY, 13), text_color=C["texto_sec"]).pack(pady=40)
        else:
            # Encabezados
            f_header = ctk.CTkFrame(self.scroll_notas, fg_color="#1E3A5F", corner_radius=5)
            f_header.pack(fill="x", padx=5, pady=2)
            
            ctk.CTkLabel(f_header, text="N°", width=35, font=(FONT_BODY, 11, "bold")).pack(side="left", padx=5)
            ctk.CTkLabel(f_header, text="Nombre del Estudiante", width=180, anchor="w", font=(FONT_BODY, 11, "bold")).pack(side="left", padx=5)
            
            for desc, tipo in h_notas:
                # Mostrar descripción corta en el header
                lbl_h = ctk.CTkLabel(f_header, text=desc[:8], width=65, font=(FONT_BODY, 10, "bold"), text_color=C["cian"])
                lbl_h.pack(side="left", padx=3)
                
            ctk.CTkLabel(f_header, text="Promedio", width=80, font=(FONT_BODY, 11, "bold"), text_color=C["verde"]).pack(side="right", padx=10)

            # Estudiantes
            for i, row_data in enumerate(r_notas):
                bg = C["card"] if i % 2 == 0 else C["card_alt"]
                f_row = ctk.CTkFrame(self.scroll_notas, fg_color=bg, corner_radius=3)
                f_row.pack(fill="x", padx=5, pady=1)

                ctk.CTkLabel(f_row, text=f"{row_data['id']}", width=35, font=(FONT_BODY, 11)).pack(side="left", padx=5)
                ctk.CTkLabel(f_row, text=row_data['nombre'], width=180, anchor="w", font=(FONT_BODY, 11)).pack(side="left", padx=5)

                for nota in row_data["notas"]:
                    color = C["texto"]
                    if isinstance(nota, (int, float)):
                        if nota < 3.0:
                            color = C["rojo"]
                        elif nota >= 4.5:
                            color = C["verde"]
                    lbl_n = ctk.CTkLabel(f_row, text=f"{nota}" if nota != "" else "-", width=65, font=(FONT_BODY, 11), text_color=color)
                    lbl_n.pack(side="left", padx=3)

                prom = row_data["promedio"]
                prom_str = f"{prom:.2f}" if isinstance(prom, (int, float)) else (f"{prom}" if prom else "-")
                color_prom = C["texto"]
                if isinstance(prom, (int, float)):
                    if prom < 3.0:
                        color_prom = C["rojo"]
                    elif prom >= 4.0:
                        color_prom = C["verde"]
                
                ctk.CTkLabel(f_row, text=prom_str, width=80, font=(FONT_BODY, 11, "bold"), text_color=color_prom).pack(side="right", padx=10)

        # ─── 2. ASISTENCIA ───
        self._renderizar_stats_asis(data["stats_asis"])
        h_asis = data["headers_asis"]
        r_asis = data["rows_asis"]

        if not h_asis:
            ctk.CTkLabel(self.scroll_asis, text="No hay registros de asistencia para este grado y trimestre.",
                         font=(FONT_BODY, 13), text_color=C["texto_sec"]).pack(pady=40)
        else:
            # Encabezados
            f_header = ctk.CTkFrame(self.scroll_asis, fg_color="#1E3A5F", corner_radius=5)
            f_header.pack(fill="x", padx=5, pady=2)
            
            ctk.CTkLabel(f_header, text="N°", width=35, font=(FONT_BODY, 11, "bold")).pack(side="left", padx=5)
            ctk.CTkLabel(f_header, text="Nombre del Estudiante", width=180, anchor="w", font=(FONT_BODY, 11, "bold")).pack(side="left", padx=5)
            
            for f in h_asis:
                lbl_h = ctk.CTkLabel(f_header, text=f, width=65, font=(FONT_BODY, 10, "bold"), text_color=C["cian"])
                lbl_h.pack(side="left", padx=3)
                
            ctk.CTkLabel(f_header, text="Asistencia %", width=80, font=(FONT_BODY, 11, "bold"), text_color=C["verde"]).pack(side="right", padx=10)

            # Estudiantes
            for i, row_data in enumerate(r_asis):
                bg = C["card"] if i % 2 == 0 else C["card_alt"]
                f_row = ctk.CTkFrame(self.scroll_asis, fg_color=bg, corner_radius=3)
                f_row.pack(fill="x", padx=5, pady=1)

                ctk.CTkLabel(f_row, text=f"{row_data['id']}", width=35, font=(FONT_BODY, 11)).pack(side="left", padx=5)
                ctk.CTkLabel(f_row, text=row_data['nombre'], width=180, anchor="w", font=(FONT_BODY, 11)).pack(side="left", padx=5)

                for st in row_data["asistencia"]:
                    color = C["texto"]
                    if st == "A":
                        color = C["rojo"]
                    elif st == "T":
                        color = C["amarillo"]
                    elif st == "E":
                        color = C["acento2"]
                    elif st == "P":
                        color = C["verde"]
                    lbl_a = ctk.CTkLabel(f_row, text=st, width=65, font=(FONT_BODY, 11), text_color=color)
                    lbl_a.pack(side="left", padx=3)

                pct = row_data["porcentaje"]
                pct_str = f"{pct}%" if row_data["dias"] > 0 else "-"
                color_pct = C["verde"] if pct >= 90 else (C["amarillo"] if pct >= 80 else C["rojo"])
                
                ctk.CTkLabel(f_row, text=pct_str, width=80, font=(FONT_BODY, 11, "bold"), text_color=color_pct).pack(side="right", padx=10)

    def _renderizar_stats_notas(self, stats):
        for w in self.stats_notas_frame.winfo_children():
            w.destroy()
            
        # Grid layout for cards
        self.stats_notas_frame.grid_columnconfigure((0, 1, 2, 3), weight=1)
        
        cards = [
            ("📝 Actividades", f"{stats['actividades']}", "Evaluaciones Tomadas", C["cian"]),
            ("📊 Promedio Salón", f"{stats['promedio_grupal']:.2f}", "Promedio General", C["amarillo"]),
            ("🎓 Aprobación", f"{stats['tasa_aprobacion']}%", "Estudiantes con >= 3.0", C["verde"]),
            ("🏆 Nota Máxima", f"{stats['max_nota']:.1f}", "Rendimiento Superior", C["purpura"])
        ]
        
        for idx, (title, val, desc, color) in enumerate(cards):
            card = ctk.CTkFrame(self.stats_notas_frame, fg_color=C["card"], corner_radius=8, border_color=color, border_width=1)
            card.grid(row=0, column=idx, padx=5, pady=5, sticky="nsew")
            
            ctk.CTkLabel(card, text=title, font=(FONT_BODY, 12, "bold"), text_color=color).pack(anchor="w", padx=15, pady=(10, 2))
            ctk.CTkLabel(card, text=val, font=(FONT_TITLE, 22, "bold"), text_color=C["texto"]).pack(anchor="w", padx=15, pady=2)
            ctk.CTkLabel(card, text=desc, font=(FONT_BODY, 10), text_color=C["texto_sec"]).pack(anchor="w", padx=15, pady=(2, 10))

    def _renderizar_stats_asis(self, stats):
        for w in self.stats_asis_frame.winfo_children():
            w.destroy()
            
        self.stats_asis_frame.grid_columnconfigure((0, 1, 2, 3), weight=1)
        
        cards = [
            ("📅 Días Tomados", f"{stats['total_dias']}", "Fechas de Asistencia", C["cian"]),
            ("📈 Asistencia Prom.", f"{stats['asistencia_promedio']}%", "Asistencia del Grupo", C["verde"]),
            ("⭐ Perfecta (100%)", f"{stats['asistencia_perfecta']}", "Asistencia Inmaculada", C["amarillo"]),
            ("⚠️ En Alerta (<90%)", f"{stats['alerta_count']}", "Estudiantes con Inasistencias", C["rojo"])
        ]
        
        for idx, (title, val, desc, color) in enumerate(cards):
            card = ctk.CTkFrame(self.stats_asis_frame, fg_color=C["card"], corner_radius=8, border_color=color, border_width=1)
            card.grid(row=0, column=idx, padx=5, pady=5, sticky="nsew")
            
            ctk.CTkLabel(card, text=title, font=(FONT_BODY, 12, "bold"), text_color=color).pack(anchor="w", padx=15, pady=(10, 2))
            ctk.CTkLabel(card, text=val, font=(FONT_TITLE, 22, "bold"), text_color=C["texto"]).pack(anchor="w", padx=15, pady=2)
            ctk.CTkLabel(card, text=desc, font=(FONT_BODY, 10), text_color=C["texto_sec"]).pack(anchor="w", padx=15, pady=(2, 10))

    def _renderizar_error(self):
        self.loading = False
        self._limpiar_scroll(self.scroll_notas)
        self._limpiar_scroll(self.scroll_asis)
        
        ctk.CTkLabel(self.scroll_notas, text="⚠️ Ocurrió un error o no hay hoja de notas válida en el Excel para este grupo.",
                     font=(FONT_BODY, 13), text_color=C["rojo"]).pack(pady=40)
        ctk.CTkLabel(self.scroll_asis, text="⚠️ Ocurrió un error o no hay hoja de asistencia válida en el Excel para este grupo.",
                     font=(FONT_BODY, 13), text_color=C["rojo"]).pack(pady=40)

    def _limpiar_scroll(self, scroll):
        for w in scroll.winfo_children():
            w.destroy()
