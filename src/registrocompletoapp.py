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
        self.tab_habitos = self.tabs.add("🧠 Hábitos del Grupo")

        # Configurar pestañas
        self.tab_notas.grid_columnconfigure(0, weight=1)
        self.tab_notas.grid_rowconfigure(1, weight=1)

        self.tab_asis.grid_columnconfigure(0, weight=1)
        self.tab_asis.grid_rowconfigure(1, weight=1)

        self.tab_habitos.grid_columnconfigure(0, weight=1)
        self.tab_habitos.grid_rowconfigure(1, weight=1)

        # Sub-frames para Notas
        self.stats_notas_frame = ctk.CTkFrame(self.tab_notas, fg_color="transparent")
        self.stats_notas_frame.grid(row=0, column=0, sticky="ew", padx=10, pady=10)

        self.scroll_notas = ctk.CTkScrollableFrame(self.tab_notas, fg_color=C["card_alt"])
        self.scroll_notas.grid(row=1, column=0, sticky="nsew", padx=10, pady=(0, 10))

        # Sub-frames para Asistencia
        self.stats_asis_frame = ctk.CTkFrame(self.tab_asis, fg_color="transparent")
        self.stats_asis_frame.grid(row=0, column=0, sticky="ew", padx=10, pady=10)

        self.scroll_asis = ctk.CTkScrollableFrame(self.tab_asis, fg_color=C["card_alt"])
        self.scroll_asis.grid(row=1, column=0, sticky="nsew", padx=10, pady=(0, 10))

        # Sub-frames para Hábitos
        self.stats_habitos_frame = ctk.CTkFrame(self.tab_habitos, fg_color="transparent")
        self.stats_habitos_frame.grid(row=0, column=0, sticky="ew", padx=10, pady=10)

        self.scroll_habitos = ctk.CTkScrollableFrame(self.tab_habitos, fg_color=C["card_alt"])
        self.scroll_habitos.grid(row=1, column=0, sticky="nsew", padx=10, pady=(0, 10))

    def actualizar_vista(self):
        self.cargar_datos_asincrono()

    def cargar_datos_asincrono(self):
        if self.loading:
            return
        self.loading = True

        # Limpiar e insertar cargadores
        self._limpiar_scroll(self.scroll_notas)
        self._limpiar_scroll(self.scroll_asis)
        self._limpiar_scroll(self.scroll_habitos)

        lbl_cargando_n = ctk.CTkLabel(self.scroll_notas, text="🔄 Cargando registros de notas desde SQL...", 
                                      font=(FONT_BODY, 13, "bold"), text_color=C["cian"])
        lbl_cargando_n.pack(pady=40)

        lbl_cargando_a = ctk.CTkLabel(self.scroll_asis, text="🔄 Cargando asistencia desde SQL...", 
                                      font=(FONT_BODY, 13, "bold"), text_color=C["cian"])
        lbl_cargando_a.pack(pady=40)

        lbl_cargando_h = ctk.CTkLabel(self.scroll_habitos, text="🔄 Cargando hábitos desde SQL...", 
                                      font=(FONT_BODY, 13, "bold"), text_color=C["cian"])
        lbl_cargando_h.pack(pady=40)

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
        # NOTA: Aunque conserva el nombre por compatibilidad, ahora lee 100% de SQLite para alto rendimiento.
        cursor = self.engine.db_conn.cursor()
        
        # Obtener grado_id
        cursor.execute("SELECT id FROM grados WHERE nombre = ? AND seccion = 'A' AND modalidad = ?;", (grado, self.engine.modalidad))
        row_g = cursor.fetchone()
        if not row_g:
            return None
        grado_id = row_g[0]
        
        # Obtener materia_id
        cursor.execute("SELECT id FROM materias WHERE nombre = ? AND grado_id = ?;", (materia, grado_id))
        row_m = cursor.fetchone()
        if not row_m:
            return None
        materia_id = row_m[0]
        
        # Obtener estudiantes: incluir retirados que estaban activos en el trimestre consultado.
        try:
            import datetime
            anio = datetime.date.today().year
        except Exception:
            anio = 2026
        trim_inicio = {1: f"{anio}-03-01", 2: f"{anio}-06-01", 3: f"{anio}-09-01"}

        trim_num_prev = int(trimestre.replace("Trimestre ", ""))
        inicio_trim = trim_inicio.get(trim_num_prev, f"{anio}-01-01")

        # Incluir estudiantes activos + retirados que tengan fecha_retiro >= inicio del trimestre
        cursor.execute("""
            SELECT id, nombre, estado, fecha_retiro FROM estudiantes
            WHERE grado_id = ?
              AND (estado = 'Activo'
                   OR (estado = 'Retirado' AND (fecha_retiro IS NULL OR fecha_retiro >= ?)))
            ORDER BY CAST(id AS INTEGER);
        """, (grado_id, inicio_trim))
        estudiantes = []
        for r in cursor.fetchall():
            nombre_dec = self.engine.db_manager.desencriptar_campo(r[1])
            es_retirado = r[2] == "Retirado"
            estudiantes.append({
                "id":          int(r[0]) if str(r[0]).isdigit() else r[0],
                "nombre":      nombre_dec + " ⚠️(R)" if es_retirado else nombre_dec,
                "retirado":    es_retirado,
            })
            
        trim_num = int(trimestre.replace("Trimestre ", ""))
        
        # 1. NOTAS
        cursor.execute("""
            SELECT MIN(id), tipo, descripcion, fecha 
            FROM notas 
            WHERE materia_id = ? AND trimestre = ?
            GROUP BY tipo, descripcion, fecha
            ORDER BY MIN(id);
        """, (materia_id, trim_num))
        tasks = []
        for row in cursor.fetchall():
            tasks.append({
                "tipo": row[1],
                "descripcion": row[2],
                "fecha": row[3]
            })
            
        promedios_finales = self.engine.obtener_promedios_reales(grado, materia, trimestre)
        student_ids_str = [str(e["id"]) for e in estudiantes]
        placeholders = ",".join(["?"] * len(student_ids_str)) if student_ids_str else "NULL"
        
        # 1. NOTAS (Búsqueda en lote)
        grades_lookup = {}
        if student_ids_str:
            cursor.execute("""
                SELECT estudiante_id, tipo, descripcion, fecha, valor
                FROM notas
                WHERE materia_id = ? AND trimestre = ?;
            """, (materia_id, trim_num))
            for row in cursor.fetchall():
                est_id_db = int(row[0]) if str(row[0]).isdigit() else row[0]
                grades_lookup[(est_id_db, row[1], row[2], row[3])] = row[4]

        rows_notas = []
        all_valid_notes = []
        
        for est in estudiantes:
            student_grades = []
            for t_info in tasks:
                val = grades_lookup.get((est["id"], t_info["tipo"], t_info["descripcion"], t_info["fecha"]))
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
            
        # Formatear encabezados de notas con el título y la fecha en una línea inferior
        headers_notas = []
        for t_info in tasks:
            desc_with_date = f"{t_info['descripcion']}\n({t_info['fecha']})"
            headers_notas.append((desc_with_date, t_info["tipo"]))
            
        stats_notas = {"actividades": len(headers_notas), "promedio_grupal": 0.0, "tasa_aprobacion": 0.0, "max_nota": 0.0}
        if all_valid_notes:
            stats_notas["promedio_grupal"] = round(sum(all_valid_notes) / len(all_valid_notes), 2)
            stats_notas["max_nota"] = round(max(all_valid_notes), 2)
            
        promedios_validos = [p for p in promedios_finales.values() if isinstance(p, (int, float))]
        if promedios_validos:
            aprobados = sum(1 for p in promedios_validos if p >= 3.0)
            stats_notas["tasa_aprobacion"] = round((aprobados / len(promedios_validos)) * 100, 1)

        # 2. ASISTENCIA (Búsqueda en lote)
        headers_asis = []
        rows_asis = []
        stats_asis = {"total_dias": 0, "asistencia_promedio": 0.0, "asistencia_perfecta": 0, "alerta_count": 0}
        
        fechas = self.engine.obtener_fechas_asistencia(grado, trimestre)
        total_asis_pcts = []
        
        asis_lookup = {}
        if student_ids_str:
            cursor.execute(f"""
                SELECT estudiante_id, fecha, estado
                FROM asistencia
                WHERE estudiante_id IN ({placeholders}) AND trimestre = ?;
            """, (*student_ids_str, trim_num))
            for row in cursor.fetchall():
                est_id_db = int(row[0]) if str(row[0]).isdigit() else row[0]
                asis_lookup[(est_id_db, row[1])] = row[2]
        
        for est in estudiantes:
            student_status = []
            ausencias = 0
            total_dias = 0
            for f in fechas:
                val = asis_lookup.get((est["id"], f))
                
                # Mapear estado de SQLite al valor de la UI (. -> P, - -> A, T -> T, E -> E)
                ui_val = "P"
                if val == "-":
                    ui_val = "A"
                    ausencias += 1
                elif val == "T":
                    ui_val = "T"
                elif val == "E":
                    ui_val = "E"
                elif val == ".":
                    ui_val = "P"
                
                student_status.append(ui_val)
                if val is not None and str(val).strip():
                    total_dias += 1
                    
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
            
        headers_asis = fechas
        stats_asis["total_dias"] = len(headers_asis)
        if total_asis_pcts:
            stats_asis["asistencia_promedio"] = round(sum(total_asis_pcts) / len(total_asis_pcts), 1)

        # 3. HÁBITOS Y APTITUDES (Búsqueda en lote)
        from habapp import CRITERIOS_PRIMARIA, CRITERIOS_PREMEDIA_MEDIA
        es_primaria = any(g in grado for g in ["1°", "2°", "3°", "4°", "5°", "6°"])
        criterios = CRITERIOS_PRIMARIA if es_primaria else CRITERIOS_PREMEDIA_MEDIA
        
        rows_habitos = []
        stats_habitos = {"total_evaluaciones": 0, "s_count": 0, "r_count": 0, "x_count": 0}
        
        mapa_trim = {"Trimestre 1": 1, "Trimestre 2": 2, "Trimestre 3": 3}
        num_trim_actual = mapa_trim.get(trimestre, 1)
        
        habitos_trim_lookup = {}
        habitos_cont_lookup = {}
        if student_ids_str:
            cursor.execute(f"""
                SELECT estudiante_id, criterio_codigo, frecuencia, nota
                FROM habitos
                WHERE estudiante_id IN ({placeholders}) AND trimestre = ?;
            """, (*student_ids_str, num_trim_actual))
            for row in cursor.fetchall():
                est_id_db = int(row[0]) if str(row[0]).isdigit() else row[0]
                crit_code = row[1]
                frecuencia = row[2]
                nota = row[3]
                if frecuencia == "Trimestral":
                    habitos_trim_lookup[(est_id_db, crit_code)] = nota
                else:
                    if (est_id_db, crit_code) not in habitos_cont_lookup:
                        habitos_cont_lookup[(est_id_db, crit_code)] = []
                    habitos_cont_lookup[(est_id_db, crit_code)].append(nota)

        # Buscar hábitos de trimestres anteriores en lote
        habitos_prev_trim_lookup = {}
        habitos_prev_cont_lookup = {}
        if num_trim_actual > 1 and student_ids_str:
            prev_trims = list(range(1, num_trim_actual))
            prev_trims_placeholders = ",".join(["?"] * len(prev_trims))
            cursor.execute(f"""
                SELECT estudiante_id, trimestre, criterio_codigo, frecuencia, nota
                FROM habitos
                WHERE estudiante_id IN ({placeholders}) AND trimestre IN ({prev_trims_placeholders});
            """, (*student_ids_str, *prev_trims))
            for row in cursor.fetchall():
                est_id_db = int(row[0]) if str(row[0]).isdigit() else row[0]
                t_num = row[1]
                crit_code = row[2]
                frecuencia = row[3]
                nota = row[4]
                if frecuencia == "Trimestral":
                    habitos_prev_trim_lookup[(est_id_db, t_num, crit_code)] = nota
                else:
                    if (est_id_db, t_num, crit_code) not in habitos_prev_cont_lookup:
                        habitos_prev_cont_lookup[(est_id_db, t_num, crit_code)] = []
                    habitos_prev_cont_lookup[(est_id_db, t_num, crit_code)].append(nota)

        for est in estudiantes:
            id_est_str = str(est["id"])
            est_id_key = est["id"]
            
            # Criterios del trimestre actual
            consolidado_actual = {}
            for crit in criterios:
                crit_name = crit[0]
                
                # Intentar cargar evaluación Trimestral
                nota_trim = habitos_trim_lookup.get((est_id_key, crit_name))
                
                # Cargar evaluaciones continuas
                other_notes = [n for n in habitos_cont_lookup.get((est_id_key, crit_name), []) if n in ('S', 'R', 'X')]
                
                s_c = other_notes.count("S")
                r_c = other_notes.count("R")
                x_c = other_notes.count("X")
                
                final_val = "-"
                if nota_trim and nota_trim in ('S', 'R', 'X'):
                    final_val = nota_trim
                    if final_val == "S": s_c += 1
                    elif final_val == "R": r_c += 1
                    elif final_val == "X": x_c += 1
                elif other_notes:
                    counts = {"S": s_c, "R": r_c, "X": x_c}
                    for r in ["X", "R", "S"]:
                        if counts[r] > 0 and counts[r] >= max(counts.values()):
                            final_val = r
                            break
                            
                stats_habitos["s_count"] += s_c
                stats_habitos["r_count"] += r_c
                stats_habitos["x_count"] += x_c
                stats_habitos["total_evaluaciones"] += (1 if nota_trim else 0) + len(other_notes)
                
                consolidado_actual[crit_name] = {
                    "final": final_val,
                    "detalle": f"S:{s_c} R:{r_c} X:{x_c}" if len(other_notes) > 1 else ""
                }
                
            # Trimestres anteriores
            consolidado_anteriores = {}
            for t_num in range(1, num_trim_actual):
                t_name = f"Trimestre {t_num}"
                consolidado_anteriores[t_name] = {}
                for crit in criterios:
                    crit_name = crit[0]
                    
                    nota_trim_prev = habitos_prev_trim_lookup.get((est_id_key, t_num, crit_name))
                    
                    if nota_trim_prev and nota_trim_prev in ('S', 'R', 'X'):
                        consolidado_anteriores[t_name][crit_name] = nota_trim_prev
                    else:
                        other_notes_prev = [n for n in habitos_prev_cont_lookup.get((est_id_key, t_num, crit_name), []) if n in ('S', 'R', 'X')]
                        if other_notes_prev:
                            counts = {"S": other_notes_prev.count("S"), "R": other_notes_prev.count("R"), "X": other_notes_prev.count("X")}
                            res = "-"
                            for r in ["X", "R", "S"]:
                                if counts[r] > 0 and counts[r] >= max(counts.values()):
                                    res = r
                                    break
                            consolidado_anteriores[t_name][crit_name] = res
                        else:
                            consolidado_anteriores[t_name][crit_name] = "-"
                            
            rows_habitos.append({
                "id": est["id"],
                "nombre": est["nombre"],
                "evals_actual": consolidado_actual,
                "evals_anteriores": consolidado_anteriores
            })
            
        return {
            "headers_notas": headers_notas,
            "rows_notas": rows_notas,
            "stats_notas": stats_notas,
            "headers_asis": headers_asis,
            "rows_asis": rows_asis,
            "stats_asis": stats_asis,
            "habitos_criterios": [crit[0] for crit in criterios],
            "rows_habitos": rows_habitos,
            "stats_habitos": stats_habitos
        }

    def _renderizar_datos(self, data):
        self.loading = False
        self._limpiar_scroll(self.scroll_notas)
        self._limpiar_scroll(self.scroll_asis)
        self._limpiar_scroll(self.scroll_habitos)

        if not data:
            self._renderizar_error()
            return

        # ─── 3. HÁBITOS ───
        self._renderizar_stats_habitos(data["stats_habitos"])
        self._renderizar_habitos(data["rows_habitos"], data["habitos_criterios"])

        # ─── 1. NOTAS ───
        self._renderizar_stats_notas(data["stats_notas"])
        h_notas = data["headers_notas"]
        r_notas = data["rows_notas"]

        if not h_notas:
            ctk.CTkLabel(self.scroll_notas, text="No hay calificaciones registradas para esta materia y trimestre.",
                         font=(FONT_BODY, 13), text_color=C["texto_sec"]).pack(pady=40)
        else:
            # Encabezados
            f_header = ctk.CTkFrame(self.scroll_notas, fg_color=C["badge_bg"], corner_radius=5)
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
                f_row, widgets = self._obtener_fila_reciclada(self.scroll_notas, i, bg)

                widgets["num"].configure(text=f"{i+1}")
                widgets["nombre"].configure(text=row_data['nombre'])

                for lbl in widgets["valores"]:
                    lbl.pack_forget()
                while len(widgets["valores"]) < len(row_data["notas"]):
                    lbl_val = ctk.CTkLabel(f_row, text="", width=65, font=(FONT_BODY, 11))
                    widgets["valores"].append(lbl_val)

                for j, nota in enumerate(row_data["notas"]):
                    lbl_val = widgets["valores"][j]
                    color = C["texto"]
                    if isinstance(nota, (int, float)):
                        if nota < 3.0:
                            color = C["rojo"]
                        elif nota >= 4.5:
                            color = C["verde"]
                    lbl_val.configure(text=f"{nota}" if nota != "" else "-", text_color=color)
                    lbl_val.pack(side="left", padx=3)

                prom = row_data["promedio"]
                prom_str = f"{prom:.2f}" if isinstance(prom, (int, float)) else (f"{prom}" if prom else "-")
                color_prom = C["texto"]
                if isinstance(prom, (int, float)):
                    if prom < 3.0:
                        color_prom = C["rojo"]
                    elif prom >= 4.0:
                        color_prom = C["verde"]
                
                widgets["resumen"].configure(text=prom_str, text_color=color_prom)
            self._limpiar_filas_excedentes(self.scroll_notas, len(r_notas))

        # ─── 2. ASISTENCIA ───
        self._renderizar_stats_asis(data["stats_asis"])
        h_asis = data["headers_asis"]
        r_asis = data["rows_asis"]

        if not h_asis:
            ctk.CTkLabel(self.scroll_asis, text="No hay registros de asistencia para este grado y trimestre.",
                         font=(FONT_BODY, 13), text_color=C["texto_sec"]).pack(pady=40)
        else:
            # Encabezados
            f_header = ctk.CTkFrame(self.scroll_asis, fg_color=C["badge_bg"], corner_radius=5)
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
                f_row, widgets = self._obtener_fila_reciclada(self.scroll_asis, i, bg)

                widgets["num"].configure(text=f"{i+1}")
                widgets["nombre"].configure(text=row_data['nombre'])

                for lbl in widgets["valores"]:
                    lbl.pack_forget()
                while len(widgets["valores"]) < len(row_data["asistencia"]):
                    lbl_val = ctk.CTkLabel(f_row, text="", width=65, font=(FONT_BODY, 11))
                    widgets["valores"].append(lbl_val)

                for j, st in enumerate(row_data["asistencia"]):
                    lbl_val = widgets["valores"][j]
                    color = C["texto"]
                    if st == "A":
                        color = C["rojo"]
                    elif st == "T":
                        color = C["amarillo"]
                    elif st == "E":
                        color = C["acento2"]
                    elif st == "P":
                        color = C["verde"]
                    lbl_val.configure(text=st, text_color=color)
                    lbl_val.pack(side="left", padx=3)

                pct = row_data["porcentaje"]
                pct_str = f"{pct}%" if row_data["dias"] > 0 else "-"
                color_pct = C["verde"] if pct >= 90 else (C["amarillo"] if pct >= 80 else C["rojo"])
                
                widgets["resumen"].configure(text=pct_str, text_color=color_pct)
            self._limpiar_filas_excedentes(self.scroll_asis, len(r_asis))

    def _renderizar_stats_notas(self, stats):
        # Actualizar valores de los cards existentes si ya fueron creados (evita recrear widgets)
        if hasattr(self, "_stats_notas_labels") and self._stats_notas_labels:
            vals = [str(stats['actividades']),
                    f"{stats['promedio_grupal']:.2f}",
                    f"{stats['tasa_aprobacion']}%",
                    f"{stats['max_nota']:.1f}"]
            for lbl, val in zip(self._stats_notas_labels, vals):
                try:
                    lbl.configure(text=val)
                except Exception:
                    pass
            return

        for w in self.stats_notas_frame.winfo_children():
            w.destroy()

        # Grid layout for cards
        self.stats_notas_frame.grid_columnconfigure((0, 1, 2, 3), weight=1)
        self._stats_notas_labels = []

        cards = [
            ("📝 Actividades",   f"{stats['actividades']}",        "Evaluaciones Tomadas",    C["cian"]),
            ("📊 Promedio Salón",f"{stats['promedio_grupal']:.2f}","Promedio General",         C["amarillo"]),
            ("🎓 Aprobación",    f"{stats['tasa_aprobacion']}%",   "Estudiantes con >= 3.0",  C["verde"]),
            ("🏆 Nota Máxima",   f"{stats['max_nota']:.1f}",       "Rendimiento Superior",    C["purpura"])
        ]

        for idx, (title, val, desc, color) in enumerate(cards):
            card = ctk.CTkFrame(self.stats_notas_frame, fg_color=C["card"], corner_radius=8,
                                border_color=color, border_width=1)
            card.grid(row=0, column=idx, padx=5, pady=5, sticky="nsew")
            ctk.CTkLabel(card, text=title, font=(FONT_BODY, 12, "bold"), text_color=color).pack(anchor="w", padx=15, pady=(10, 2))
            lbl_val = ctk.CTkLabel(card, text=val, font=(FONT_TITLE, 22, "bold"), text_color=C["texto"])
            lbl_val.pack(anchor="w", padx=15, pady=2)
            ctk.CTkLabel(card, text=desc, font=(FONT_BODY, 10), text_color=C["texto_sec"]).pack(anchor="w", padx=15, pady=(2, 10))
            self._stats_notas_labels.append(lbl_val)

    def _renderizar_stats_asis(self, stats):
        # Actualizar valores de los cards existentes si ya fueron creados
        if hasattr(self, "_stats_asis_labels") and self._stats_asis_labels:
            vals = [str(stats['total_dias']),
                    f"{stats['asistencia_promedio']}%",
                    str(stats['asistencia_perfecta']),
                    str(stats['alerta_count'])]
            for lbl, val in zip(self._stats_asis_labels, vals):
                try:
                    lbl.configure(text=val)
                except Exception:
                    pass
            return

        for w in self.stats_asis_frame.winfo_children():
            w.destroy()

        self.stats_asis_frame.grid_columnconfigure((0, 1, 2, 3), weight=1)
        self._stats_asis_labels = []

        cards = [
            ("📅 Días Tomados",      f"{stats['total_dias']}",              "Fechas de Asistencia",           C["cian"]),
            ("📈 Asistencia Prom.",  f"{stats['asistencia_promedio']}%",    "Asistencia del Grupo",           C["verde"]),
            ("⭐ Perfecta (100%)",   f"{stats['asistencia_perfecta']}",     "Asistencia Inmaculada",          C["amarillo"]),
            ("⚠️ En Alerta (<90%)", f"{stats['alerta_count']}",            "Estudiantes con Inasistencias",  C["rojo"])
        ]

        for idx, (title, val, desc, color) in enumerate(cards):
            card = ctk.CTkFrame(self.stats_asis_frame, fg_color=C["card"], corner_radius=8,
                                border_color=color, border_width=1)
            card.grid(row=0, column=idx, padx=5, pady=5, sticky="nsew")
            ctk.CTkLabel(card, text=title, font=(FONT_BODY, 12, "bold"), text_color=color).pack(anchor="w", padx=15, pady=(10, 2))
            lbl_val = ctk.CTkLabel(card, text=val, font=(FONT_TITLE, 22, "bold"), text_color=C["texto"])
            lbl_val.pack(anchor="w", padx=15, pady=2)
            ctk.CTkLabel(card, text=desc, font=(FONT_BODY, 10), text_color=C["texto_sec"]).pack(anchor="w", padx=15, pady=(2, 10))
            self._stats_asis_labels.append(lbl_val)

    def _renderizar_error(self):
        self.loading = False
        self._limpiar_scroll(self.scroll_notas)
        self._limpiar_scroll(self.scroll_asis)
        self._limpiar_scroll(self.scroll_habitos)
        
        ctk.CTkLabel(self.scroll_notas, text="⚠️ Ocurrió un error o no hay hoja de notas válida en el Excel para este grupo.",
                     font=(FONT_BODY, 13), text_color=C["rojo"]).pack(pady=40)
        ctk.CTkLabel(self.scroll_asis, text="⚠️ Ocurrió un error o no hay hoja de asistencia válida en el Excel para este grupo.",
                     font=(FONT_BODY, 13), text_color=C["rojo"]).pack(pady=40)
        ctk.CTkLabel(self.scroll_habitos, text="⚠️ Ocurrió un error o no hay evaluaciones de hábitos registradas.",
                     font=(FONT_BODY, 13), text_color=C["rojo"]).pack(pady=40)

    def _limpiar_scroll(self, scroll):
        pool_frames = set()
        if hasattr(scroll, "_pool_filas"):
            for f_row, _ in scroll._pool_filas:
                pool_frames.add(f_row)
                f_row.pack_forget()
        for w in scroll.winfo_children():
            if w not in pool_frames:
                w.destroy()

    def _obtener_fila_reciclada(self, scroll, index, fg_color):
        if not hasattr(scroll, "_pool_filas"):
            scroll._pool_filas = []
        if index < len(scroll._pool_filas):
            f_row, widgets = scroll._pool_filas[index]
            f_row.configure(fg_color=fg_color)
            f_row.pack(fill="x", padx=5, pady=1)
            return f_row, widgets
        f_row = ctk.CTkFrame(scroll, fg_color=fg_color, corner_radius=3)
        f_row.pack(fill="x", padx=5, pady=1)
        widgets = {
            "num": ctk.CTkLabel(f_row, text="", width=35, font=(FONT_BODY, 11)),
            "nombre": ctk.CTkLabel(f_row, text="", width=180, anchor="w", font=(FONT_BODY, 11)),
            "valores": [],
            "resumen": ctk.CTkLabel(f_row, text="", width=80, font=(FONT_BODY, 11, "bold")),
        }
        widgets["num"].pack(side="left", padx=5)
        widgets["nombre"].pack(side="left", padx=5)
        widgets["resumen"].pack(side="right", padx=10)
        scroll._pool_filas.append((f_row, widgets))
        return f_row, widgets

    def _obtener_fila_habito_reciclada(self, scroll, index, fg_color):
        if not hasattr(scroll, "_pool_filas"):
            scroll._pool_filas = []
        if index < len(scroll._pool_filas):
            f_row, widgets = scroll._pool_filas[index]
            f_row.configure(fg_color=fg_color)
            f_row.pack(fill="x", padx=5, pady=4)
            return f_row, widgets
        f_row = ctk.CTkFrame(scroll, fg_color=fg_color, corner_radius=8)
        f_row.pack(fill="x", padx=5, pady=4)
        header = ctk.CTkFrame(f_row, fg_color="transparent")
        header.pack(fill="x", padx=15, pady=(8, 4))
        lbl_num = ctk.CTkLabel(header, text="", font=(FONT_BODY, 12, "bold"), text_color=C["texto"])
        lbl_num.pack(side="left")
        lbl_name = ctk.CTkLabel(header, text="", font=(FONT_BODY, 12, "bold"), text_color=C["texto"])
        lbl_name.pack(side="left", padx=10)
        grid_frame = ctk.CTkFrame(f_row, fg_color="transparent")
        grid_frame.pack(fill="x", padx=15, pady=(2, 6))
        ant_frame = ctk.CTkFrame(f_row, fg_color=C["fondo"], corner_radius=6)
        widgets = {
            "num": lbl_num,
            "nombre": lbl_name,
            "grid_frame": grid_frame,
            "item_widgets": [],
            "ant_frame": ant_frame,
            "ant_labels": [],
            "ant_title": ctk.CTkLabel(ant_frame, text="⏮️ Trimestres Anteriores:", font=(FONT_BODY, 10, "bold"), text_color=C["cian"])
        }
        widgets["ant_title"].pack(anchor="w", padx=10, pady=(4, 2))
        scroll._pool_filas.append((f_row, widgets))
        return f_row, widgets

    def _limpiar_filas_excedentes(self, scroll, num_requeridos):
        if not hasattr(scroll, "_pool_filas"):
            return
        for idx in range(num_requeridos, len(scroll._pool_filas)):
            f_row, _ = scroll._pool_filas[idx]
            f_row.pack_forget()

    def _renderizar_stats_habitos(self, stats):
        for w in self.stats_habitos_frame.winfo_children():
            w.destroy()
            
        self.stats_habitos_frame.grid_columnconfigure((0, 1, 2, 3), weight=1)
        
        cards = [
            ("🧠 Total Evals", f"{stats['total_evaluaciones']}", "Evaluaciones Realizadas", C["cian"]),
            ("🟢 Satisfactorios (S)", f"{stats['s_count']}", "Desempeño Satisfactorio", C["verde"]),
            ("🟡 Regulares (R)", f"{stats['r_count']}", "Desempeño Regular", C["amarillo"]),
            ("🔴 No Satisface (X)", f"{stats['x_count']}", "No Satisface", C["rojo"])
        ]
        
        for idx, (title, val, desc, color) in enumerate(cards):
            card = ctk.CTkFrame(self.stats_habitos_frame, fg_color=C["card"], corner_radius=8, border_color=color, border_width=1)
            card.grid(row=0, column=idx, padx=5, pady=5, sticky="nsew")
            
            ctk.CTkLabel(card, text=title, font=(FONT_BODY, 12, "bold"), text_color=color).pack(anchor="w", padx=15, pady=(10, 2))
            ctk.CTkLabel(card, text=val, font=(FONT_TITLE, 22, "bold"), text_color=C["texto"]).pack(anchor="w", padx=15, pady=2)
            ctk.CTkLabel(card, text=desc, font=(FONT_BODY, 10), text_color=C["texto_sec"]).pack(anchor="w", padx=15, pady=(2, 10))

    def _renderizar_habitos(self, rows_habitos, habitos_criterios):
        if not rows_habitos:
            ctk.CTkLabel(self.scroll_habitos, text="No hay registros de hábitos para este grado y trimestre.",
                         font=(FONT_BODY, 13), text_color=C["texto_sec"]).pack(pady=40)
            return

        for i, row_data in enumerate(rows_habitos):
            bg = C["card"] if i % 2 == 0 else C["card_alt"]
            f_row, widgets = self._obtener_fila_habito_reciclada(self.scroll_habitos, i, bg)

            widgets["num"].configure(text=f"{i+1}.")
            widgets["nombre"].configure(text=row_data['nombre'])

            grid_frame = widgets["grid_frame"]
            grid_frame.columnconfigure((0, 1, 2), weight=1)

            # Ocultar todos los frames de ítems actuales
            for iw in widgets["item_widgets"]:
                iw["frame"].grid_forget()

            # Configurar y empaquetar los necesarios
            for idx, crit_name in enumerate(habitos_criterios):
                c_row = idx // 3
                c_col = idx % 3

                while len(widgets["item_widgets"]) <= idx:
                    item_frame = ctk.CTkFrame(grid_frame, fg_color="transparent")
                    badge = ctk.CTkLabel(item_frame, text="", font=(FONT_BODY, 10, "bold"), text_color="white", corner_radius=4)
                    badge.pack(side="left", padx=(0, 6))
                    lbl_n = ctk.CTkLabel(item_frame, text="", font=(FONT_BODY, 11), text_color=C["texto"])
                    lbl_n.pack(side="left")
                    lbl_det = ctk.CTkLabel(item_frame, text="", font=(FONT_BODY, 9), text_color=C["texto_dim"])
                    widgets["item_widgets"].append({
                        "frame": item_frame,
                        "badge": badge,
                        "name": lbl_n,
                        "det": lbl_det
                    })

                iw = widgets["item_widgets"][idx]
                iw["frame"].grid(row=c_row, column=c_col, sticky="w", padx=10, pady=4)

                crit_data = row_data["evals_actual"].get(crit_name, {"final": "-", "detalle": ""})
                final_val = crit_data["final"]
                detalle_val = crit_data["detalle"]

                color = C["texto_sec"]
                if final_val == "S": color = C["verde"]
                elif final_val == "R": color = C["amarillo"]
                elif final_val == "X": color = C["rojo"]

                iw["badge"].configure(text=f" {final_val} ", fg_color=color)
                iw["name"].configure(text=crit_name)

                iw["det"].pack_forget()
                if detalle_val:
                    iw["det"].configure(text=f" ({detalle_val})")
                    iw["det"].pack(side="left", padx=2)

            # Trimestres anteriores
            anteriores_texts = []
            for prev_t, prev_crit_vals in row_data["evals_anteriores"].items():
                active_crits = []
                for c_name, val in prev_crit_vals.items():
                    if val != "-":
                        active_crits.append(f"{c_name}: {val}")
                if active_crits:
                    anteriores_texts.append(f"⏱️ {prev_t}: " + " | ".join(active_crits))

            ant_frame = widgets["ant_frame"]
            ant_frame.pack_forget()
            for lbl in widgets["ant_labels"]:
                lbl.pack_forget()

            if anteriores_texts:
                ant_frame.pack(fill="x", padx=15, pady=(4, 10))
                for k, txt in enumerate(anteriores_texts):
                    while len(widgets["ant_labels"]) <= k:
                        lbl = ctk.CTkLabel(ant_frame, text="", font=(FONT_BODY, 10), text_color=C["texto_sec"], justify="left", wraplength=850)
                        widgets["ant_labels"].append(lbl)
                    lbl = widgets["ant_labels"][k]
                    lbl.configure(text=txt)
                    lbl.pack(anchor="w", padx=20, pady=2)
                    
        self._limpiar_filas_excedentes(self.scroll_habitos, len(rows_habitos))
