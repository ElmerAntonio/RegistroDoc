import os
import json
import datetime
import tkinter as tk
import threading
from tkinter import messagebox
import customtkinter as ctk
from config import BASE_DIR

try:
    from docx import Document
    from docx.shared import Pt, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    DOCX_DISPONIBLE = True
except ImportError:
    DOCX_DISPONIBLE = False

# Criterios oficiales de Panamá (MEDUCA)
CRITERIOS_PRIMARIA = [
    ("Responsabilidad", "Asistencia a clases y entrega puntual de tareas."),
    ("Orden y Aseo", "Cuidado de útiles, libretas y aseo personal."),
    ("Organización del Trabajo", "Planificación y estructura al hacer deberes."),
    ("Autodominio y Confianza en sí mismo", "Control emocional, seguridad y autonomía."),
    ("Iniciativa", "Proactividad, participación y propuestas nuevas."),
    ("Cooperación", "Trabajo en equipo y compañerismo en el aula."),
    ("Respeto a la propiedad ajena", "Cuidado de bienes del aula y de compañeros.")
]

CRITERIOS_PREMEDIA_MEDIA = [
    ("Responsabilidad", "Cumplimiento y compromiso con deberes y tareas."),
    ("Puntualidad", "Llegada a tiempo a clase y entrega de trabajos."),
    ("Honradez", "Honestidad académica (no copiar, no mentir) y rectitud."),
    ("Conciencia Cívica", "Respeto a símbolos patrios, normas y valores cívicos."),
    ("Organización del Trabajo", "Método ordenado y planificación de actividades."),
    ("Autodominio y Confianza en sí mismo", "Autocontrol emocional frente a retos y autoestima."),
    ("Iniciativa", "Proactividad para aprender y emprender tareas."),
    ("Cooperación", "Empatía, compañerismo y trabajo en grupo."),
    ("Respeto a la propiedad ajena", "Cuidado de materiales ajenos y del centro."),
    ("Modales", "Cortesía, vocabulario correcto y buen trato."),
    ("Orden y Aseo", "Limpieza de trabajos y excelente presentación personal."),
    ("Empleo del tiempo libre", "Uso productivo y sano de recesos y tiempo libre.")
]

class HabitosFrame(ctk.CTkFrame):
    def __init__(self, master, engine, **kwargs):
        super().__init__(master, fg_color="transparent", **kwargs)
        self.engine = engine

        # Estado interno
        self.estudiantes = []
        self.evaluaciones_temporales = {}  # id_estudiante -> {criterio: valor}
        self.combo_widgets = {}
        self.criterios_activos = []

        # Ruta de datos JSON
        self.ruta_json = os.path.join(BASE_DIR, "..", "Expedientes_Estudiantes", "habitos_evaluaciones.json")

        self.crear_interfaz()
        self.al_cambiar_grado_trimestre()

    def crear_interfaz(self):
        # 1. Contenedor Superior (Filtros y Controles)
        self.frame_filtros = ctk.CTkFrame(self, fg_color="#1a2436", corner_radius=10)
        self.frame_filtros.pack(fill="x", padx=10, pady=10)

        # Grado
        ctk.CTkLabel(self.frame_filtros, text="Grado:", font=("Outfit", 12, "bold")).grid(row=0, column=0, padx=10, pady=10, sticky="w")
        self.combo_grado = ctk.CTkComboBox(
            self.frame_filtros, 
            values=self.engine.obtener_grados_activos() or ["Sin grados"], 
            command=self.al_cambiar_grado_trimestre,
            width=140
        )
        self.combo_grado.grid(row=0, column=1, padx=5, pady=10)

        # Trimestre
        ctk.CTkLabel(self.frame_filtros, text="Trimestre:", font=("Outfit", 12, "bold")).grid(row=0, column=2, padx=10, pady=10, sticky="w")
        self.combo_trimestre = ctk.CTkComboBox(
            self.frame_filtros, 
            values=["Trimestre 1", "Trimestre 2", "Trimestre 3"], 
            command=self.al_cambiar_grado_trimestre,
            width=140
        )
        self.combo_trimestre.grid(row=0, column=3, padx=5, pady=10)

        # Botón Autorellenar con IA
        self.btn_sugerir = ctk.CTkButton(
            self.frame_filtros,
            text="🧠 Autorellenar con IA",
            font=("Outfit", 12, "bold"),
            fg_color="#3B82F6",
            hover_color="#2563EB",
            command=self.autorellenar_con_ia,
            width=160
        )
        self.btn_sugerir.grid(row=0, column=4, padx=15, pady=10)

        # Botón Guardar
        self.btn_guardar_todo = ctk.CTkButton(
            self.frame_filtros,
            text="💾 GUARDAR EVALUACIONES",
            font=("Outfit", 12, "bold"),
            fg_color="#10B981",
            hover_color="#059669",
            command=self.guardar_evaluaciones,
            width=190
        )
        self.btn_guardar_todo.grid(row=0, column=5, padx=5, pady=10)

        # 2. Contenedor de la Tabla Matricial (Scrollable)
        self.scroll_tabla = ctk.CTkScrollableFrame(self, fg_color="#121c2c", corner_radius=10)
        self.scroll_tabla.pack(fill="both", expand=True, padx=10, pady=5)

    def al_cambiar_grado_trimestre(self, *args):
        grado = self.combo_grado.get()
        if not grado or grado == "Sin grados":
            return
        self.cargar_tabla_matricial()

    def cargar_tabla_matricial(self):
        # Limpiar tabla anterior
        for widget in self.scroll_tabla.winfo_children():
            try:
                widget.destroy()
            except Exception:
                pass

        grado = self.combo_grado.get()
        trimestre = self.combo_trimestre.get()
        if not grado or grado == "Sin grados":
            return

        # Cargar estudiantes
        self.estudiantes = self.engine.obtener_estudiantes_completos(grado)
        if not self.estudiantes or "Sin estudiantes" in self.estudiantes[0]['nombre']:
            lbl = ctk.CTkLabel(self.scroll_tabla, text="No hay estudiantes inscritos en este grado.", font=("Outfit", 14, "italic"), text_color="#A3A3A3")
            lbl.pack(pady=40)
            return

        # Determinar criterios según modalidad
        es_primaria = any(g in grado for g in ["1°", "2°", "3°", "4°", "5°", "6°"])
        criterios_info = CRITERIOS_PRIMARIA if es_primaria else CRITERIOS_PREMEDIA_MEDIA
        self.criterios_activos = [crit[0] for crit in criterios_info]

        # Cargar evaluaciones guardadas desde SQLite (Tarea 12 y 19)
        trimestre_num = int(trimestre.replace("Trimestre ", ""))
        try:
            cursor = self.engine.db_conn.cursor()
            est_ids = [str(e['id']) for e in self.estudiantes]
            placeholders = ",".join(["?"] * len(est_ids))
            
            cursor.execute(f"""
                SELECT estudiante_id, criterio_codigo, nota 
                FROM habitos 
                WHERE estudiante_id IN ({placeholders}) AND trimestre = ?;
            """, est_ids + [trimestre_num])
            rows = cursor.fetchall()
            
            # Inicializar self.evaluaciones_temporales
            self.evaluaciones_temporales = {}
            for est in self.estudiantes:
                self.evaluaciones_temporales[str(est['id'])] = {crit: "-" for crit in self.criterios_activos}
                
            for est_id, crit_code, nota in rows:
                if est_id in self.evaluaciones_temporales and crit_code in self.evaluaciones_temporales[est_id]:
                    self.evaluaciones_temporales[est_id][crit_code] = nota
        except Exception as e:
            print(f"[!] Error cargando hábitos de SQLite: {e}")
            self.evaluaciones_temporales = {}
            for est in self.estudiantes:
                self.evaluaciones_temporales[str(est['id'])] = {crit: "-" for crit in self.criterios_activos}

        # Configurar columnas de grid
        self.scroll_tabla.grid_columnconfigure(0, minsize=220)  # Nombre Estudiante
        for c_idx in range(1, len(self.criterios_activos) + 1):
            self.scroll_tabla.grid_columnconfigure(c_idx, minsize=90)

        # ─── Fila 0: Encabezados de Criterios Completos ───
        lbl_est = ctk.CTkLabel(self.scroll_tabla, text="Estudiante", font=("Outfit", 12, "bold"), text_color="#3B82F6", anchor="w")
        lbl_est.grid(row=0, column=0, padx=10, pady=10, sticky="w")

        for c_idx, crit in enumerate(self.criterios_activos, 1):
            lbl_crit = ctk.CTkLabel(self.scroll_tabla, text=crit, font=("Outfit", 10, "bold"), text_color="#22D3EE", wraplength=85, justify="center")
            lbl_crit.grid(row=0, column=c_idx, padx=5, pady=10)

        # ─── Filas de Estudiantes ───
        self.combo_widgets = {}
        for r_idx, est in enumerate(self.estudiantes, 1):
            est_id = str(est['id'])
            
            # Nombre de estudiante
            lbl_nom = ctk.CTkLabel(self.scroll_tabla, text=est['nombre'], font=("Outfit", 11, "bold"), text_color="#F3F4F6", anchor="w")
            lbl_nom.grid(row=r_idx, column=0, padx=10, pady=4, sticky="w")
            
            # Dropdowns para los criterios
            for c_idx, crit in enumerate(self.criterios_activos, 1):
                val_actual = self.evaluaciones_temporales[est_id].get(crit, "-")
                
                def get_color(val):
                    return "#10B981" if val == "S" else ("#F59E0B" if val == "R" else ("#EF4444" if val == "X" else "#4B5563"))
                
                opt_menu = ctk.CTkOptionMenu(
                    self.scroll_tabla,
                    values=["-", "S", "R", "X"],
                    width=65,
                    height=24,
                    font=("Outfit", 10, "bold"),
                    fg_color=get_color(val_actual),
                    button_color=get_color(val_actual),
                    button_hover_color="#374151",
                    command=lambda val, eid=est_id, cn=crit: self.cambiar_valor_celda(eid, cn, val)
                )
                opt_menu.set(val_actual)
                opt_menu.grid(row=r_idx, column=c_idx, padx=5, pady=4)
                
                self.combo_widgets[(est_id, crit)] = opt_menu

    def cambiar_valor_celda(self, est_id, criterio, valor):
        self.evaluaciones_temporales[est_id][criterio] = valor
        opt_menu = self.combo_widgets.get((est_id, criterio))
        if opt_menu:
            color = "#10B981" if valor == "S" else ("#F59E0B" if valor == "R" else ("#EF4444" if valor == "X" else "#4B5563"))
            opt_menu.configure(fg_color=color, button_color=color)

    def autorellenar_con_ia(self):
        grado = self.combo_grado.get()
        trimestre = self.combo_trimestre.get()
        if not self.estudiantes:
            return

        confirmar = messagebox.askyesno(
            "Autorellenar con IA",
            "¿Desea analizar el rendimiento, asistencia y disciplina de todos los estudiantes para rellenar las sugerencias automáticamente?\n\nEsto sobrescribirá las selecciones actuales no guardadas."
        )
        if not confirmar:
            return

        self.btn_sugerir.configure(text="⏳ ANALIZANDO...", state="disabled")
        self.update()

        try:
            for est in self.estudiantes:
                id_est = est['id']
                est_id_str = str(id_est)
                
                stats_asistencia = self.engine.obtener_estadisticas_asistencia(grado, trimestre, id_est)
                stats_notas = self.engine.obtener_tareas_sin_nota(grado, trimestre, id_est)
                
                reportes_disciplina = []
                carpeta_exp = os.path.join(BASE_DIR, "..", "Expedientes_Estudiantes")
                nombre_archivo = f"{est['nombre']} - {grado.replace('°', '')}.docx".replace("/", "-")
                ruta_word = os.path.join(carpeta_exp, nombre_archivo)
                if os.path.exists(ruta_word) and DOCX_DISPONIBLE:
                    try:
                        doc = Document(ruta_word)
                        if doc.tables:
                            tabla = doc.tables[0]
                            for row in tabla.rows[1:]:
                                if len(row.cells) >= 4:
                                    tipo = row.cells[2].text.strip()
                                    motivo = row.cells[3].text.strip()
                                    if any(k in tipo for k in ["Conducta", "Citación", "Reporte", "Disciplina"]):
                                        reportes_disciplina.append(motivo)
                    except Exception:
                        pass
                
                sugerencias = {}
                aus_sin_just = stats_asistencia.get("ausencias", 0)
                tardanzas = stats_asistencia.get("tardanzas", 0)
                tareas_vacias = stats_notas.get("total_vacias", 0)
                
                # Responsabilidad
                resp_val = "S"
                if aus_sin_just >= 4 or tareas_vacias >= 3:
                    resp_val = "X"
                elif aus_sin_just >= 2 or tareas_vacias >= 1:
                    resp_val = "R"
                sugerencias["Responsabilidad"] = resp_val
                
                # Puntualidad
                if "Puntualidad" in self.criterios_activos:
                    punt_val = "S"
                    if tardanzas >= 6:
                        punt_val = "X"
                    elif tardanzas >= 3:
                        punt_val = "R"
                    sugerencias["Puntualidad"] = punt_val
                
                # Organización del Trabajo
                org_val = "S"
                if tareas_vacias >= 3:
                    org_val = "X"
                elif tareas_vacias >= 1:
                    org_val = "R"
                sugerencias["Organización del Trabajo"] = org_val
                
                # Conducta, autodominio, cooperación, etc.
                con_val = "S"
                coop_val = "S"
                mod_val = "S"
                hon_val = "S"
                
                if reportes_disciplina:
                    con_val = "R"
                    coop_val = "R"
                    mod_val = "R"
                    hon_val = "R"
                    if any("Grave" in r or "Suspensión" in r for r in reportes_disciplina):
                        con_val = "X"
                        coop_val = "X"
                        mod_val = "X"
                
                sugerencias["Autodominio y Confianza en sí mismo"] = con_val
                sugerencias["Cooperación"] = coop_val
                if "Modales" in self.criterios_activos:
                    sugerencias["Modales"] = mod_val
                if "Honradez" in self.criterios_activos:
                    sugerencias["Honradez"] = hon_val
                
                for crit in self.criterios_activos:
                    if crit not in sugerencias:
                        sugerencias[crit] = "S"
                
                for crit, val in sugerencias.items():
                    self.evaluaciones_temporales[est_id_str][crit] = val
                    opt_menu = self.combo_widgets.get((est_id_str, crit))
                    if opt_menu:
                        opt_menu.set(val)
                        color = "#10B981" if val == "S" else ("#F59E0B" if val == "R" else ("#EF4444" if val == "X" else "#4B5563"))
                        opt_menu.configure(fg_color=color, button_color=color)

            root = self.winfo_toplevel()
            if hasattr(root, "mostrar_toast"):
                root.mostrar_toast("✓ Autocompletado IA finalizado", color="#10B981")
        except Exception as e:
            messagebox.showerror("Error", f"Error en el análisis de IA: {e}")
        finally:
            self.btn_sugerir.configure(text="🧠 Autorellenar con IA", state="normal")

    def guardar_evaluaciones(self):
        grado = self.combo_grado.get()
        trimestre = self.combo_trimestre.get()
        trimestre_num = int(trimestre.replace("Trimestre ", ""))

        if not self.evaluaciones_temporales:
            messagebox.showwarning("Atención", "No hay datos de evaluación para guardar.")
            return

        self.btn_guardar_todo.configure(text="⏳ GUARDANDO...", state="disabled")
        self.btn_sugerir.configure(state="disabled")
        self.update()

        def tarea_fondo():
            try:
                cursor = self.engine.db_conn.cursor()
                # 1. Guardar en SQLite (Base de datos principal)
                for est_id, crit_vals in self.evaluaciones_temporales.items():
                    for crit_code, nota in crit_vals.items():
                        cursor.execute("""
                            SELECT id FROM habitos 
                            WHERE estudiante_id = ? AND trimestre = ? AND criterio_codigo = ?;
                        """, (est_id, trimestre_num, crit_code))
                        row = cursor.fetchone()
                        if row:
                            cursor.execute("""
                                UPDATE habitos SET nota = ? WHERE id = ?;
                            """, (nota, row[0]))
                        else:
                            cursor.execute("""
                                INSERT INTO habitos (estudiante_id, trimestre, criterio_codigo, nota) 
                                VALUES (?, ?, ?, ?);
                            """, (est_id, trimestre_num, crit_code, nota))
                
                self.engine.db_conn.commit()
                self.engine.db_manager.guardar_cifrado()

                # 2. Guardar en JSON/Word (Expedientes locales) para compatibilidad
                datos_json = {}
                if os.path.exists(self.ruta_json):
                    try:
                        with open(self.ruta_json, "r", encoding="utf-8") as f:
                            datos_json = json.load(f)
                    except Exception:
                        pass
                
                clave_json = f"{grado}::{trimestre}::Trimester-Grid::Hoy"
                datos_json[clave_json] = {
                    "frecuencia": "Trimester-Grid",
                    "periodo": "Hoy",
                    "estudiantes": {str(k): v for k, v in self.evaluaciones_temporales.items()}
                }
                
                carpeta_exp = os.path.dirname(self.ruta_json)
                if not os.path.exists(carpeta_exp):
                    os.makedirs(carpeta_exp)
                
                with open(self.ruta_json, "w", encoding="utf-8") as f:
                    json.dump(datos_json, f, ensure_ascii=False, indent=4)

                # Actualizar Word de cada estudiante si está disponible
                if DOCX_DISPONIBLE:
                    for est in self.estudiantes:
                        id_est = est['id']
                        id_est_str = str(id_est)
                        if id_est_str in self.evaluaciones_temporales:
                            nombre_archivo = f"{est['nombre']} - {grado.replace('°', '')}.docx".replace("/", "-")
                            ruta_word = os.path.join(carpeta_exp, nombre_archivo)
                            
                            if os.path.exists(ruta_word):
                                try:
                                    doc = Document(ruta_word)
                                    if doc.tables:
                                        tabla = doc.tables[0]
                                        crit_vals = self.evaluaciones_temporales[id_est_str]
                                        eval_str = ", ".join([f"{k}: {v}" for k, v in crit_vals.items()])
                                        
                                        tipo_registro = f"Hábitos ({trimestre})"
                                        fecha_act = datetime.date.today().strftime("%d-%m-%Y")
                                        
                                        fila_existente = None
                                        for row in tabla.rows[1:]:
                                            if len(row.cells) >= 4:
                                                t = row.cells[2].text.strip()
                                                if t == tipo_registro:
                                                    fila_existente = row
                                                    break
                                        
                                        if fila_existente:
                                            fila_existente.cells[1].text = fecha_act
                                            fila_existente.cells[3].text = eval_str
                                        else:
                                            num_reg = str(len(tabla.rows))
                                            fila = tabla.add_row()
                                            fila.cells[0].text = num_reg
                                            fila.cells[1].text = fecha_act
                                            fila.cells[2].text = tipo_registro
                                            fila.cells[3].text = eval_str
                                        
                                        doc.save(ruta_word)
                                except Exception as e:
                                    print(f"Error escribiendo en Word de {est['nombre']}: {e}")

                self.after(0, lambda: self.finalizar_guardado(True, ""))
            except Exception as e:
                self.after(0, lambda: self.finalizar_guardado(False, str(e)))

        threading.Thread(target=tarea_fondo, daemon=True).start()

    def finalizar_guardado(self, exito, error_msg):
        self.btn_guardar_todo.configure(text="💾 GUARDAR EVALUACIONES", state="normal")
        self.btn_sugerir.configure(state="normal")
        
        if not exito:
            messagebox.showerror("Error", f"No se pudo guardar la evaluación: {error_msg}")
            return
            
        root = self.winfo_toplevel()
        if hasattr(root, "mostrar_toast"):
            root.mostrar_toast("✓ Evaluaciones guardadas con éxito", color="#10B981")
        else:
            messagebox.showinfo("Éxito", "Evaluaciones registradas y guardadas correctamente.")

    def actualizar_vista(self):
        opciones = self.engine.obtener_grados_activos() or ["Sin grados"]
        old_sel = self.combo_grado.get()
        self.combo_grado.configure(values=opciones)
        if old_sel in opciones:
            self.combo_grado.set(old_sel)
        else:
            self.combo_grado.set(opciones[0])
        self.al_cambiar_grado_trimestre()
