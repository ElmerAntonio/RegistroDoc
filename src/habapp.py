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
        self.estudiante_seleccionado = None
        self.evaluaciones_temporales = {}  # id_estudiante -> {criterio: valor}
        self.frecuencia_bloqueada = None

        # Ruta de datos JSON
        self.ruta_json = os.path.join(BASE_DIR, "..", "Expedientes_Estudiantes", "habitos_evaluaciones.json")

        self.crear_interfaz()

    def crear_interfaz(self):
        # 1. Contenedor Superior (Filtros y Configuración)
        self.frame_filtros = ctk.CTkFrame(self, fg_color="#1a2436", corner_radius=10)
        self.frame_filtros.pack(fill="x", padx=10, pady=10)

        # Grado
        ctk.CTkLabel(self.frame_filtros, text="Grado:", font=("Outfit", 12, "bold")).grid(row=0, column=0, padx=10, pady=10, sticky="w")
        self.combo_grado = ctk.CTkComboBox(
            self.frame_filtros, 
            values=self.engine.obtener_grados_activos(), 
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

        # Frecuencia
        ctk.CTkLabel(self.frame_filtros, text="Frecuencia:", font=("Outfit", 12, "bold")).grid(row=0, column=4, padx=10, pady=10, sticky="w")
        self.btn_frecuencia = ctk.CTkSegmentedButton(
            self.frame_filtros,
            values=["Diario", "Semanal", "Mensual"],
            command=self.al_cambiar_frecuencia,
            width=240,
            selected_color="#3B82F6"
        )
        self.btn_frecuencia.set("Diario")
        self.btn_frecuencia.grid(row=0, column=5, padx=5, pady=10)

        # Selector de Fecha/Periodo
        self.lbl_periodo = ctk.CTkLabel(self.frame_filtros, text="Fecha:", font=("Outfit", 12, "bold"))
        self.lbl_periodo.grid(row=0, column=6, padx=10, pady=10, sticky="w")

        # Widget de entrada dinámico para el período
        self.periodo_var = tk.StringVar()
        self.entry_periodo = ctk.CTkEntry(self.frame_filtros, textvariable=self.periodo_var, width=120)
        self.entry_periodo.grid(row=0, column=7, padx=5, pady=10)
        self.periodo_var.set(datetime.date.today().strftime("%d-%m-%Y"))

        # Label de bloqueo de frecuencia
        self.lbl_bloqueo = ctk.CTkLabel(self.frame_filtros, text="", text_color="#F59E0B", font=("Outfit", 11, "italic"))
        self.lbl_bloqueo.grid(row=0, column=8, padx=10, pady=10, columnspan=2)

        # 2. Contenedor Principal (Layout de Dos Columnas)
        self.frame_principal = ctk.CTkFrame(self, fg_color="transparent")
        self.frame_principal.pack(fill="both", expand=True, padx=10, pady=5)

        # Columna Izquierda: Lista de Estudiantes
        self.frame_estudiantes = ctk.CTkFrame(self.frame_principal, fg_color="#121c2c", width=260)
        self.frame_estudiantes.pack(side="left", fill="y", padx=(0, 5))
        self.frame_estudiantes.pack_propagate(False)

        ctk.CTkLabel(self.frame_estudiantes, text="Estudiantes", font=("Outfit", 14, "bold"), text_color="#3B82F6").pack(anchor="w", padx=10, pady=10)
        
        self.scroll_estudiantes = ctk.CTkScrollableFrame(self.frame_estudiantes, fg_color="transparent")
        self.scroll_estudiantes.pack(fill="both", expand=True, padx=5, pady=5)

        # Botón para Guardar Todo
        self.btn_guardar_todo = ctk.CTkButton(
            self.frame_estudiantes,
            text="💾 GUARDAR EVALUACIONES",
            font=("Outfit", 12, "bold"),
            fg_color="#10B981",
            hover_color="#059669",
            command=self.guardar_evaluaciones
        )
        self.btn_guardar_todo.pack(fill="x", padx=10, pady=10)

        # Columna Derecha: Formulario de Evaluación
        self.frame_formulario = ctk.CTkFrame(self.frame_principal, fg_color="#121c2c")
        self.frame_formulario.pack(side="right", fill="both", expand=True, padx=(5, 0))

        # Cabecera del Estudiante Seleccionado
        self.lbl_nombre_estudiante = ctk.CTkLabel(
            self.frame_formulario,
            text="Seleccione un estudiante de la lista",
            font=("Outfit", 16, "bold"),
            text_color="#F3F4F6",
            anchor="w"
        )
        self.lbl_nombre_estudiante.pack(fill="x", padx=15, pady=15)

        # Sub-contenedor del Formulario
        self.frame_criterios_form = ctk.CTkScrollableFrame(self.frame_formulario, fg_color="transparent")
        self.frame_criterios_form.pack(fill="both", expand=True, padx=10, pady=5)

        # Panel de Sugerencias
        self.frame_sugerencia = ctk.CTkFrame(self.frame_formulario, fg_color="#1a2436", height=130)
        self.frame_sugerencia.pack(fill="x", padx=15, pady=15)
        self.frame_sugerencia.pack_propagate(False)

        self.btn_sugerir = ctk.CTkButton(
            self.frame_sugerencia,
            text="🧠 Sugerencia de IA",
            font=("Outfit", 12, "bold"),
            fg_color="#3B82F6",
            hover_color="#2563EB",
            command=self.generar_sugerencia,
            width=140
        )
        self.btn_sugerir.pack(side="left", padx=15, pady=15)

        self.txt_explicacion = ctk.CTkTextbox(
            self.frame_sugerencia,
            fg_color="#0F1923",
            font=("Outfit", 11),
            text_color="#A3A3A3",
            wrap="word"
        )
        self.txt_explicacion.pack(side="right", fill="both", expand=True, padx=15, pady=15)
        self.txt_explicacion.insert("1.0", "Presione el botón para generar sugerencias conceptuales basadas en asistencia, notas y comportamiento.")
        self.txt_explicacion.configure(state="disabled")

        # Cargar datos iniciales
        self.al_cambiar_grado_trimestre()

    def al_cambiar_frecuencia(self, frecuencia):
        # Actualiza el widget de selector de período según la frecuencia seleccionada
        if frecuencia == "Diario":
            self.lbl_periodo.configure(text="Fecha:")
            self.entry_periodo.grid_remove()
            self.entry_periodo = ctk.CTkEntry(self.frame_filtros, textvariable=self.periodo_var, width=120)
            self.entry_periodo.grid(row=0, column=7, padx=5, pady=10)
            self.periodo_var.set(datetime.date.today().strftime("%d-%m-%Y"))
        elif frecuencia == "Semanal":
            self.lbl_periodo.configure(text="Semana:")
            self.entry_periodo.grid_remove()
            semanas = [f"Semana {i}" for i in range(1, 16)]
            self.combo_periodo = ctk.CTkComboBox(self.frame_filtros, values=semanas, variable=self.periodo_var, width=120)
            self.combo_periodo.grid(row=0, column=7, padx=5, pady=10)
            self.periodo_var.set("Semana 1")
            self.entry_periodo = self.combo_periodo
        elif frecuencia == "Mensual":
            self.lbl_periodo.configure(text="Mes:")
            self.entry_periodo.grid_remove()
            meses = ["Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]
            self.combo_periodo = ctk.CTkComboBox(self.frame_filtros, values=meses, variable=self.periodo_var, width=120)
            self.combo_periodo.grid(row=0, column=7, padx=5, pady=10)
            self.periodo_var.set("Marzo")
            self.entry_periodo = self.combo_periodo

        self.cargar_evaluaciones_guardadas()

    def al_cambiar_grado_trimestre(self, *args):
        grado = self.combo_grado.get()
        trimestre = self.combo_trimestre.get()

        # 1. Cargar estudiantes
        self.estudiantes = self.engine.obtener_estudiantes_completos(grado)
        
        # Limpiar formulario y selección
        self.estudiante_seleccionado = None
        self.evaluaciones_temporales = {}
        self.lbl_nombre_estudiante.configure(text="Seleccione un estudiante de la lista")
        for widget in self.frame_criterios_form.winfo_children():
            widget.destroy()

        # 2. Verificar bloqueo de frecuencia
        self.verificar_bloqueo_frecuencia(grado, trimestre)

        # 3. Cargar la lista en la UI
        self.refrescar_lista_estudiantes()
        self.cargar_evaluaciones_guardadas()

    def verificar_bloqueo_frecuencia(self, grado, trimestre):
        # Carga el JSON y verifica si ya hay evaluaciones para este grado/trimestre
        self.frecuencia_bloqueada = None
        self.lbl_bloqueo.configure(text="")
        self.btn_frecuencia.configure(state="normal")

        if os.path.exists(self.ruta_json):
            try:
                with open(self.ruta_json, "r", encoding="utf-8") as f:
                    datos = json.load(f)
                
                # Buscamos claves del tipo "{grado}::{trimestre}::[Frecuencia]::[Periodo]"
                patron = f"{grado}::{trimestre}::"
                for clave in datos.keys():
                    if clave.startswith(patron):
                        partes = clave.split("::")
                        if len(partes) >= 3:
                            self.frecuencia_bloqueada = partes[2]
                            break
            except Exception as e:
                print(f"Error cargando JSON de hábitos: {e}")

        if self.frecuencia_bloqueada:
            self.btn_frecuencia.set(self.frecuencia_bloqueada)
            self.btn_frecuencia.configure(state="disabled")
            self.lbl_bloqueo.configure(text=f"⚠️ Frecuencia bloqueada en [{self.frecuencia_bloqueada}] para este trimestre.")
            self.al_cambiar_frecuencia(self.frecuencia_bloqueada)

    def refrescar_lista_estudiantes(self):
        # Limpiar lista anterior
        for widget in self.scroll_estudiantes.winfo_children():
            widget.destroy()

        if not self.estudiantes or "Sin estudiantes" in self.estudiantes[0]['nombre']:
            return

        self.botones_estudiantes = {}
        for est in self.estudiantes:
            btn = ctk.CTkButton(
                self.scroll_estudiantes,
                text=f"{est['nombre'][:22]}",
                font=("Outfit", 12),
                fg_color="transparent",
                text_color="#A3A3A3",
                hover_color="#1a2436",
                anchor="w",
                command=lambda e=est: self.seleccionar_estudiante(e)
            )
            btn.pack(fill="x", pady=2)
            self.botones_estudiantes[est['id']] = btn

    def seleccionar_estudiante(self, est):
        # Deseleccionar anterior
        if self.estudiante_seleccionado:
            self.botones_estudiantes[self.estudiante_seleccionado['id']].configure(
                fg_color="transparent", text_color="#A3A3A3"
            )

        self.estudiante_seleccionado = est
        self.botones_estudiantes[est['id']].configure(
            fg_color="#3B82F6", text_color="#FFFFFF"
        )
        self.lbl_nombre_estudiante.configure(text=f"Evaluando: {est['nombre']}")

        # Limpiar explicación de sugerencia
        self.txt_explicacion.configure(state="normal")
        self.txt_explicacion.delete("1.0", "end")
        self.txt_explicacion.insert("1.0", "Presione el botón para generar sugerencias conceptuales basadas en asistencia, notas y comportamiento.")
        self.txt_explicacion.configure(state="disabled")

        # Cargar criterios correspondientes (Primaria o Premedia)
        self.cargar_formulario_criterios()

    def cargar_formulario_criterios(self):
        for widget in self.frame_criterios_form.winfo_children():
            widget.destroy()

        grado = self.combo_grado.get()
        # Determinar si es primaria o premedia/media basándose en el grado
        es_primaria = any(g in grado for g in ["1°", "2°", "3°", "4°", "5°", "6°"])
        criterios = CRITERIOS_PRIMARIA if es_primaria else CRITERIOS_PREMEDIA_MEDIA

        id_est = self.estudiante_seleccionado['id']
        if id_est not in self.evaluaciones_temporales:
            self.evaluaciones_temporales[id_est] = {crit[0]: "S" for crit in criterios}

        self.widgets_criterios = {}

        for crit_name, desc in criterios:
            row = ctk.CTkFrame(self.frame_criterios_form, fg_color="transparent")
            row.pack(fill="x", pady=8, padx=5)

            # Información del criterio
            info_frame = ctk.CTkFrame(row, fg_color="transparent")
            info_frame.pack(side="left", fill="both", expand=True)
            
            lbl_crit = ctk.CTkLabel(info_frame, text=crit_name, font=("Outfit", 12, "bold"), text_color="#F3F4F6", anchor="w")
            lbl_crit.pack(fill="x")
            
            lbl_desc = ctk.CTkLabel(info_frame, text=desc, font=("Outfit", 10), text_color="#8E9AA8", anchor="w")
            lbl_desc.pack(fill="x")

            # Botones segmentados S / R / X
            val_actual = self.evaluaciones_temporales[id_est].get(crit_name, "S")
            seg_btn = ctk.CTkSegmentedButton(
                row,
                values=["S", "R", "X"],
                width=120,
                selected_color="#10B981" if val_actual == "S" else ("#F59E0B" if val_actual == "R" else "#EF4444"),
                command=lambda val, cn=crit_name: self.cambiar_valor_criterio(cn, val)
            )
            seg_btn.set(val_actual)
            seg_btn.pack(side="right", padx=10)

            self.widgets_criterios[crit_name] = seg_btn

    def cambiar_valor_criterio(self, criterio, valor):
        if not self.estudiante_seleccionado:
            return
        id_est = self.estudiante_seleccionado['id']
        self.evaluaciones_temporales[id_est][criterio] = valor

        # Cambiar color según valor
        color_btn = "#10B981" if valor == "S" else ("#F59E0B" if valor == "R" else "#EF4444")
        self.widgets_criterios[criterio].configure(selected_color=color_btn)

    def cargar_evaluaciones_guardadas(self):
        grado = self.combo_grado.get()
        trimestre = self.combo_trimestre.get()
        frecuencia = self.btn_frecuencia.get()
        periodo = self.periodo_var.get()

        clave = f"{grado}::{trimestre}::{frecuencia}::{periodo}"
        
        if os.path.exists(self.ruta_json):
            try:
                with open(self.ruta_json, "r", encoding="utf-8") as f:
                    datos = json.load(f)
                
                if clave in datos:
                    evals_guardadas = datos[clave].get("estudiantes", {})
                    # Cargar en temporales
                    for id_est, crit_vals in evals_guardadas.items():
                        self.evaluaciones_temporales[int(id_est)] = crit_vals
            except Exception as e:
                print(f"Error cargando evaluaciones guardadas: {e}")

        # Refrescar formulario si hay alguien seleccionado
        if self.estudiante_seleccionado:
            self.cargar_formulario_criterios()

    def generar_sugerencia(self):
        if not self.estudiante_seleccionado:
            messagebox.showwarning("Atención", "Debe seleccionar un estudiante de la lista.")
            return

        grado = self.combo_grado.get()
        trimestre = self.combo_trimestre.get()
        est = self.estudiante_seleccionado
        id_est = est['id']

        self.btn_sugerir.configure(text="⏳ ANALIZANDO...", state="disabled")
        self.update()

        # 1. Obtener estadísticas de asistencia del Excel
        stats_asistencia = self.engine.obtener_estadisticas_asistencia(grado, trimestre, id_est)

        # 2. Obtener tareas con notas vacías del Excel
        stats_notas = self.engine.obtener_tareas_sin_nota(grado, trimestre, id_est)

        # 3. Leer observaciones y reportes del Word del estudiante
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
                            if "Conducta" in tipo or "Citación" in tipo or "Reporte" in tipo or "Disciplina" in tipo:
                                reportes_disciplina.append(f"{tipo}: {motivo}")
            except Exception as e:
                print(f"Error leyendo Word de comportamiento para {est['nombre']}: {e}")

        # Lógica de sugerencias conceptuales
        sugerencias = {}
        comentarios_analisis = []

        es_primaria = any(g in grado for g in ["1°", "2°", "3°", "4°", "5°", "6°"])
        criterios = [crit[0] for crit in (CRITERIOS_PRIMARIA if es_primaria else CRITERIOS_PREMEDIA_MEDIA)]

        # --- Responsabilidad & Puntualidad ---
        limite_ausencias_r = 2
        limite_ausencias_x = 4
        limite_tardanzas_r = 3
        limite_tardanzas_x = 6
        limite_tareas_r = 1
        limite_tareas_x = 3

        aus_sin_just = stats_asistencia["ausencias"]
        tardanzas = stats_asistencia["tardanzas"]
        tareas_vacias = stats_notas["total_vacias"]

        # Responsabilidad
        resp_val = "S"
        if aus_sin_just >= limite_ausencias_x or tareas_vacias >= limite_tareas_x:
            resp_val = "X"
            comentarios_analisis.append(f"• Responsabilidad -> Sugerido 'X' por excesiva inasistencia ({aus_sin_just} ausencias) o tareas sin entregar ({tareas_vacias}).")
        elif aus_sin_just >= limite_ausencias_r or tareas_vacias >= limite_tareas_r:
            resp_val = "R"
            comentarios_analisis.append(f"• Responsabilidad -> Sugerido 'R' por {aus_sin_just} ausencias o {tareas_vacias} tareas vacías.")
        sugerencias["Responsabilidad"] = resp_val

        # Puntualidad (Solo para Premedia/Media)
        if "Puntualidad" in criterios:
            punt_val = "S"
            if tardanzas >= limite_tardanzas_x:
                punt_val = "X"
                comentarios_analisis.append(f"• Puntualidad -> Sugerido 'X' por tardanzas frecuentes ({tardanzas} tardanzas).")
            elif tardanzas >= limite_tardanzas_r:
                punt_val = "R"
                comentarios_analisis.append(f"• Puntualidad -> Sugerido 'R' por {tardanzas} tardanzas.")
            sugerencias["Puntualidad"] = punt_val

        # Organización del Trabajo
        org_val = "S"
        if tareas_vacias >= limite_tareas_x:
            org_val = "X"
            comentarios_analisis.append(f"• Org. del Trabajo -> Sugerido 'X' debido a {tareas_vacias} tareas vacías.")
        elif tareas_vacias >= limite_tareas_r:
            org_val = "R"
            comentarios_analisis.append(f"• Org. del Trabajo -> Sugerido 'R' debido a {tareas_vacias} tareas vacías.")
        sugerencias["Organización del Trabajo"] = org_val

        # Autodominio y Confianza en sí mismo / Cooperación / Modales / Honradez
        con_val = "S"
        coop_val = "S"
        mod_val = "S"
        hon_val = "S"

        if reportes_disciplina:
            comentarios_analisis.append(f"• Comportamiento -> Encontrado(s) {len(reportes_disciplina)} reporte(s) en su expediente:")
            for rep in reportes_disciplina[:2]:
                comentarios_analisis.append(f"  - {rep[:60]}...")
            con_val = "R"
            coop_val = "R"
            mod_val = "R"
            hon_val = "R"
            
            # Si tiene reportes serios
            if any("Grave" in r or "Suspensión" in r for r in reportes_disciplina):
                con_val = "X"
                coop_val = "X"
                mod_val = "X"

        sugerencias["Autodominio y Confianza en sí mismo"] = con_val
        sugerencias["Cooperación"] = coop_val
        if "Modales" in criterios:
            sugerencias["Modales"] = mod_val
        if "Honradez" in criterios:
            sugerencias["Honradez"] = hon_val

        # Rellenar criterios por defecto que no dependan directamente de estadísticas
        for crit in criterios:
            if crit not in sugerencias:
                sugerencias[crit] = "S"

        # Aplicar sugerencias en temporales
        for crit, val in sugerencias.items():
            self.evaluaciones_temporales[id_est][crit] = val
            if crit in self.widgets_criterios:
                self.widgets_criterios[crit].set(val)
                color_btn = "#10B981" if val == "S" else ("#F59E0B" if val == "R" else "#EF4444")
                self.widgets_criterios[crit].configure(selected_color=color_btn)

        # Mostrar explicaciones en el TextBox
        self.txt_explicacion.configure(state="normal")
        self.txt_explicacion.delete("1.0", "end")
        
        texto_explicativo = f"Análisis de {est['nombre']}:\n"
        texto_explicativo += f"- Asistencia: {stats_asistencia['total_dias']} días activos, {stats_asistencia['ausencias']} faltas, {stats_asistencia['tardanzas']} tardanzas.\n"
        texto_explicativo += f"- Notas: {stats_notas['total_vacias']} tareas con notas vacías.\n"
        if reportes_disciplina:
            texto_explicativo += f"- Conducta: {len(reportes_disciplina)} anotaciones en Word.\n"
        else:
            texto_explicativo += "- Conducta: Sin anotaciones de disciplina.\n"

        if comentarios_analisis:
            texto_explicativo += "\nRecomendaciones de Ajuste:\n" + "\n".join(comentarios_analisis)
        else:
            texto_explicativo += "\n✓ Rendimiento excelente: Sugerido 'S' (Satisfactorio) en todos los campos."

        self.txt_explicacion.insert("1.0", texto_explicativo)
        self.txt_explicacion.configure(state="disabled")

        self.btn_sugerir.configure(text="🧠 Sugerencia de IA", state="normal")

    def toggle_saving_state(self, is_saving):
        state = "disabled" if is_saving else "normal"
        self.btn_guardar_todo.configure(state=state)
        self.btn_sugerir.configure(state=state)
        self.combo_grado.configure(state=state)
        self.combo_trimestre.configure(state=state)
        self.btn_frecuencia.configure(state=state)
        if is_saving:
            self.btn_guardar_todo.configure(text="⏳ GUARDANDO...")
        else:
            self.btn_guardar_todo.configure(text="💾 GUARDAR EVALUACIONES")

    def guardar_evaluaciones(self):
        grado = self.combo_grado.get()
        trimestre = self.combo_trimestre.get()
        frecuencia = self.btn_frecuencia.get()
        periodo = self.periodo_var.get().strip()

        if not periodo:
            messagebox.showwarning("Atención", "Debe colocar una fecha o período válido.")
            return

        if not self.evaluaciones_temporales:
            messagebox.showwarning("Atención", "No hay datos de evaluación para guardar.")
            return

        self.toggle_saving_state(True)

        def tarea_fondo():
            clave = f"{grado}::{trimestre}::{frecuencia}::{periodo}"

            datos = {}
            if os.path.exists(self.ruta_json):
                try:
                    with open(self.ruta_json, "r", encoding="utf-8") as f:
                        datos = json.load(f)
                except Exception as e:
                    print(f"Error cargando JSON para escribir: {e}")

            datos[clave] = {
                "frecuencia": frecuencia,
                "periodo": periodo,
                "estudiantes": {str(k): v for k, v in self.evaluaciones_temporales.items()}
            }

            carpeta_exp = os.path.dirname(self.ruta_json)
            if not os.path.exists(carpeta_exp):
                os.makedirs(carpeta_exp)

            try:
                with open(self.ruta_json, "w", encoding="utf-8") as f:
                    json.dump(datos, f, ensure_ascii=False, indent=4)
            except Exception as e:
                self.after(0, lambda: self.finalizar_guardado_habitos(False, str(e), grado, trimestre))
                return

            if DOCX_DISPONIBLE:
                for est in self.estudiantes:
                    id_est = est['id']
                    if id_est in self.evaluaciones_temporales:
                        nombre_archivo = f"{est['nombre']} - {grado.replace('°', '')}.docx".replace("/", "-")
                        ruta_word = os.path.join(carpeta_exp, nombre_archivo)
                        
                        if os.path.exists(ruta_word):
                            try:
                                doc = Document(ruta_word)
                                if doc.tables:
                                    tabla = doc.tables[0]
                                    
                                    crit_vals = self.evaluaciones_temporales[id_est]
                                    eval_str = ", ".join([f"{k}: {v}" for k, v in crit_vals.items()])
                                    
                                    tipo_registro = f"Hábitos ({frecuencia}: {periodo})"
                                    
                                    fila_existente = None
                                    for row in tabla.rows[1:]:
                                        if len(row.cells) >= 4:
                                            f = row.cells[1].text.strip()
                                            t = row.cells[2].text.strip()
                                            if f == periodo and t == tipo_registro:
                                                fila_existente = row
                                                break
                                    
                                    if fila_existente:
                                        fila_existente.cells[3].text = eval_str
                                    else:
                                        num_reg = str(len(tabla.rows))
                                        fila = tabla.add_row()
                                        fila.cells[0].text = num_reg
                                        fila.cells[1].text = periodo
                                        fila.cells[2].text = tipo_registro
                                        fila.cells[3].text = eval_str
                                    
                                    doc.save(ruta_word)
                            except Exception as e:
                                print(f"Error escribiendo en Word de {est['nombre']}: {e}")

            self.after(0, lambda: self.finalizar_guardado_habitos(True, "", grado, trimestre))

        threading.Thread(target=tarea_fondo, daemon=True).start()

    def finalizar_guardado_habitos(self, exito, error_msg, grado, trimestre):
        self.toggle_saving_state(False)
        if not exito:
            messagebox.showerror("Error", f"No se pudo guardar la evaluación: {error_msg}")
            return
            
        self.verificar_bloqueo_frecuencia(grado, trimestre)

        root = self.winfo_toplevel()
        if hasattr(root, "mostrar_toast"):
            root.mostrar_toast("✓ Evaluaciones guardadas con éxito", color="#10B981")
        else:
            messagebox.showinfo("Éxito", "Evaluaciones registradas y guardadas correctamente.")

    def actualizar_vista(self):
        """Recarga la lista de grados y estudiantes."""
        opciones = self.engine.obtener_grados_activos() or ["Sin datos"]
        old_sel = self.combo_grado.get()
        self.combo_grado.configure(values=opciones)
        if old_sel in opciones:
            self.combo_grado.set(old_sel)
        else:
            self.combo_grado.set(opciones[0])
        self.al_cambiar_grado_trimestre()

