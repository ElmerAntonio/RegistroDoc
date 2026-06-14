"""
RegistroDoc Pro — Centro de Ayuda Interactivo
Diseño único: guías paso-a-paso, glosario MEDUCA, tips, atajos, checklist.
100% offline — sin dependencia de internet.
"""
import os, datetime, random
import customtkinter as ctk
from tkinter import messagebox
from config import BASE_DIR
from theme import C, FONT_TITLE, FONT_BODY, FONT_SIZES

# ─── DATOS OFFLINE ──────────────────────────────────────────────
TIPS_DIA = [
    "💡 Use Ctrl+S en cualquier pantalla para guardar rápido.",
    "💡 Marque 'Todos Presentes' y luego corrija solo las excepciones.",
    "💡 El botón 🧠 Sugerencia de IA en Hábitos analiza notas y asistencia automáticamente.",
    "💡 Guarde un respaldo en USB cada viernes — su trabajo es valioso.",
    "💡 Si el Excel está abierto en otro programa, ciérrelo antes de guardar aquí.",
    "💡 Use la pestaña 'Modificar' para corregir notas o asistencia ya guardadas.",
    "💡 Las Excusas (E) no restan puntos de asistencia al alumno.",
    "💡 Use Ctrl+1 a Ctrl+9 para navegar rápido entre secciones.",
    "💡 El Dashboard muestra alertas de alumnos en riesgo automáticamente.",
    "💡 Puede buscar alumnos en el Dashboard escribiendo su nombre.",
    "💡 Las observaciones se guardan automáticamente en Word por alumno.",
    "💡 En Hábitos, la frecuencia se bloquea después del primer guardado del trimestre.",
]

GLOSARIO = [
    ("S (Satisfactorio)", "El alumno cumple bien con el criterio evaluado. Color verde."),
    ("R (Regular)", "El alumno cumple parcialmente. Necesita mejorar. Color amarillo."),
    ("X (No Satisface)", "El alumno no cumple con el criterio. Requiere atención. Color rojo."),
    ("Parcial / Diaria", "Nota de tareas, talleres o actividades de clase. Valor: 1.0 a 5.0."),
    ("Apreciación", "Nota de cuadernos, participación o proyectos. Valor: 1.0 a 5.0."),
    ("Examen", "Prueba trimestral oficial. La descripción se genera automáticamente."),
    ("P (Presente)", "El alumno asistió a clase. Se guarda como '.' en Excel."),
    ("A (Ausente)", "El alumno faltó sin justificación. Se guarda como '-' en Excel."),
    ("T (Tardanza)", "El alumno llegó tarde. Se guarda como 'T' en Excel."),
    ("E (Excusa)", "Falta justificada con certificado. NO resta puntos de asistencia."),
    ("Trimestre", "Período académico (3 por año). Cada uno tiene sus propias notas."),
    ("Expediente", "Documento Word individual por alumno con su historial completo."),
    ("Sincronizar", "Copiar los datos de configuración a todas las hojas del Excel."),
    ("Consejero", "Profesor encargado de un grupo/grado específico."),
]

ATAJOS = [
    ("Ctrl + 1", "Ir al Dashboard (Inicio)"),
    ("Ctrl + 2", "Ir a Estudiantes"),
    ("Ctrl + 3", "Ir a Notas"),
    ("Ctrl + 4", "Ir a Asistencia"),
    ("Ctrl + 5", "Ir a Observaciones"),
    ("Ctrl + 6", "Ir a Hábitos"),
    ("Ctrl + 7", "Ir a Reportes"),
    ("Ctrl + 8", "Ir a Gráficos"),
    ("Ctrl + 9", "Ir a Impresión"),
    ("Ctrl + S", "Guardar en la pantalla actual"),
    ("F1", "Abrir esta Ayuda"),
    ("Escape", "Volver al Dashboard"),
]

GUIAS = [
    {
        "icono": "📝", "titulo": "Poner Notas a mis Alumnos",
        "color": C["info"],
        "pasos": [
            "Presione 'Notas' en el menú izquierdo (o Ctrl+3)",
            "Seleccione el Grado y la Materia en los desplegables",
            "Elija el Trimestre y Tipo de nota (Parcial, Apreciación o Examen)",
            "Escriba la Fecha y una Descripción corta del trabajo",
            "Coloque la nota de cada alumno (1.0 a 5.0) en su casilla",
            "Presione el botón verde 💾 GUARDAR NUEVA",
        ]
    },
    {
        "icono": "📅", "titulo": "Pasar Lista de Asistencia",
        "color": C["success"],
        "pasos": [
            "Presione 'Asistencia' en el menú izquierdo (o Ctrl+4)",
            "Seleccione el Grado y Trimestre",
            "Escriba la fecha de hoy (DD-MM)",
            "Por defecto todos están en P (Presente)",
            "Cambie a A, T o E solo los que faltaron/llegaron tarde",
            "Escriba el motivo de la ausencia o tardanza",
            "Presione 💾 GUARDAR ASISTENCIA",
        ]
    },
    {
        "icono": "👤", "titulo": "Agregar un Alumno Nuevo",
        "color": C["cian"],
        "pasos": [
            "Presione 'Estudiantes' en el menú izquierdo (o Ctrl+2)",
            "En la barra verde superior escriba: APELLIDO, Nombre",
            "Opcionalmente escriba la cédula y seleccione el sexo",
            "Presione 'Guardar Nuevo'",
            "El alumno aparecerá al final de la lista",
        ]
    },
    {
        "icono": "🔍", "titulo": "Registrar una Observación de Conducta",
        "color": C["warning"],
        "pasos": [
            "Presione 'Observaciones' en el menú izquierdo (o Ctrl+5)",
            "Seleccione el Grado y haga clic en el nombre del alumno",
            "Elija la categoría: Conducta, Académico, Citación, etc.",
            "Puede elegir una plantilla o escribir libremente",
            "Presione 📝 GUARDAR EN EXPEDIENTE OFICIAL",
            "Se creará un documento Word en la carpeta Expedientes",
        ]
    },
    {
        "icono": "🧠", "titulo": "Calificar Hábitos y Aptitudes",
        "color": C["purpura"],
        "pasos": [
            "Presione 'Hábitos' en el menú izquierdo (o Ctrl+6)",
            "Seleccione Grado, Trimestre y Frecuencia (Diario/Semanal/Mensual)",
            "⚠️ La frecuencia se BLOQUEA después del primer guardado",
            "Haga clic en un alumno de la lista izquierda",
            "Califique cada criterio: S (verde), R (amarillo), X (rojo)",
            "Use el botón 🧠 Sugerencia para que el sistema analice al alumno",
            "Presione 💾 GUARDAR EVALUACIONES",
        ]
    },
    {
        "icono": "🖨️", "titulo": "Imprimir Planillas y Reportes",
        "color": C["rosa"],
        "pasos": [
            "Presione 'Impresión' en el menú izquierdo (o Ctrl+9)",
            "Seleccione el tipo: Portada, Planilla, Asistencia, Resumen...",
            "Seleccione Grado y Materia si aplica",
            "Presione '📂 Abrir en Excel' para ver antes de imprimir",
            "O presione '🖨️ Enviar a Impresora' para imprimir directo",
        ]
    },
    {
        "icono": "📋", "titulo": "Ver Registro Completo Consolidado",
        "color": C["cian"],
        "pasos": [
            "Presione 'Registro Completo' en el menú izquierdo",
            "Seleccione el Grado, Trimestre y Materia a visualizar",
            "Pestaña Calificaciones: Vea métricas y cuadrícula de notas de todo el salón",
            "Pestaña Asistencia: Vea la matriz completa de asistencia por fechas del trimestre",
        ]
    },
    {
        "icono": "✍️", "titulo": "Usar Corrector y Guardar Plantillas",
        "color": C["warning"],
        "pasos": [
            "En Observaciones, redacte la conducta o académico libremente",
            "Las palabras mal escritas se subrayarán en rojo de forma automática",
            "Haga clic derecho en una palabra subrayada para seleccionar sugerencias",
            "Use el botón '💾 Guardar como nueva plantilla' para clasificar y guardarla",
        ]
    },
]

FAQS = [
    ("¿Por qué no puedo cambiar la frecuencia de Hábitos?",
     "Una vez que guarda la primera calificación en el trimestre, la frecuencia se bloquea para mantener los datos ordenados. Debe terminar el trimestre o borrar las evaluaciones para cambiarla."),
    ("¿Qué significa 'E' en la asistencia?",
     "'E' es una Falta Justificada (Excusa). A diferencia de la falta '-', las fórmulas de MEDUCA NO la restarán, por lo que el alumno no perderá puntaje."),
    ("¿Dónde se guardan los expedientes de conducta?",
     "En la carpeta 'Expedientes_Estudiantes' como archivos Word (.docx). Hay uno por alumno con todo su historial."),
    ("¿Cómo corrijo una nota ya guardada?",
     "En la pantalla de Notas, cambie a la pestaña 'Modificar', seleccione la descripción del trabajo, haga los cambios y presione 'ACTUALIZAR EXCEL'."),
    ("¿Cómo cambio una Ausencia a Excusa?",
     "En Asistencia → pestaña 'Modificar' → seleccione la fecha → cambie A por E → escriba la justificación → ACTUALIZAR EXCEL."),
    ("El programa da error al guardar, ¿qué hago?",
     "Asegúrese de que el archivo Excel NO esté abierto en Microsoft Excel. Ciérrelo completamente e intente de nuevo."),
    ("¿Cómo hago un respaldo de seguridad?",
     "En el Dashboard hay un botón '💾 Respaldo'. También puede copiar manualmente el archivo Excel y la carpeta Expedientes a un USB."),
    ("¿Qué son las notas vacías que menciona el sistema?",
     "Son celdas donde se creó una tarea pero no se le puso nota al alumno. El sistema las cuenta como 'no entregadas' para las sugerencias."),
    ("¿Puedo usar el programa sin internet?",
     "¡Sí! RegistroDoc Pro funciona 100% offline. No necesita conexión a internet para nada."),
    ("¿Cómo agrego una materia nueva a un grado?",
     "En Configuración → pestaña 'Gestión de Materias' → seleccione el grado, escriba el nombre y presione 'Crear Materia'."),
    ("¿Qué pasa si elimino un grado por error?",
     "La eliminación es permanente. Por eso siempre haga respaldos antes de eliminar. Puede restaurar desde un respaldo anterior."),
    ("¿Cómo sincronizo los datos de portada?",
     "En Configuración → pestaña 'Datos de Portada' → llene todos los campos → presione 'SINCRONIZAR Y SOBREESCRIBIR EXCEL'."),
]


class HelpFrame(ctk.CTkFrame):
    def __init__(self, master, engine, **kwargs):
        super().__init__(master, fg_color="transparent", **kwargs)
        self.engine = engine
        self._widgets_buscables = []
        self._tip_idx = random.randint(0, len(TIPS_DIA) - 1)
        self.crear_interfaz()

    def crear_interfaz(self):
        # ═══ HEADER ═══
        hdr = ctk.CTkFrame(self, fg_color=C["card"], corner_radius=12,
                           border_width=1, border_color=C["cian"])
        hdr.pack(fill="x", padx=15, pady=(12, 6))

        hdr_left = ctk.CTkFrame(hdr, fg_color="transparent")
        hdr_left.pack(side="left", fill="x", expand=True, padx=18, pady=12)

        ctk.CTkLabel(hdr_left, text="📚 Centro de Ayuda RegistroDoc Pro",
                     font=ctk.CTkFont(FONT_TITLE, 22, "bold"),
                     text_color=C["cian"]).pack(anchor="w")
        ctk.CTkLabel(hdr_left, text="Todo lo que necesita saber, sin internet, en un solo lugar.",
                     font=ctk.CTkFont(FONT_BODY, 12),
                     text_color=C["texto_sec"]).pack(anchor="w", pady=(2, 0))

        # Buscador
        self.var_busqueda = ctk.StringVar()
        self.var_busqueda.trace_add("write", self._filtrar)
        srch = ctk.CTkEntry(hdr, textvariable=self.var_busqueda,
                            placeholder_text="🔍 Buscar en toda la ayuda...",
                            font=ctk.CTkFont(FONT_BODY, 12), width=280,
                            fg_color=C["input"], border_color=C["borde"],
                            text_color=C["texto"])
        srch.pack(side="right", padx=18, pady=12)

        # ═══ TIP DEL DÍA ═══
        tip_frame = ctk.CTkFrame(self, fg_color=C["card_alt"], corner_radius=10,
                                 border_width=1, border_color=C["warning"])
        tip_frame.pack(fill="x", padx=15, pady=(6, 4))

        tip_inner = ctk.CTkFrame(tip_frame, fg_color="transparent")
        tip_inner.pack(fill="x", padx=14, pady=10)

        self.lbl_tip = ctk.CTkLabel(tip_inner,
                                     text=TIPS_DIA[self._tip_idx],
                                     font=ctk.CTkFont(FONT_BODY, 12, "bold"),
                                     text_color=C["amarillo"], anchor="w")
        self.lbl_tip.pack(side="left", fill="x", expand=True)

        ctk.CTkButton(tip_inner, text="Otro tip ↻", width=80, height=26,
                      fg_color=C["badge_bg"], hover_color=C["hover"],
                      font=ctk.CTkFont(FONT_BODY, 11),
                      command=self._siguiente_tip).pack(side="right")

        # ═══ TABVIEW PRINCIPAL ═══
        self.tabview = ctk.CTkTabview(self,
            segmented_button_selected_color=C["cian"],
            segmented_button_selected_hover_color=C["hover"],
            segmented_button_unselected_hover_color=C["hover"],
            text_color=C["texto"])
        self.tabview.pack(fill="both", expand=True, padx=15, pady=6)

        tab_guias = self.tabview.add("📖 Guías Paso a Paso")
        tab_faqs = self.tabview.add("💬 Preguntas Frecuentes")
        tab_glosario = self.tabview.add("📘 Glosario MEDUCA")
        tab_atajos = self.tabview.add("⌨️ Atajos de Teclado")
        tab_manual = self.tabview.add("📄 Manual Completo")

        self._crear_tab_guias(tab_guias)
        self._crear_tab_faqs(tab_faqs)
        self._crear_tab_glosario(tab_glosario)
        self._crear_tab_atajos(tab_atajos)
        self._crear_tab_manual(tab_manual)

    # ─── TAB: GUÍAS PASO A PASO ────────────────────────────────────
    def _crear_tab_guias(self, parent):
        scroll = ctk.CTkScrollableFrame(parent, fg_color="transparent",
                                        scrollbar_button_color=C["borde"])
        scroll.pack(fill="both", expand=True, padx=4, pady=4)

        for guia in GUIAS:
            card = self._guia_card(scroll, guia)
            self._widgets_buscables.append({
                "widget": card,
                "texto": guia["titulo"] + " " + " ".join(guia["pasos"])
            })

    def _guia_card(self, parent, guia):
        card = ctk.CTkFrame(parent, fg_color=C["card"], corner_radius=12,
                            border_width=1, border_color=C["borde"])
        card.pack(fill="x", pady=6, padx=4)

        # Header de la guía
        hdr = ctk.CTkFrame(card, fg_color="transparent")
        hdr.pack(fill="x", padx=16, pady=(14, 8))

        # Icono + título
        ctk.CTkLabel(hdr, text=f'{guia["icono"]}  {guia["titulo"]}',
                     font=ctk.CTkFont(FONT_TITLE, 15, "bold"),
                     text_color=guia["color"], anchor="w"
                     ).pack(side="left")

        badge = ctk.CTkLabel(hdr, text=f'{len(guia["pasos"])} pasos',
                             font=ctk.CTkFont(FONT_BODY, 10, "bold"),
                             text_color=C["texto_sec"],
                             fg_color=C["badge_bg"], corner_radius=10,
                             padx=8, pady=2)
        badge.pack(side="right")

        # Pasos numerados
        for i, paso in enumerate(guia["pasos"], 1):
            row = ctk.CTkFrame(card, fg_color="transparent")
            row.pack(fill="x", padx=16, pady=2)

            # Número circular
            num_color = guia["color"]
            ctk.CTkLabel(row, text=str(i), width=26, height=26,
                         fg_color=num_color, corner_radius=13,
                         font=ctk.CTkFont(FONT_BODY, 11, "bold"),
                         text_color=C["fondo"]).pack(side="left", padx=(0, 10))

            ctk.CTkLabel(row, text=paso,
                         font=ctk.CTkFont(FONT_BODY, 12),
                         text_color=C["texto"], anchor="w",
                         wraplength=600, justify="left"
                         ).pack(side="left", fill="x", expand=True)

        # Espaciado final
        ctk.CTkFrame(card, fg_color="transparent", height=8).pack()

        # Hover effect
        card.bind("<Enter>", lambda e: card.configure(border_color=guia["color"]))
        card.bind("<Leave>", lambda e: card.configure(border_color=C["borde"]))
        return card

    # ─── TAB: FAQs ─────────────────────────────────────────────────
    def _crear_tab_faqs(self, parent):
        scroll = ctk.CTkScrollableFrame(parent, fg_color="transparent",
                                        scrollbar_button_color=C["borde"])
        scroll.pack(fill="both", expand=True, padx=4, pady=4)

        for pregunta, respuesta in FAQS:
            card = ctk.CTkFrame(scroll, fg_color=C["card"], corner_radius=10,
                                border_width=1, border_color=C["borde"])
            card.pack(fill="x", pady=5, padx=4)

            ctk.CTkLabel(card, text=f"❓ {pregunta}",
                         font=ctk.CTkFont(FONT_TITLE, 13, "bold"),
                         text_color=C["amarillo"], anchor="w",
                         wraplength=700, justify="left"
                         ).pack(fill="x", padx=15, pady=(12, 4))

            ctk.CTkLabel(card, text=f"→ {respuesta}",
                         font=ctk.CTkFont(FONT_BODY, 12),
                         text_color=C["texto"], anchor="w",
                         wraplength=700, justify="left"
                         ).pack(fill="x", padx=15, pady=(0, 12))

            card.bind("<Enter>", lambda e, c=card: c.configure(border_color=C["amarillo"]))
            card.bind("<Leave>", lambda e, c=card: c.configure(border_color=C["borde"]))

            self._widgets_buscables.append({
                "widget": card, "texto": pregunta + " " + respuesta
            })

    # ─── TAB: GLOSARIO MEDUCA ──────────────────────────────────────
    def _crear_tab_glosario(self, parent):
        scroll = ctk.CTkScrollableFrame(parent, fg_color="transparent",
                                        scrollbar_button_color=C["borde"])
        scroll.pack(fill="both", expand=True, padx=4, pady=4)

        ctk.CTkLabel(scroll, text="📘 Términos y Conceptos del Sistema MEDUCA",
                     font=ctk.CTkFont(FONT_TITLE, 16, "bold"),
                     text_color=C["cian"]).pack(anchor="w", padx=10, pady=(8, 12))

        for i, (termino, definicion) in enumerate(GLOSARIO):
            bg = C["card"] if i % 2 == 0 else C["card_alt"]
            row = ctk.CTkFrame(scroll, fg_color=bg, corner_radius=8)
            row.pack(fill="x", pady=2, padx=4)

            ctk.CTkLabel(row, text=termino,
                         font=ctk.CTkFont(FONT_BODY, 13, "bold"),
                         text_color=C["cian"], anchor="w", width=240
                         ).pack(side="left", padx=14, pady=10)

            ctk.CTkLabel(row, text=definicion,
                         font=ctk.CTkFont(FONT_BODY, 12),
                         text_color=C["texto"], anchor="w",
                         wraplength=500, justify="left"
                         ).pack(side="left", fill="x", expand=True, padx=(0, 14), pady=10)

            self._widgets_buscables.append({
                "widget": row, "texto": termino + " " + definicion
            })

    # ─── TAB: ATAJOS ───────────────────────────────────────────────
    def _crear_tab_atajos(self, parent):
        scroll = ctk.CTkScrollableFrame(parent, fg_color="transparent",
                                        scrollbar_button_color=C["borde"])
        scroll.pack(fill="both", expand=True, padx=4, pady=4)

        ctk.CTkLabel(scroll, text="⌨️ Atajos de Teclado — Trabaje más rápido",
                     font=ctk.CTkFont(FONT_TITLE, 16, "bold"),
                     text_color=C["cian"]).pack(anchor="w", padx=10, pady=(8, 12))

        for tecla, accion in ATAJOS:
            row = ctk.CTkFrame(scroll, fg_color=C["card"], corner_radius=8)
            row.pack(fill="x", pady=3, padx=4)

            # Key badge
            ctk.CTkLabel(row, text=f"  {tecla}  ",
                         font=ctk.CTkFont("Consolas" if os.name == "nt" else "monospace", 13, "bold"),
                         fg_color=C["badge_bg"], corner_radius=6,
                         text_color=C["cian"], padx=6, pady=4
                         ).pack(side="left", padx=14, pady=8)

            ctk.CTkLabel(row, text=accion,
                         font=ctk.CTkFont(FONT_BODY, 13),
                         text_color=C["texto"], anchor="w"
                         ).pack(side="left", padx=12, pady=8)

    # ─── TAB: MANUAL COMPLETO ──────────────────────────────────────
    def _crear_tab_manual(self, parent):
        f_top = ctk.CTkFrame(parent, fg_color="transparent")
        f_top.pack(fill="x", padx=10, pady=(5, 8))

        ctk.CTkLabel(f_top, text="📖 Manual de Usuario Oficial — v3.0",
                     font=ctk.CTkFont(FONT_TITLE, 14, "bold"),
                     text_color=C["cian"]).pack(side="left", pady=5)

        ctk.CTkButton(f_top, text="📄 Abrir / Imprimir",
                      font=ctk.CTkFont(FONT_TITLE, 11, "bold"),
                      fg_color=C["cian"], hover_color=C["hover"],
                      text_color=C["fondo"], command=self._abrir_manual,
                      height=30, width=160).pack(side="right", padx=5)

        # Contenedor con navegación lateral + contenido
        body = ctk.CTkFrame(parent, fg_color="transparent")
        body.pack(fill="both", expand=True, padx=10, pady=4)
        body.columnconfigure(0, weight=0)
        body.columnconfigure(1, weight=1)
        body.rowconfigure(0, weight=1)

        # ─── Menú de navegación lateral ───
        nav = ctk.CTkScrollableFrame(body, fg_color=C["card"], width=200,
                                      corner_radius=8, scrollbar_button_color=C["borde"])
        nav.grid(row=0, column=0, sticky="nsew", padx=(0, 6))

        ctk.CTkLabel(nav, text="📑 Ir a sección:",
                     font=ctk.CTkFont(FONT_BODY, 11, "bold"),
                     text_color=C["cian"]).pack(anchor="w", padx=8, pady=(8, 6))

        secciones = [
            ("1. Iniciar Programa", "1. CÓMO INICIAR"),
            ("2. Dashboard", "2. PANTALLA DE INICIO"),
            ("3. Estudiantes", "3. MÓDULO DE ESTUDIANTES"),
            ("4. Notas", "4. MÓDULO DE CALIFICACIONES"),
            ("5. Asistencia", "5. MÓDULO DE ASISTENCIA"),
            ("6. Observaciones", "6. MÓDULO DE OBSERVACIONES"),
            ("7. Hábitos", "7. MÓDULO DE HÁBITOS"),
            ("8. Tareas ¡NUEVO!", "8. MÓDULO DE TAREAS"),
            ("9. Reportes/Gráficos", "9. MÓDULO DE REPORTES"),
            ("10. Impresión", "10. MÓDULO DE IMPRESIÓN"),
            ("11. Configuración", "11. CONFIGURACIÓN"),
            ("12. Atajos Teclado", "12. ATAJOS DE TECLADO"),
            ("13. Respaldos", "13. COPIAS DE SEGURIDAD"),
            ("14. Reuniones ¡NUEVO!", "14. REUNIONES"),
            ("15. Reg. Completo ¡NUEVO!", "15. REGISTRO COMPLETO"),
            ("16. Problemas", "16. SOLUCIÓN DE PROBLEMAS"),
            ("17. Glosario", "17. GLOSARIO"),
            ("18. Base de Datos", "17. ESTRUCTURA DE BASE DE DATOS"),
            ("19. Seg. e ISO", "18. CUMPLIMIENTO DE SEGURIDAD"),
        ]

        # ─── Visor de texto ───
        self._manual_textbox = ctk.CTkTextbox(body, fg_color=C["card"],
                                               text_color=C["texto"],
                                               font=("Consolas", 11), wrap="word",
                                               border_width=1, border_color=C["borde"])
        self._manual_textbox.grid(row=0, column=1, sticky="nsew")

        ruta = os.path.join(BASE_DIR, "..", "docs", "Manual_Usuario_RegistroDoc.md")
        contenido = ""
        if os.path.exists(ruta):
            try:
                with open(ruta, "r", encoding="utf-8") as f:
                    contenido = f.read()
            except Exception as e:
                contenido = f"Error al cargar: {e}"
        else:
            contenido = "Manual no encontrado."

        self._manual_textbox.insert("1.0", contenido)
        self._manual_textbox.configure(state="disabled")

        # Crear botones de navegación
        def _hacer_navegar(buscar_texto):
            def _navegar():
                self._manual_textbox.configure(state="normal")
                inner = self._manual_textbox._textbox
                idx = inner.search(buscar_texto, "38.0", nocase=True)
                if idx:
                    inner.see(idx)
                    inner.mark_set("insert", idx)
                    try:
                        inner.tag_remove("sel", "1.0", "end")
                        inner.tag_add("sel", idx, f"{idx} lineend")
                    except Exception:
                        pass
                self._manual_textbox.configure(state="disabled")
                inner.update_idletasks()
            return _navegar

        for titulo, buscar in secciones:
            es_nuevo = "NUEVO" in titulo
            color_txt = C["amarillo"] if es_nuevo else C["texto_sec"]

            ctk.CTkButton(nav, text=titulo,
                          font=ctk.CTkFont(FONT_BODY, 10),
                          fg_color="transparent", hover_color=C["hover"],
                          text_color=color_txt, anchor="w",
                          height=26, command=_hacer_navegar(buscar)).pack(fill="x", padx=4, pady=1)

    # ─── FUNCIONES AUXILIARES ──────────────────────────────────────
    def _abrir_manual(self):
        ruta = os.path.join(BASE_DIR, "..", "docs", "Manual_Usuario_RegistroDoc.md")
        if os.path.exists(ruta):
            try:
                os.startfile(ruta)
            except Exception as e:
                messagebox.showerror("Error", f"No se pudo abrir: {e}")
        else:
            messagebox.showerror("Error", "Manual no encontrado.")

    def _siguiente_tip(self):
        self._tip_idx = (self._tip_idx + 1) % len(TIPS_DIA)
        self.lbl_tip.configure(text=TIPS_DIA[self._tip_idx])

    def _filtrar(self, *args):
        termino = self.var_busqueda.get().strip().lower()
        for item in self._widgets_buscables:
            if not termino or termino in item["texto"].lower():
                item["widget"].pack(fill="x", pady=5, padx=4)
            else:
                item["widget"].pack_forget()
