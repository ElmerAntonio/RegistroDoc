"""
RegistroDoc Pro — Pantalla de Reuniones
Genera actas de reuniones de docentes y con padres de familia.
100% offline.
"""
import customtkinter as ctk
from tkinter import messagebox
import datetime
import os

from theme import C, FONT_TITLE, FONT_BODY
from reuniones import generar_acta_docentes, generar_acta_padres


class ReunionesFrame(ctk.CTkFrame):
    def __init__(self, master, engine, **kwargs):
        super().__init__(master, fg_color="transparent", **kwargs)
        self.engine = engine
        self.crear_interfaz()

    def crear_interfaz(self):
        # Header
        hdr = ctk.CTkFrame(self, fg_color=C["card"], corner_radius=12,
                           border_width=1, border_color=C["cian"])
        hdr.pack(fill="x", padx=15, pady=(12, 8))

        ctk.CTkLabel(hdr, text="📋 Generador de Actas de Reuniones",
                     font=ctk.CTkFont(FONT_TITLE, 20, "bold"),
                     text_color=C["cian"]).pack(side="left", padx=18, pady=12)

        # Tabs
        self.tabs = ctk.CTkTabview(self,
            segmented_button_selected_color=C["cian"],
            segmented_button_selected_hover_color=C["hover"])
        self.tabs.pack(fill="both", expand=True, padx=15, pady=6)

        tab_docentes = self.tabs.add("👨‍🏫 Reunión de Docentes")
        tab_padres = self.tabs.add("👪 Reunión con Padres")

        self._crear_tab_docentes(tab_docentes)
        self._crear_tab_padres(tab_padres)

    def _crear_tab_docentes(self, parent):
        scroll = ctk.CTkScrollableFrame(parent, fg_color="transparent")
        scroll.pack(fill="both", expand=True, padx=4, pady=4)

        # Datos generales
        ctk.CTkLabel(scroll, text="📝 Datos de la Reunión",
                     font=ctk.CTkFont(FONT_TITLE, 15, "bold"),
                     text_color=C["cian"]).pack(anchor="w", padx=10, pady=(8, 6))

        row1 = ctk.CTkFrame(scroll, fg_color="transparent")
        row1.pack(fill="x", padx=10, pady=4)

        # Fecha
        ctk.CTkLabel(row1, text="Fecha:", font=ctk.CTkFont(FONT_BODY, 12)).pack(side="left")
        self.d_fecha = ctk.CTkEntry(row1, width=120, fg_color=C["input"])
        self.d_fecha.insert(0, datetime.datetime.now().strftime("%d-%m-%Y"))
        self.d_fecha.pack(side="left", padx=(6, 20))

        ctk.CTkLabel(row1, text="Hora Inicio:", font=ctk.CTkFont(FONT_BODY, 12)).pack(side="left")
        self.d_hora_ini = ctk.CTkEntry(row1, width=80, fg_color=C["input"])
        self.d_hora_ini.insert(0, "8:00 AM")
        self.d_hora_ini.pack(side="left", padx=(6, 20))

        ctk.CTkLabel(row1, text="Hora Fin:", font=ctk.CTkFont(FONT_BODY, 12)).pack(side="left")
        self.d_hora_fin = ctk.CTkEntry(row1, width=80, fg_color=C["input"])
        self.d_hora_fin.insert(0, "10:00 AM")
        self.d_hora_fin.pack(side="left", padx=6)

        row2 = ctk.CTkFrame(scroll, fg_color="transparent")
        row2.pack(fill="x", padx=10, pady=4)

        ctk.CTkLabel(row2, text="Lugar:", font=ctk.CTkFont(FONT_BODY, 12)).pack(side="left")
        self.d_lugar = ctk.CTkEntry(row2, width=250, fg_color=C["input"])
        self.d_lugar.insert(0, "Sala de Reuniones")
        self.d_lugar.pack(side="left", padx=(6, 20))

        ctk.CTkLabel(row2, text="Acta N°:", font=ctk.CTkFont(FONT_BODY, 12)).pack(side="left")
        self.d_num_acta = ctk.CTkEntry(row2, width=60, fg_color=C["input"])
        self.d_num_acta.insert(0, "001")
        self.d_num_acta.pack(side="left", padx=6)

        # Asistentes
        ctk.CTkLabel(scroll, text="👥 Asistentes (uno por línea):",
                     font=ctk.CTkFont(FONT_BODY, 12, "bold"),
                     text_color=C["texto"]).pack(anchor="w", padx=10, pady=(12, 4))

        self.d_asistentes = ctk.CTkTextbox(scroll, height=80, fg_color=C["input"],
                                            font=(FONT_BODY, 12))
        self.d_asistentes.pack(fill="x", padx=10, pady=2)

        # Temas
        ctk.CTkLabel(scroll, text="📌 Temas Tratados (uno por línea):",
                     font=ctk.CTkFont(FONT_BODY, 12, "bold"),
                     text_color=C["texto"]).pack(anchor="w", padx=10, pady=(12, 4))

        self.d_temas = ctk.CTkTextbox(scroll, height=80, fg_color=C["input"],
                                       font=(FONT_BODY, 12))
        self.d_temas.pack(fill="x", padx=10, pady=2)

        # Acuerdos
        ctk.CTkLabel(scroll, text="✅ Acuerdos y Compromisos (uno por línea):",
                     font=ctk.CTkFont(FONT_BODY, 12, "bold"),
                     text_color=C["texto"]).pack(anchor="w", padx=10, pady=(12, 4))

        self.d_acuerdos = ctk.CTkTextbox(scroll, height=80, fg_color=C["input"],
                                          font=(FONT_BODY, 12))
        self.d_acuerdos.pack(fill="x", padx=10, pady=2)

        # Observaciones
        ctk.CTkLabel(scroll, text="📝 Observaciones:",
                     font=ctk.CTkFont(FONT_BODY, 12)).pack(anchor="w", padx=10, pady=(12, 4))

        self.d_obs = ctk.CTkEntry(scroll, fg_color=C["input"],
                                   placeholder_text="Observaciones adicionales...")
        self.d_obs.pack(fill="x", padx=10, pady=2)

        # Botón generar
        ctk.CTkButton(scroll, text="📄 GENERAR ACTA DE REUNIÓN",
                      font=ctk.CTkFont(FONT_TITLE, 14, "bold"),
                      fg_color="#10B981", hover_color="#059669",
                      height=42, command=self._generar_docentes).pack(
            fill="x", padx=10, pady=(16, 8))

    def _crear_tab_padres(self, parent):
        scroll = ctk.CTkScrollableFrame(parent, fg_color="transparent")
        scroll.pack(fill="both", expand=True, padx=4, pady=4)

        ctk.CTkLabel(scroll, text="👪 Datos de la Citación",
                     font=ctk.CTkFont(FONT_TITLE, 15, "bold"),
                     text_color=C["cian"]).pack(anchor="w", padx=10, pady=(8, 6))

        # Datos del alumno
        row1 = ctk.CTkFrame(scroll, fg_color="transparent")
        row1.pack(fill="x", padx=10, pady=4)

        ctk.CTkLabel(row1, text="Alumno:", font=ctk.CTkFont(FONT_BODY, 12)).pack(side="left")
        grados = self.engine.obtener_grados_activos() or ["Sin datos"]
        self.p_grado = ctk.CTkOptionMenu(row1, values=grados, width=80,
                                          fg_color=C["input"],
                                          command=self._cargar_alumnos)
        self.p_grado.pack(side="left", padx=6)

        self.p_alumno = ctk.CTkOptionMenu(row1, values=["Seleccione grado"],
                                           width=200, fg_color=C["input"])
        self.p_alumno.pack(side="left", padx=6)

        row2 = ctk.CTkFrame(scroll, fg_color="transparent")
        row2.pack(fill="x", padx=10, pady=4)

        ctk.CTkLabel(row2, text="Acudiente:", font=ctk.CTkFont(FONT_BODY, 12)).pack(side="left")
        self.p_acudiente = ctk.CTkEntry(row2, width=200, fg_color=C["input"],
                                         placeholder_text="Nombre del padre/madre")
        self.p_acudiente.pack(side="left", padx=(6, 16))

        ctk.CTkLabel(row2, text="Parentesco:", font=ctk.CTkFont(FONT_BODY, 12)).pack(side="left")
        self.p_parentesco = ctk.CTkOptionMenu(row2,
            values=["Padre", "Madre", "Abuelo/a", "Tío/a", "Tutor Legal", "Otro"],
            width=120, fg_color=C["input"])
        self.p_parentesco.pack(side="left", padx=6)

        row3 = ctk.CTkFrame(scroll, fg_color="transparent")
        row3.pack(fill="x", padx=10, pady=4)

        ctk.CTkLabel(row3, text="Teléfono:", font=ctk.CTkFont(FONT_BODY, 12)).pack(side="left")
        self.p_tel = ctk.CTkEntry(row3, width=140, fg_color=C["input"])
        self.p_tel.pack(side="left", padx=(6, 16))

        ctk.CTkLabel(row3, text="Fecha:", font=ctk.CTkFont(FONT_BODY, 12)).pack(side="left")
        self.p_fecha = ctk.CTkEntry(row3, width=120, fg_color=C["input"])
        self.p_fecha.insert(0, datetime.datetime.now().strftime("%d-%m-%Y"))
        self.p_fecha.pack(side="left", padx=6)

        # Motivo
        ctk.CTkLabel(scroll, text="Motivo de la reunión:",
                     font=ctk.CTkFont(FONT_BODY, 12, "bold")).pack(anchor="w", padx=10, pady=(12, 4))

        self.p_motivo = ctk.CTkOptionMenu(scroll,
            values=["Bajo rendimiento académico", "Conducta inapropiada",
                    "Inasistencias frecuentes", "Situación familiar",
                    "Mérito y reconocimiento", "Adaptación escolar",
                    "Seguimiento de compromisos", "Otro motivo"],
            fg_color=C["input"])
        self.p_motivo.pack(fill="x", padx=10, pady=2)

        # Descripción
        ctk.CTkLabel(scroll, text="Descripción de la situación:",
                     font=ctk.CTkFont(FONT_BODY, 12, "bold")).pack(anchor="w", padx=10, pady=(12, 4))

        self.p_desc = ctk.CTkTextbox(scroll, height=80, fg_color=C["input"],
                                      font=(FONT_BODY, 12))
        self.p_desc.pack(fill="x", padx=10, pady=2)

        # Acuerdos padre
        ctk.CTkLabel(scroll, text="Acuerdos con el padre/madre:",
                     font=ctk.CTkFont(FONT_BODY, 12, "bold")).pack(anchor="w", padx=10, pady=(12, 4))

        self.p_acuerdos = ctk.CTkTextbox(scroll, height=60, fg_color=C["input"],
                                          font=(FONT_BODY, 12))
        self.p_acuerdos.pack(fill="x", padx=10, pady=2)

        # Compromisos alumno
        ctk.CTkLabel(scroll, text="Compromisos del alumno:",
                     font=ctk.CTkFont(FONT_BODY, 12, "bold")).pack(anchor="w", padx=10, pady=(12, 4))

        self.p_compromisos = ctk.CTkTextbox(scroll, height=60, fg_color=C["input"],
                                             font=(FONT_BODY, 12))
        self.p_compromisos.pack(fill="x", padx=10, pady=2)

        # Botón generar
        ctk.CTkButton(scroll, text="📄 GENERAR ACTA DE CITACIÓN",
                      font=ctk.CTkFont(FONT_TITLE, 14, "bold"),
                      fg_color="#10B981", hover_color="#059669",
                      height=42, command=self._generar_padres).pack(
            fill="x", padx=10, pady=(16, 8))

        # Cargar alumnos del primer grado
        self._cargar_alumnos(grados[0])

    def _cargar_alumnos(self, grado):
        try:
            ests = self.engine.obtener_estudiantes_completos(grado)
            nombres = [e["nombre"] for e in ests] if ests else ["Sin alumnos"]
            self.p_alumno.configure(values=nombres)
            self.p_alumno.set(nombres[0])
        except Exception:
            pass

    def _generar_docentes(self):
        try:
            datos_gen = self.engine.obtener_datos_generales()
        except Exception:
            datos_gen = {}

        asist_txt = self.d_asistentes.get("1.0", "end").strip()
        temas_txt = self.d_temas.get("1.0", "end").strip()
        acuerdos_txt = self.d_acuerdos.get("1.0", "end").strip()

        datos = {
            "escuela": datos_gen.get("escuela_nombre", "Centro Educativo"),
            "director": datos_gen.get("director_nombre", ""),
            "fecha": self.d_fecha.get().strip(),
            "hora_inicio": self.d_hora_ini.get().strip(),
            "hora_fin": self.d_hora_fin.get().strip(),
            "lugar": self.d_lugar.get().strip(),
            "numero_acta": self.d_num_acta.get().strip(),
            "ano": datos_gen.get("ano_lectivo", "2026"),
            "convocante": "Dirección",
            "asistentes": [a.strip() for a in asist_txt.split("\n") if a.strip()],
            "temas": [t.strip() for t in temas_txt.split("\n") if t.strip()],
            "acuerdos": [a.strip() for a in acuerdos_txt.split("\n") if a.strip()],
            "observaciones": self.d_obs.get().strip(),
        }

        ruta = generar_acta_docentes(datos)
        if ruta:
            root = self.winfo_toplevel()
            if hasattr(root, "mostrar_toast"):
                root.mostrar_toast("✅ Acta generada con éxito", color="#10B981")
        else:
            messagebox.showerror("Error", "No se pudo generar el acta. Verifique python-docx.")

    def _generar_padres(self):
        try:
            datos_gen = self.engine.obtener_datos_generales()
        except Exception:
            datos_gen = {}

        datos = {
            "escuela": datos_gen.get("escuela_nombre", "Centro Educativo"),
            "docente": datos_gen.get("docente_nombre", ""),
            "fecha": self.p_fecha.get().strip(),
            "hora": datetime.datetime.now().strftime("%I:%M %p"),
            "alumno": self.p_alumno.get(),
            "grado": self.p_grado.get(),
            "acudiente": self.p_acudiente.get().strip(),
            "parentesco": self.p_parentesco.get(),
            "telefono": self.p_tel.get().strip(),
            "motivo": self.p_motivo.get(),
            "descripcion": self.p_desc.get("1.0", "end").strip(),
            "acuerdos_padre": self.p_acuerdos.get("1.0", "end").strip(),
            "acuerdos_alumno": self.p_compromisos.get("1.0", "end").strip(),
        }

        ruta = generar_acta_padres(datos)
        if ruta:
            root = self.winfo_toplevel()
            if hasattr(root, "mostrar_toast"):
                root.mostrar_toast("✅ Acta de citación generada", color="#10B981")
        else:
            messagebox.showerror("Error", "No se pudo generar. Verifique python-docx.")
