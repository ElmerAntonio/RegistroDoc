"""
RegistroDoc Pro — Pantalla de Gestión de Tareas Programadas
Permite al docente crear, ver y gestionar tareas con anticipación.
100% offline.
"""
import customtkinter as ctk
from tkinter import messagebox
import datetime

from theme import C, FONT_TITLE, FONT_BODY
from tareas import (cargar_tareas, agregar_tarea, marcar_completada,
                    eliminar_tarea, obtener_pendientes)


class TareasFrame(ctk.CTkFrame):
    def __init__(self, master, engine, **kwargs):
        super().__init__(master, fg_color="transparent", **kwargs)
        self.engine = engine
        self.editando_tarea_id = None
        self.crear_interfaz()

    def actualizar_vista(self):
        """Refresca la lista de tareas al volver a esta pantalla."""
        self._refrescar_lista()

    def crear_interfaz(self):
        # ═══ HEADER ═══
        hdr = ctk.CTkFrame(self, fg_color=C["card"], corner_radius=12,
                           border_width=1, border_color=C["cian"])
        hdr.pack(fill="x", padx=15, pady=(12, 8))

        ctk.CTkLabel(hdr, text="📋 Tareas y Evaluaciones Programadas",
                     font=ctk.CTkFont(FONT_TITLE, 20, "bold"),
                     text_color=C["cian"]).pack(side="left", padx=18, pady=12)

        ctk.CTkLabel(hdr, text="Programe trabajos con anticipación y reciba recordatorios",
                     font=ctk.CTkFont(FONT_BODY, 11),
                     text_color=C["texto_sec"]).pack(side="left", padx=10, pady=12)

        # ═══ CUERPO: 2 COLUMNAS ═══
        body = ctk.CTkFrame(self, fg_color="transparent")
        body.pack(fill="both", expand=True, padx=15, pady=4)
        body.columnconfigure(0, weight=4)
        body.columnconfigure(1, weight=6)
        body.rowconfigure(0, weight=1)

        # ─── Columna Izquierda: Formulario ───
        self._crear_formulario(body)

        # ─── Columna Derecha: Lista de Tareas ───
        self._crear_lista(body)

    def _crear_formulario(self, parent):
        form = ctk.CTkFrame(parent, fg_color=C["card"], corner_radius=12,
                            border_width=1, border_color=C["borde"])
        form.grid(row=0, column=0, sticky="nsew", padx=(0, 8))

        self.lbl_form_title = ctk.CTkLabel(form, text="➕ Nueva Tarea",
                                            font=ctk.CTkFont(FONT_TITLE, 16, "bold"),
                                            text_color=C["cian"])
        self.lbl_form_title.pack(anchor="w", padx=16, pady=(14, 8))

        # Título
        ctk.CTkLabel(form, text="Título de la tarea:",
                      font=ctk.CTkFont(FONT_BODY, 12),
                      text_color=C["texto"]).pack(anchor="w", padx=16, pady=(8, 2))
        self.entry_titulo = ctk.CTkEntry(form, fg_color=C["input"],
                                          border_color=C["borde"],
                                          placeholder_text="Ej: Examen de Matemáticas")
        self.entry_titulo.pack(fill="x", padx=16, pady=(0, 6))

        # Grados (Selección Múltiple)
        ctk.CTkLabel(form, text="Grados/Grupos:",
                      font=ctk.CTkFont(FONT_BODY, 12),
                      text_color=C["texto"]).pack(anchor="w", padx=16, pady=(4, 2))
        
        self.var_todos = ctk.BooleanVar(value=False)
        self.chk_todos = ctk.CTkCheckBox(form, text="Seleccionar todos", variable=self.var_todos,
                                         font=ctk.CTkFont(FONT_BODY, 11), command=self._toggle_todos)
        self.chk_todos.pack(anchor="w", padx=16, pady=2)
        
        grados = self.engine.obtener_grados_activos() or ["Sin datos"]
        grados_frame = ctk.CTkScrollableFrame(form, height=75, fg_color=C["input"],
                                             border_color=C["borde"], border_width=1)
        grados_frame.pack(fill="x", padx=16, pady=(0, 6))
        
        self.grado_checkboxes = {}
        for g in grados:
            chk = ctk.CTkCheckBox(grados_frame, text=g, font=ctk.CTkFont(FONT_BODY, 11))
            chk.pack(anchor="w", padx=5, pady=2)
            self.grado_checkboxes[g] = chk

        # Materia (Lista unificada)
        ctk.CTkLabel(form, text="Materia:",
                      font=ctk.CTkFont(FONT_BODY, 12),
                      text_color=C["texto"]).pack(anchor="w", padx=16, pady=(4, 2))
        materias = ["General"]
        try:
            for g in grados:
                if g != "Sin datos":
                    mats = self.engine.obtener_materias_por_grado(g)
                    if mats:
                        for m in mats:
                            if m not in materias and m not in ["No hay materias", "No hay materias registradas"]:
                                materias.append(m)
        except Exception:
            pass
        self.combo_materia = ctk.CTkOptionMenu(form, values=materias,
                                                fg_color=C["input"])
        self.combo_materia.pack(fill="x", padx=16, pady=(0, 6))

        # Tipo
        ctk.CTkLabel(form, text="Tipo de evaluación:",
                      font=ctk.CTkFont(FONT_BODY, 12),
                      text_color=C["texto"]).pack(anchor="w", padx=16, pady=(4, 2))
        self.combo_tipo = ctk.CTkOptionMenu(form,
            values=["Parcial", "Apreciación", "Examen", "Asistencia", "Hábitos", "Otro"],
            fg_color=C["input"])
        self.combo_tipo.pack(fill="x", padx=16, pady=(0, 6))

        # Fecha límite
        ctk.CTkLabel(form, text="Fecha límite (DD-MM-YYYY):",
                      font=ctk.CTkFont(FONT_BODY, 12),
                      text_color=C["texto"]).pack(anchor="w", padx=16, pady=(4, 2))
        self.entry_fecha = ctk.CTkEntry(form, fg_color=C["input"],
                                         border_color=C["borde"],
                                         placeholder_text="Ej: 15-06-2026")
        manana = (datetime.datetime.now() + datetime.timedelta(days=1)).strftime("%d-%m-%Y")
        self.entry_fecha.insert(0, manana)
        self.entry_fecha.pack(fill="x", padx=16, pady=(0, 6))

        # Descripción
        ctk.CTkLabel(form, text="Descripción (opcional):",
                      font=ctk.CTkFont(FONT_BODY, 12),
                      text_color=C["texto"]).pack(anchor="w", padx=16, pady=(4, 2))
        self.entry_desc = ctk.CTkEntry(form, fg_color=C["input"],
                                        border_color=C["borde"],
                                        placeholder_text="Detalles adicionales...")
        self.entry_desc.pack(fill="x", padx=16, pady=(0, 12))

        # Botones de Formulario
        self.buttons_frame = ctk.CTkFrame(form, fg_color="transparent")
        self.buttons_frame.pack(fill="x", padx=16, pady=(8, 16))
        
        self.btn_guardar = ctk.CTkButton(self.buttons_frame, text="💾 PROGRAMAR TAREA",
                                          font=ctk.CTkFont(FONT_TITLE, 14, "bold"),
                                          fg_color="#10B981", hover_color="#059669",
                                          height=42, command=self._guardar_tarea)
        self.btn_guardar.pack(side="left", fill="x", expand=True)
        
        self.btn_cancelar = ctk.CTkButton(self.buttons_frame, text="❌ CANCELAR",
                                           font=ctk.CTkFont(FONT_TITLE, 12, "bold"),
                                           fg_color="#475569", hover_color="#334155",
                                           height=42, command=self._cancelar_edicion)

    def _crear_lista(self, parent):
        lista_frame = ctk.CTkFrame(parent, fg_color=C["card"], corner_radius=12,
                                    border_width=1, border_color=C["borde"])
        lista_frame.grid(row=0, column=1, sticky="nsew")

        # Header de la lista
        lhdr = ctk.CTkFrame(lista_frame, fg_color="transparent")
        lhdr.pack(fill="x", padx=16, pady=(14, 8))

        ctk.CTkLabel(lhdr, text="📅 Tareas Programadas",
                     font=ctk.CTkFont(FONT_TITLE, 16, "bold"),
                     text_color=C["cian"]).pack(side="left")

        ctk.CTkButton(lhdr, text="🔄 Refrescar", width=100, height=28,
                      fg_color=C["badge_bg"], hover_color=C["hover"],
                      font=ctk.CTkFont(FONT_BODY, 11),
                      command=self._refrescar_lista).pack(side="right")

        # Scroll de tareas
        self.scroll_tareas = ctk.CTkScrollableFrame(lista_frame, fg_color="transparent",
                                                      scrollbar_button_color=C["borde"])
        self.scroll_tareas.pack(fill="both", expand=True, padx=8, pady=(0, 8))

        self._refrescar_lista()

    def _refrescar_lista(self):
        for w in self.scroll_tareas.winfo_children():
            w.destroy()

        pendientes = obtener_pendientes()
        completadas = [t for t in cargar_tareas() if t.get("completada")]

        if not pendientes and not completadas:
            ctk.CTkLabel(self.scroll_tareas,
                         text="📭 No hay tareas programadas.\nUse el formulario para crear una.",
                         font=ctk.CTkFont(FONT_BODY, 13),
                         text_color=C["texto_sec"]).pack(pady=40)
            return

        # Pendientes
        if pendientes:
            ctk.CTkLabel(self.scroll_tareas, text="⏳ Pendientes",
                         font=ctk.CTkFont(FONT_TITLE, 13, "bold"),
                         text_color=C["amarillo"]).pack(anchor="w", padx=8, pady=(4, 4))

            for t in pendientes:
                self._renderizar_tarea(t, pendiente=True)

        # Completadas (últimas 5)
        if completadas:
            ctk.CTkLabel(self.scroll_tareas, text="✅ Completadas",
                         font=ctk.CTkFont(FONT_TITLE, 13, "bold"),
                         text_color=C["verde"]).pack(anchor="w", padx=8, pady=(12, 4))

            for t in completadas[-5:]:
                self._renderizar_tarea(t, pendiente=False)

    def _toggle_todos(self):
        val = self.var_todos.get()
        for chk in self.grado_checkboxes.values():
            if val:
                chk.select()
            else:
                chk.deselect()

    def _iniciar_edicion(self, tarea):
        self.editando_tarea_id = tarea["id"]
        
        # Change headers and buttons
        self.lbl_form_title.configure(text="✏️ Editar Tarea")
        self.btn_guardar.configure(text="💾 GUARDAR CAMBIOS")
        self.btn_cancelar.pack(side="right", fill="x", expand=True, padx=(8, 0))
        
        # Fill entry fields
        self.entry_titulo.delete(0, "end")
        self.entry_titulo.insert(0, tarea.get("titulo", ""))
        
        self.entry_fecha.delete(0, "end")
        self.entry_fecha.insert(0, tarea.get("fecha_limite", ""))
        
        self.entry_desc.delete(0, "end")
        self.entry_desc.insert(0, tarea.get("descripcion", ""))
        
        self.combo_tipo.set(tarea.get("tipo", "Otro"))
        self.combo_materia.set(tarea.get("materia", "General"))
        
        # Check corresponding grade and uncheck others
        target_grade = tarea.get("grado", "")
        for g, chk in self.grado_checkboxes.items():
            if g == target_grade:
                chk.select()
            else:
                chk.deselect()
        self.var_todos.set(False)

    def _cancelar_edicion(self):
        self.editando_tarea_id = None
        
        # Reset form labels and buttons
        self.lbl_form_title.configure(text="➕ Nueva Tarea")
        self.btn_guardar.configure(text="💾 PROGRAMAR TAREA")
        self.btn_cancelar.pack_forget()
        
        # Clear entries
        self.entry_titulo.delete(0, "end")
        self.entry_desc.delete(0, "end")
        manana = (datetime.datetime.now() + datetime.timedelta(days=1)).strftime("%d-%m-%Y")
        self.entry_fecha.delete(0, "end")
        self.entry_fecha.insert(0, manana)
        
        # Uncheck grades
        for chk in self.grado_checkboxes.values():
            chk.deselect()
        self.var_todos.set(False)

    def _renderizar_tarea(self, tarea, pendiente=True):
        urgencia = tarea.get("_urgencia", "normal")
        colores_borde = {
            "vencida": "#EF4444", "hoy": "#F59E0B",
            "urgente": "#FB923C", "normal": C["borde"]
        }
        borde = colores_borde.get(urgencia, C["borde"]) if pendiente else "#334155"

        card = ctk.CTkFrame(self.scroll_tareas, fg_color=C["card_alt"],
                            corner_radius=8, border_width=1, border_color=borde)
        card.pack(fill="x", pady=3, padx=4)

        inner = ctk.CTkFrame(card, fg_color="transparent")
        inner.pack(fill="x", padx=12, pady=8)

        # Info
        info_f = ctk.CTkFrame(inner, fg_color="transparent")
        info_f.pack(side="left", fill="x", expand=True)

        titulo_color = C["texto"] if pendiente else C["texto_sec"]
        ctk.CTkLabel(info_f, text=tarea["titulo"],
                     font=ctk.CTkFont(FONT_BODY, 13, "bold"),
                     text_color=titulo_color, anchor="w").pack(anchor="w")

        detalle = f"{tarea.get('tipo', '')} — {tarea.get('grado', '')} — {tarea.get('materia', '')} — {tarea.get('fecha_limite', '')}"
        ctk.CTkLabel(info_f, text=detalle,
                     font=ctk.CTkFont(FONT_BODY, 10),
                     text_color=C["texto_sec"], anchor="w").pack(anchor="w")

        # Botones de acción
        if pendiente:
            btn_f = ctk.CTkFrame(inner, fg_color="transparent")
            btn_f.pack(side="right")

            tid = tarea["id"]
            ctk.CTkButton(btn_f, text="✅", width=32, height=28,
                          fg_color="#10B981", hover_color="#059669",
                          font=ctk.CTkFont(FONT_BODY, 14),
                          command=lambda i=tid: self._completar(i)).pack(side="left", padx=2)

            ctk.CTkButton(btn_f, text="✏️", width=32, height=28,
                          fg_color="#6366F1", hover_color="#4F46E5",
                          font=ctk.CTkFont(FONT_BODY, 12),
                          command=lambda t=tarea: self._iniciar_edicion(t)).pack(side="left", padx=2)

            ctk.CTkButton(btn_f, text="🗑️", width=32, height=28,
                          fg_color="#EF4444", hover_color="#DC2626",
                          font=ctk.CTkFont(FONT_BODY, 14),
                          command=lambda i=tid: self._eliminar(i)).pack(side="left", padx=2)

    def _guardar_tarea(self):
        titulo = self.entry_titulo.get().strip()
        if not titulo:
            messagebox.showwarning("Atención", "Escriba un título para la tarea.")
            return

        fecha = self.entry_fecha.get().strip()
        if not fecha:
            messagebox.showwarning("Atención", "Escriba una fecha límite.")
            return

        import datetime
        try:
            partes = fecha.split("-")
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
                    f"La fecha límite ({fecha}) está fuera del período escolar.\n\n"
                    f"Rango escolar: {t1_start.strftime('%d-%m-%Y')} a {t3_end.strftime('%d-%m-%Y')}.\n"
                    "Por favor ingrese una fecha dentro de este rango."
                )
                return
        except Exception:
            pass

        # Check which checkboxes are checked
        grados_seleccionados = [g for g, chk in self.grado_checkboxes.items() if chk.get()]
        if not grados_seleccionados:
            messagebox.showwarning("Atención", "Seleccione al menos un Grado/Grupo.")
            return

        materia = self.combo_materia.get()
        tipo = self.combo_tipo.get()
        desc = self.entry_desc.get().strip()

        if self.editando_tarea_id is not None:
            # Edit mode: Delete old task first
            eliminar_tarea(self.editando_tarea_id)
            self.editando_tarea_id = None
            
        for g in grados_seleccionados:
            agregar_tarea(
                titulo=titulo,
                grado=g,
                materia=materia,
                tipo=tipo,
                fecha_limite=fecha,
                descripcion=desc
            )

        # Clear formulario
        self.entry_titulo.delete(0, "end")
        self.entry_desc.delete(0, "end")
        manana = (datetime.datetime.now() + datetime.timedelta(days=1)).strftime("%d-%m-%Y")
        self.entry_fecha.delete(0, "end")
        self.entry_fecha.insert(0, manana)
        
        for chk in self.grado_checkboxes.values():
            chk.deselect()
        self.var_todos.set(False)

        # Change form back to normal
        self.lbl_form_title.configure(text="➕ Nueva Tarea")
        self.btn_guardar.configure(text="💾 PROGRAMAR TAREA")
        self.btn_cancelar.pack_forget()

        # Refrescar lista
        self._refrescar_lista()

        # Toast
        root = self.winfo_toplevel()
        if hasattr(root, "mostrar_toast"):
            root.mostrar_toast("✅ Tarea(s) guardadas con éxito", color="#10B981")

    def _completar(self, tarea_id):
        marcar_completada(tarea_id)
        self._refrescar_lista()
        root = self.winfo_toplevel()
        if hasattr(root, "mostrar_toast"):
            root.mostrar_toast("✅ Tarea marcada como completada", color="#10B981")

    def _eliminar(self, tarea_id):
        if messagebox.askyesno("Confirmar", "¿Eliminar esta tarea?"):
            eliminar_tarea(tarea_id)
            self._refrescar_lista()
