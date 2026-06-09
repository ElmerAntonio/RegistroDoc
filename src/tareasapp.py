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

        ctk.CTkLabel(form, text="➕ Nueva Tarea",
                     font=ctk.CTkFont(FONT_TITLE, 16, "bold"),
                     text_color=C["cian"]).pack(anchor="w", padx=16, pady=(14, 8))

        # Título
        ctk.CTkLabel(form, text="Título de la tarea:",
                     font=ctk.CTkFont(FONT_BODY, 12),
                     text_color=C["texto"]).pack(anchor="w", padx=16, pady=(8, 2))
        self.entry_titulo = ctk.CTkEntry(form, fg_color=C["input"],
                                          border_color=C["borde"],
                                          placeholder_text="Ej: Examen de Matemáticas")
        self.entry_titulo.pack(fill="x", padx=16, pady=(0, 6))

        # Grado
        ctk.CTkLabel(form, text="Grado:",
                     font=ctk.CTkFont(FONT_BODY, 12),
                     text_color=C["texto"]).pack(anchor="w", padx=16, pady=(4, 2))
        grados = self.engine.obtener_grados_activos() or ["Sin datos"]
        self.combo_grado = ctk.CTkOptionMenu(form, values=grados,
                                              fg_color=C["input"])
        self.combo_grado.pack(fill="x", padx=16, pady=(0, 6))

        # Materia
        ctk.CTkLabel(form, text="Materia:",
                     font=ctk.CTkFont(FONT_BODY, 12),
                     text_color=C["texto"]).pack(anchor="w", padx=16, pady=(4, 2))
        materias = ["General"]
        try:
            if grados and grados[0] != "Sin datos":
                materias = self.engine.obtener_materias_por_grado(grados[0])
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
        # Poner fecha de mañana por defecto
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

        # Botón guardar
        ctk.CTkButton(form, text="💾 PROGRAMAR TAREA",
                      font=ctk.CTkFont(FONT_TITLE, 14, "bold"),
                      fg_color="#10B981", hover_color="#059669",
                      height=42, command=self._guardar_tarea).pack(
            fill="x", padx=16, pady=(8, 16))

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

        agregar_tarea(
            titulo=titulo,
            grado=self.combo_grado.get(),
            materia=self.combo_materia.get(),
            tipo=self.combo_tipo.get(),
            fecha_limite=fecha,
            descripcion=self.entry_desc.get().strip()
        )

        # Limpiar formulario
        self.entry_titulo.delete(0, "end")
        self.entry_desc.delete(0, "end")

        # Refrescar lista
        self._refrescar_lista()

        # Toast
        root = self.winfo_toplevel()
        if hasattr(root, "mostrar_toast"):
            root.mostrar_toast("✅ Tarea programada con éxito", color="#10B981")

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
