import customtkinter as ctk
import tkinter.messagebox as messagebox
import tkinter as tk
from theme import C, FONT_TITLE, FONT_BODY

class EstudiantesFrame(ctk.CTkFrame):
    def __init__(self, master, engine, **kwargs):
        super().__init__(master, fg_color="transparent", **kwargs)
        self.engine = engine
        self.entradas_estudiantes = {}

        ctk.CTkLabel(self, text="Gestión de Estudiantes",
                     font=(FONT_TITLE, 24, "bold")).pack(anchor="w", pady=(0, 10))

        # ── PANEL AGREGAR ──────────────────────────────────────────────────
        frame_agregar = ctk.CTkFrame(self, fg_color=C["success"], corner_radius=8)
        frame_agregar.pack(fill="x", pady=5, ipadx=10, ipady=10)

        ctk.CTkLabel(frame_agregar, text="➕ Agregar Alumno:",
                     font=(FONT_TITLE, 14, "bold"),
                     text_color=C["fondo"]).pack(side="left", padx=10)

        self.entry_nuevo_nombre = ctk.CTkEntry(frame_agregar, width=250,
                                               placeholder_text="APELLIDO, Nombre (Requerido)")
        self.entry_nuevo_nombre.pack(side="left", padx=10)

        self.entry_nueva_cedula = ctk.CTkEntry(frame_agregar, width=120,
                                               placeholder_text="Cédula (Opcional)")
        self.entry_nueva_cedula.pack(side="left", padx=10)
        self.entry_nueva_cedula.bind("<FocusOut>", lambda e: self._formatear_nueva_cedula())

        ctk.CTkLabel(frame_agregar, text="Sexo:", font=(FONT_BODY, 12),
                     text_color=C["fondo"]).pack(side="left", padx=(5, 2))
        self.combo_nuevo_sexo = ctk.CTkOptionMenu(frame_agregar, values=["M", "F"], width=60)
        self.combo_nuevo_sexo.set("M")
        self.combo_nuevo_sexo.pack(side="left", padx=10)

        ctk.CTkButton(frame_agregar, text="Guardar Nuevo",
                      fg_color=C["esmeralda"], hover_color=C["verde"],
                      text_color=C["fondo"],
                      command=self.agregar_nuevo).pack(side="left", padx=10)

        # ── SELECTOR DE GRADO ──────────────────────────────────────────────
        frame_controles = ctk.CTkFrame(self, fg_color=C["card_alt"], corner_radius=8)
        frame_controles.pack(fill="x", pady=10, ipadx=10, ipady=10)

        ctk.CTkLabel(frame_controles, text="Seleccione Grado:",
                     font=(FONT_BODY, 14)).pack(side="left", padx=10)

        opciones_grado = self.engine.obtener_grados_activos() or ["Sin datos"]
        self.combo_grado = ctk.CTkOptionMenu(frame_controles, values=opciones_grado,
                                             command=self._al_cambiar_grado)
        self.combo_grado.pack(side="left", padx=10)

        ctk.CTkButton(frame_controles, text="🖨️ Lista de Clase",
                      fg_color="#6366F1", hover_color="#4F46E5",
                      font=(FONT_BODY, 12, "bold"), height=32,
                      command=self.imprimir_lista_clase).pack(side="right", padx=15)

        # ── TABS: Activos / Retirados ───────────────────────────────────────
        self.tabs = ctk.CTkTabview(self, corner_radius=8)
        self.tabs.pack(fill="both", expand=True, pady=5)
        self.tab_activos   = self.tabs.add("👥 Activos")
        self.tab_retirados = self.tabs.add("🚪 Retirados")
        self.tabs.set("👥 Activos")

        # ── Tab Activos ────────────────────────────────────────────────────
        frame_enc = ctk.CTkFrame(self.tab_activos, fg_color=C["badge_bg"], corner_radius=5)
        frame_enc.pack(fill="x", pady=(5, 0), ipady=5)
        ctk.CTkLabel(frame_enc, text="N°",   width=40, font=(FONT_BODY, 13, "bold")).pack(side="left", padx=6)
        ctk.CTkLabel(frame_enc, text="APELLIDO, NOMBRE", width=320, anchor="w",
                     font=(FONT_BODY, 13, "bold")).pack(side="left", padx=6)
        ctk.CTkLabel(frame_enc, text="CÉDULA", width=130, anchor="w",
                     font=(FONT_BODY, 13, "bold")).pack(side="left", padx=6)
        ctk.CTkLabel(frame_enc, text="SEXO", width=70, anchor="w",
                     font=(FONT_BODY, 13, "bold")).pack(side="left", padx=6)
        ctk.CTkLabel(frame_enc, text="ACCIÓN", width=100, anchor="w",
                     font=(FONT_BODY, 13, "bold")).pack(side="left", padx=6)

        self.scroll_activos = ctk.CTkScrollableFrame(self.tab_activos, fg_color="transparent")
        self.scroll_activos.pack(fill="both", expand=True, pady=3)

        self.btn_guardar = ctk.CTkButton(
            self.tab_activos,
            text="💾 GUARDAR MODIFICACIONES DE LA LISTA",
            fg_color="#3B82F6", hover_color="#2563EB",
            font=(FONT_BODY, 14, "bold"), height=38,
            command=self.guardar_cambios
        )
        self.btn_guardar.pack(pady=8)

        # ── Tab Retirados ──────────────────────────────────────────────────
        frame_enc_r = ctk.CTkFrame(self.tab_retirados, fg_color=C["badge_bg"], corner_radius=5)
        frame_enc_r.pack(fill="x", pady=(5, 0), ipady=5)
        ctk.CTkLabel(frame_enc_r, text="N°",      width=35,  font=(FONT_BODY, 13, "bold")).pack(side="left", padx=5)
        ctk.CTkLabel(frame_enc_r, text="NOMBRE",  width=250, anchor="w",
                     font=(FONT_BODY, 13, "bold")).pack(side="left", padx=5)
        ctk.CTkLabel(frame_enc_r, text="GRADO",   width=60,  anchor="w",
                     font=(FONT_BODY, 13, "bold")).pack(side="left", padx=5)
        ctk.CTkLabel(frame_enc_r, text="FECHA",   width=90,  anchor="w",
                     font=(FONT_BODY, 13, "bold")).pack(side="left", padx=5)
        ctk.CTkLabel(frame_enc_r, text="MOTIVO",  width=180, anchor="w",
                     font=(FONT_BODY, 13, "bold")).pack(side="left", padx=5)
        ctk.CTkLabel(frame_enc_r, text="ACCIÓN",  width=110, anchor="w",
                     font=(FONT_BODY, 13, "bold")).pack(side="left", padx=5)

        self.scroll_retirados = ctk.CTkScrollableFrame(self.tab_retirados, fg_color="transparent")
        self.scroll_retirados.pack(fill="both", expand=True, pady=3)

        # Datos iniciales se cargan en actualizar_vista() al mostrar el frame
        self._initialized = False

    # ──────────────────────────────────────────────────────────────────────
    # Navegación y carga
    # ──────────────────────────────────────────────────────────────────────

    def _al_cambiar_grado(self, grado):
        self.cargar_lista(grado)
        self.cargar_retirados(grado)

    # ── Pool de filas recicladas (Activos) ────────────────────────────────
    def _obtener_fila_activo(self, index):
        if not hasattr(self, "_pool_activos"):
            self._pool_activos = []
        if index < len(self._pool_activos):
            f_row, widgets = self._pool_activos[index]
            f_row.pack(fill="x", pady=2)
            return f_row, widgets

        f_row = ctk.CTkFrame(self.scroll_activos, fg_color=C["card"], corner_radius=5)
        f_row.pack(fill="x", pady=2)

        num_lbl = ctk.CTkLabel(f_row, text="", width=40)
        num_lbl.pack(side="left", padx=6)

        entry_nom = ctk.CTkEntry(f_row, width=320, fg_color=C["input"])
        entry_nom.pack(side="left", padx=6)

        entry_ced = ctk.CTkEntry(f_row, width=130, fg_color=C["input"],
                                 placeholder_text="Ej: 4-123-456")
        entry_ced.pack(side="left", padx=6)

        combo_sex = ctk.CTkOptionMenu(f_row, values=["M", "F"], width=70)
        combo_sex.pack(side="left", padx=6)

        btn_retirar = ctk.CTkButton(
            f_row, text="🚪 Retirar",
            width=95, height=26,
            fg_color="#F97316", hover_color="#EA580C",
            font=(FONT_BODY, 11, "bold")
        )
        btn_retirar.pack(side="left", padx=6)

        widgets = {
            "num": num_lbl, "nombre": entry_nom, "cedula": entry_ced,
            "sexo": combo_sex, "btn": btn_retirar
        }
        self._pool_activos.append((f_row, widgets))
        return f_row, widgets

    def _limpiar_scroll_activos(self):
        pool_frames = set()
        if hasattr(self, "_pool_activos"):
            for f_row, _ in self._pool_activos:
                pool_frames.add(f_row)
                f_row.pack_forget()
        for w in self.scroll_activos.winfo_children():
            if w not in pool_frames:
                w.destroy()

    def cargar_lista(self, grado):
        self._limpiar_scroll_activos()
        self.entradas_estudiantes.clear()

        estudiantes = self.engine.obtener_estudiantes_completos(grado)

        if not estudiantes:
            ctk.CTkLabel(self.scroll_activos,
                         text=f"La lista de {grado} está vacía. ¡Agrega un alumno arriba ☝️!",
                         text_color="#94A3B8").pack(pady=20)
            return

        for i, est in enumerate(estudiantes):
            f_row, widgets = self._obtener_fila_activo(i)

            widgets["num"].configure(text=str(i + 1))

            widgets["nombre"].delete(0, "end")
            widgets["nombre"].insert(0, est["nombre"])

            widgets["cedula"].delete(0, "end")
            if est["cedula"]:
                widgets["cedula"].insert(0, est["cedula"])
            widgets["cedula"].bind("<FocusOut>", lambda e, ent=widgets["cedula"]: self._formatear_cedula_existente(ent))

            widgets["sexo"].set(est.get("sexo", "M") if est.get("sexo") in ["M", "F"] else "M")

            widgets["btn"].configure(
                command=lambda eid=est["id"], enom=est["nombre"]: self._iniciar_retiro(eid, enom)
            )

            self.entradas_estudiantes[est["id"]] = {
                "nombre": widgets["nombre"], "cedula": widgets["cedula"], "sexo": widgets["sexo"]
            }

        # Ocultar filas sobrantes del pool
        if hasattr(self, "_pool_activos"):
            for idx in range(len(estudiantes), len(self._pool_activos)):
                self._pool_activos[idx][0].pack_forget()

    # ── Pool de filas recicladas (Retirados) ──────────────────────────────
    def _obtener_fila_retirado(self, index):
        if not hasattr(self, "_pool_retirados"):
            self._pool_retirados = []
        if index < len(self._pool_retirados):
            f_row, widgets = self._pool_retirados[index]
            f_row.pack(fill="x", pady=2)
            return f_row, widgets

        f_row = ctk.CTkFrame(self.scroll_retirados, fg_color="#3B1F1F", corner_radius=5)
        f_row.pack(fill="x", pady=2)

        num_lbl = ctk.CTkLabel(f_row, text="", width=35, text_color="#F87171")
        num_lbl.pack(side="left", padx=5)
        nom_lbl = ctk.CTkLabel(f_row, text="", width=250, anchor="w", text_color="#FCA5A5")
        nom_lbl.pack(side="left", padx=5)
        grado_lbl = ctk.CTkLabel(f_row, text="", width=60, anchor="w", text_color="#94A3B8")
        grado_lbl.pack(side="left", padx=5)
        fecha_lbl = ctk.CTkLabel(f_row, text="", width=90, anchor="w", text_color="#94A3B8")
        fecha_lbl.pack(side="left", padx=5)
        motivo_lbl = ctk.CTkLabel(f_row, text="", width=180, anchor="w", text_color="#94A3B8")
        motivo_lbl.pack(side="left", padx=5)

        btn_reactivar = ctk.CTkButton(
            f_row, text="✅ Reactivar",
            width=105, height=26,
            fg_color="#16A34A", hover_color="#15803D",
            font=(FONT_BODY, 11, "bold")
        )
        btn_reactivar.pack(side="left", padx=5)

        widgets = {
            "num": num_lbl, "nombre": nom_lbl, "grado": grado_lbl,
            "fecha": fecha_lbl, "motivo": motivo_lbl, "btn": btn_reactivar
        }
        self._pool_retirados.append((f_row, widgets))
        return f_row, widgets

    def _limpiar_scroll_retirados(self):
        pool_frames = set()
        if hasattr(self, "_pool_retirados"):
            for f_row, _ in self._pool_retirados:
                pool_frames.add(f_row)
                f_row.pack_forget()
        for w in self.scroll_retirados.winfo_children():
            if w not in pool_frames:
                w.destroy()

    def cargar_retirados(self, grado):
        self._limpiar_scroll_retirados()

        retirados = self.engine.obtener_estudiantes_retirados(grado)

        if not retirados:
            ctk.CTkLabel(self.scroll_retirados,
                         text="No hay estudiantes retirados en este grado.",
                         text_color="#94A3B8").pack(pady=20)
            return

        for i, est in enumerate(retirados):
            f_row, widgets = self._obtener_fila_retirado(i)
            f_row.configure(fg_color="#3B1F1F" if i % 2 == 0 else "#2D1A1A")

            widgets["num"].configure(text=str(i + 1))
            widgets["nombre"].configure(text=est["nombre"])
            widgets["grado"].configure(text=est["grado"])
            widgets["fecha"].configure(text=est["fecha_retiro"])
            motivo = est["motivo_retiro"]
            widgets["motivo"].configure(
                text=motivo[:30] + "…" if len(motivo) > 30 else motivo
            )

            widgets["btn"].configure(
                command=lambda eid=est["id"], enom=est["nombre"]: self._reactivar(eid, enom)
            )

        # Ocultar filas sobrantes del pool de retirados
        if hasattr(self, "_pool_retirados"):
            for idx in range(len(retirados), len(self._pool_retirados)):
                self._pool_retirados[idx][0].pack_forget()

    # ──────────────────────────────────────────────────────────────────────
    # Retiro y reactivación
    # ──────────────────────────────────────────────────────────────────────

    def _iniciar_retiro(self, est_id, nombre):
        """Muestra el diálogo de retiro y ejecuta el proceso si se confirma."""
        dialogo = _DialogoRetiro(self.winfo_toplevel(), nombre)
        self.wait_window(dialogo)

        if not dialogo.confirmado:
            return

        motivo     = dialogo.motivo
        acudiente  = dialogo.acudiente

        exito, resultado = self.engine.retirar_estudiante(est_id, motivo, acudiente)
        root = self.winfo_toplevel()

        if exito:
            # Generar expediente de retiro
            try:
                self._generar_expediente_retiro(est_id, nombre, motivo, acudiente)
            except Exception:
                pass

            if hasattr(root, "mostrar_toast"):
                root.mostrar_toast(
                    f"🚪 {resultado} retirado. Expediente actualizado.",
                    color="#F97316"
                )
            self.cargar_lista(self.combo_grado.get())
            self.cargar_retirados(self.combo_grado.get())
        else:
            messagebox.showerror("Error", f"No se pudo retirar al estudiante.\n{resultado}")

    def _reactivar(self, est_id, nombre):
        if not messagebox.askyesno(
            "Reactivar Estudiante",
            f"¿Desea reactivar a {nombre}?\n\nEl estudiante volverá a aparecer en todas las listas activas."
        ):
            return

        exito, resultado = self.engine.reactivar_estudiante(est_id)
        root = self.winfo_toplevel()

        if exito:
            if hasattr(root, "mostrar_toast"):
                root.mostrar_toast(f"✅ {resultado} reactivado.", color="#10B981")
            self.cargar_lista(self.combo_grado.get())
            self.cargar_retirados(self.combo_grado.get())
        else:
            messagebox.showerror("Error", f"No se pudo reactivar.\n{resultado}")

    def _generar_expediente_retiro(self, est_id, nombre, motivo, acudiente):
        """Genera el expediente Word de retiro (delega a documentos_maestro)."""
        try:
            from rdsecurity import cargar_config_segura
            from documentos_maestro import generar_expediente_retiro
            cfg = cargar_config_segura({})
            grado = self.combo_grado.get()

            # Obtener resumen de asistencia y notas del estudiante
            asis = self.engine.obtener_estadisticas_asistencia(grado, "Todos", est_id)
            notas_t1 = self.engine.obtener_notas_estudiante(nombre, grado, "Trimestre 1")
            notas_t2 = self.engine.obtener_notas_estudiante(nombre, grado, "Trimestre 2")
            notas_t3 = self.engine.obtener_notas_estudiante(nombre, grado, "Trimestre 3")

            datos = {
                "nombre":           nombre,
                "grado":            grado,
                "motivo":           motivo,
                "acudiente":        acudiente,
                "docente_nombre":   cfg.get("docente_nombre", ""),
                "escuela_nombre":   cfg.get("escuela_nombre", ""),
                "ano_lectivo":      cfg.get("ano_lectivo", "2026"),
                "asistencia":       asis,
                "notas_t1":         notas_t1,
                "notas_t2":         notas_t2,
                "notas_t3":         notas_t3,
            }
            # Obtener cédula del estudiante
            cursor = self.engine.db_conn.cursor()
            cursor.execute("SELECT cedula FROM estudiantes WHERE id = ?;", (str(est_id),))
            r = cursor.fetchone()
            datos["cedula"] = self.engine.db_manager.desencriptar_campo(r[0]) if r and r[0] else ""

            generar_expediente_retiro(datos)
        except Exception as e:
            print(f"[dapp] No se pudo generar expediente de retiro: {e}")

    # ──────────────────────────────────────────────────────────────────────
    # Agregar / Guardar
    # ──────────────────────────────────────────────────────────────────────

    def agregar_nuevo(self):
        grado  = self.combo_grado.get()
        nombre = self.entry_nuevo_nombre.get().strip().upper()
        cedula = self.formatear_cedula_panamena(self.entry_nueva_cedula.get().strip())
        sexo   = self.combo_nuevo_sexo.get()

        if not nombre:
            messagebox.showwarning("Faltan datos", "El Apellido y Nombre son obligatorios.")
            return

        max_est = 34 if self.engine.modalidad == "primaria" else 36
        actuales = self.engine.obtener_estudiantes_completos(grado)
        if len(actuales) >= max_est:
            messagebox.showwarning(
                "Límite Alcanzado",
                f"El grado {grado} ha alcanzado el límite máximo de {max_est} estudiantes "
                f"permitido para {self.engine.modalidad.capitalize()}."
            )
            return

        exito = self.engine.agregar_estudiante(grado, nombre, cedula, sexo)
        if exito:
            root = self.winfo_toplevel()
            if hasattr(root, "mostrar_toast"):
                root.mostrar_toast(f"✓ {nombre} agregado con éxito", color="#10B981")
            self.entry_nuevo_nombre.delete(0, "end")
            self.entry_nueva_cedula.delete(0, "end")
            self.cargar_lista(grado)
        else:
            messagebox.showerror(
                "Error",
                f"No se pudo agregar. La lista podría estar llena (Máx. {max_est} alumnos)."
            )

    def guardar_cambios(self):
        grado = self.combo_grado.get()
        datos_modificados = {}

        for id_est, entries in self.entradas_estudiantes.items():
            nom = entries["nombre"].get().strip().upper()
            ced = self.formatear_cedula_panamena(entries["cedula"].get().strip())
            sex = entries["sexo"].get()
            if nom:
                datos_modificados[id_est] = {"nombre": nom, "cedula": ced, "sexo": sex}

        if self.engine.guardar_cambios_estudiantes(grado, datos_modificados):
            root = self.winfo_toplevel()
            if hasattr(root, "mostrar_toast"):
                root.mostrar_toast("✓ Cambios guardados con éxito", color="#10B981")
            self.cargar_lista(grado)
        else:
            messagebox.showerror("Error", "Hubo un problema al guardar los cambios.")

    # ──────────────────────────────────────────────────────────────────────
    # Utilidades de cédula
    # ──────────────────────────────────────────────────────────────────────

    def formatear_cedula_panamena(self, val):
        val = val.strip().upper()
        val = "".join([c for c in val if c.isalnum() or c == "-"])
        if "-" in val:
            partes = val.split("-")
            if len(partes) == 3:
                p1 = partes[0]
                if p1 in ["PE", "PI", "E", "N"] or (p1.isdigit() and 1 <= int(p1) <= 13):
                    pass
                else:
                    return val
            else:
                return val
        clean = val.replace("-", "")
        if not clean:
            return ""
        prefix, rest = "", clean
        for p in ["PE", "PI", "E", "N"]:
            if clean.startswith(p):
                prefix, rest = p, clean[len(p):]
                break
        if not prefix:
            if len(clean) >= 2 and clean[:2].isdigit():
                prov = int(clean[:2])
                if 1 <= prov <= 13:
                    prefix, rest = clean[:2], clean[2:]
                elif clean[0].isdigit():
                    prefix, rest = clean[0], clean[1:]
            elif clean and clean[0].isdigit():
                prefix, rest = clean[0], clean[1:]
        if prefix:
            if len(rest) > 3:
                split_idx = max(1, len(rest) - 4)
                if len(rest) <= 6:
                    split_idx = min(3, len(rest) - 3)
                    if split_idx < 1:
                        split_idx = 1
                return f"{prefix}-{rest[:split_idx]}-{rest[split_idx:]}"
            elif len(rest) > 0:
                return f"{prefix}-{rest}"
            return prefix
        return val

    def _formatear_nueva_cedula(self):
        val = self.entry_nueva_cedula.get()
        fmt = self.formatear_cedula_panamena(val)
        self.entry_nueva_cedula.delete(0, "end")
        self.entry_nueva_cedula.insert(0, fmt)

    def _formatear_cedula_existente(self, ent):
        val = ent.get()
        fmt = self.formatear_cedula_panamena(val)
        ent.delete(0, "end")
        ent.insert(0, fmt)

    # ──────────────────────────────────────────────────────────────────────
    # Imprimir lista / Actualizar
    # ──────────────────────────────────────────────────────────────────────

    def imprimir_lista_clase(self):
        from rdsecurity import cargar_config_segura
        from documentos_maestro import generar_lista_clase
        from utils.footer_utils import abrir_documento

        grado = self.combo_grado.get()
        if grado == "Sin datos":
            messagebox.showwarning("Atención", "No hay un grado válido seleccionado.")
            return

        estudiantes = self.engine.obtener_estudiantes_completos(grado)
        if not estudiantes:
            messagebox.showwarning("Atención", "No hay estudiantes en este grado.")
            return

        cfg = cargar_config_segura({})
        datos = {
            "grado":         grado,
            "docente_nombre": cfg.get("docente_nombre", ""),
            "escuela_nombre": cfg.get("escuela_nombre", ""),
            "ano_lectivo":   cfg.get("ano_lectivo", "2026"),
            "estudiantes":   estudiantes,
        }

        ruta = generar_lista_clase(datos)
        if ruta:
            abrir_documento(ruta)
            messagebox.showinfo("✓ Lista Generada",
                                f"Lista de clase en Word generada exitosamente:\n{ruta}")

    def actualizar_vista(self):
        """Recarga la lista de grados activos y refresca la vista."""
        opciones = self.engine.obtener_grados_activos() or ["Sin datos"]
        old_sel  = self.combo_grado.get()
        self.combo_grado.configure(values=opciones)
        if old_sel in opciones:
            self.combo_grado.set(old_sel)
        else:
            self.combo_grado.set(opciones[0])
        grado = self.combo_grado.get()
        self.cargar_lista(grado)
        self.cargar_retirados(grado)
        self._initialized = True


# ══════════════════════════════════════════════════════════════════════════════
#  Diálogo de Retiro
# ══════════════════════════════════════════════════════════════════════════════

class _DialogoRetiro(ctk.CTkToplevel):
    """Diálogo modal que pide el motivo y el nombre del acudiente al retirar un alumno."""

    def __init__(self, parent, nombre_alumno: str):
        super().__init__(parent)
        self.title("🚪 Registrar Retiro de Alumno")
        self.geometry("480x300")
        self.resizable(False, False)
        self.grab_set()
        self.lift()
        self.focus_force()
        self.protocol("WM_DELETE_WINDOW", self._cancelar)

        self.confirmado = False
        self.motivo     = ""
        self.acudiente  = ""

        # Encabezado
        ctk.CTkLabel(self, text="🚪 Registrar Retiro",
                     font=(FONT_TITLE, 18, "bold")).pack(pady=(18, 4))
        ctk.CTkLabel(self,
                     text=f"Alumno: {nombre_alumno}",
                     font=(FONT_BODY, 13),
                     text_color="#FCA5A5").pack(pady=(0, 12))

        # Motivo
        ctk.CTkLabel(self, text="Motivo del retiro *",
                     font=(FONT_BODY, 12, "bold"),
                     anchor="w").pack(fill="x", padx=30)
        self.entry_motivo = ctk.CTkEntry(self, width=420,
                                         placeholder_text="Ej: Traslado a otra escuela")
        self.entry_motivo.pack(padx=30, pady=(2, 10))
        self.entry_motivo.focus_set()

        # Nombre del acudiente
        ctk.CTkLabel(self, text="Nombre del acudiente (opcional)",
                     font=(FONT_BODY, 12, "bold"),
                     anchor="w").pack(fill="x", padx=30)
        self.entry_acudiente = ctk.CTkEntry(self, width=420,
                                             placeholder_text="Ej: María González")
        self.entry_acudiente.pack(padx=30, pady=(2, 16))

        # Botones
        f_btns = ctk.CTkFrame(self, fg_color="transparent")
        f_btns.pack()
        ctk.CTkButton(f_btns, text="✅ Confirmar Retiro",
                      fg_color="#F97316", hover_color="#EA580C",
                      width=180, font=(FONT_BODY, 13, "bold"),
                      command=self._confirmar).pack(side="left", padx=10)
        ctk.CTkButton(f_btns, text="Cancelar",
                      fg_color="#374151", hover_color="#4B5563",
                      width=120,
                      command=self.destroy).pack(side="left", padx=10)

        self.bind("<Return>",  lambda e: self._confirmar())
        self.bind("<Escape>",  lambda e: self._cancelar())
        self.protocol("WM_DELETE_WINDOW", self._cancelar)

    def _confirmar(self):
        motivo = self.entry_motivo.get().strip()
        if not motivo:
            self.entry_motivo.configure(border_color="#EF4444")
            self.entry_motivo.focus_set()
            return
        self.motivo    = motivo
        self.acudiente = self.entry_acudiente.get().strip()
        self.confirmado = True
        self._cancelar()

    def _cancelar(self):
        try:
            self.grab_release()
        except Exception:
            pass
        self.destroy()