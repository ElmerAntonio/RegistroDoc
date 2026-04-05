"""
RegistroDoc Pro — Dashboard Principal
Diseño basado en la imagen de referencia:
  - Panel principal con métricas, gráfica de línea y barra
  - Marca de agua sutil de Panamá
  - Ventana fija 1280x720 (no redimensionable)
  - Paleta: fondo #0A1628, acentos cian #00DDEB / verde #00FF88
"""

import customtkinter as ctk
import tkinter as tk
import datetime

# Matplotlib para las gráficas
import matplotlib
matplotlib.use("TkAgg")
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.figure import Figure

# ─── PALETA PRINCIPAL CORREGIDA ──────────────────────────────────────────────
C = {
    "fondo":        "#0A1628",
    "sidebar":      "#0D1F35",
    "header":       "#0D1F35",
    "card":         "#0F2744",
    "card_borde":   "#00DDEB",
    "cian":         "#00DDEB",
    "verde":        "#00FF88",
    "rojo":         "#FF4444",
    "amarillo":     "#FFD700",
    "purpura":      "#A855F7",
    "texto":        "#E2E8F0",
    "texto_sec":    "#64748B",
    "texto_dim":    "#334155",
    "input":        "#1E3A5F",
    "hover":        "#1A3352",
    "activo":       "#1A3352",
    "borde":        "#1E3A5F",
    "glow":         "#1E3A5F",
}

# Color de acento actual (cambiable)
ACENTO = C["cian"]


# ─── CLASE PRINCIPAL DEL DASHBOARD ─────────────────────────────────────────
class DashboardFrame(ctk.CTkFrame):
    """
    Panel dashboard completo.
    Parámetros:
        master      : contenedor padre
        engine      : DataEngine con los métodos del motor de datos
        app_principal: referencia a RegistroDocApp para navegar
    """

    def __init__(self, master, engine, app_principal=None, **kwargs):
        super().__init__(master, fg_color=C["fondo"], **kwargs)
        self.engine        = engine
        self.app           = app_principal
        self._acento       = ACENTO

        # To support dynamic individual search
        self._current_student = None

        self._student_cache = None

        # Stats una sola vez
        self._stats = self._cargar_stats()

        self._construir()

    # ══════════════════════════════════════════════════════════════════════
    #  DATOS
    # ══════════════════════════════════════════════════════════════════════
    def _cargar_stats(self) -> dict:
        try:
            return self.engine.get_dashboard_stats()
        except Exception:
            return {
                "total": 92, "riesgo": 7, "honor": "SANTOS FIDEL",
                "honor_prom": "4.9", "asistencia": "98%",
                "grados": ["7°", "8°", "9°"],
            }

    # ══════════════════════════════════════════════════════════════════════
    #  CONSTRUCCIÓN PRINCIPAL
    # ══════════════════════════════════════════════════════════════════════
    def _construir(self):
        self.pack_propagate(False)
        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)

        # Panel principal
        self._panel_principal(self)


    # ══════════════════════════════════════════════════════════════════════
    #  PANEL PRINCIPAL
    # ══════════════════════════════════════════════════════════════════════
    def _panel_principal(self, parent):
        panel = ctk.CTkScrollableFrame(parent, fg_color=C["fondo"],
                                       corner_radius=0,
                                       scrollbar_fg_color=C["fondo"],
                                       scrollbar_button_color=C["borde"])
        panel.grid(row=0, column=0, sticky="nsew", padx=0, pady=0)
        panel.columnconfigure(0, weight=1)
        self._panel_container = panel

        # Marca de agua Panamá (canvas con texto muy tenue)
        self._marca_agua(panel)

        # Título y fecha
        self._titulo_seccion(panel)

        # Tarjetas métricas
        self._tarjetas_metricas(panel)

        # Gráficas
        self._graficas(panel)

        # Grados activos y accesos rápidos
        self._footer_panel(panel)

    # ── Marca de agua ──────────────────────────────────────────────────
    def _marca_agua(self, parent):
        """Canvas con texto PANAMÁ muy tenue como marca de agua."""
        cv = tk.Canvas(parent, bg=C["fondo"], height=0,
                       highlightthickness=0, bd=0)
        cv.pack(fill="x")
        # Se dibuja sobre toda la pantalla con place desde el panel principal
        # (técnica: se pone un label gigante transparente atrás)
        fondo_lbl = ctk.CTkLabel(
            parent,
            text="P  A  N  A  M  Á",
            font=ctk.CTkFont("Segoe UI", 72, "bold"),
            text_color="#0D2040",        # casi invisible sobre fondo
            fg_color="transparent")
        fondo_lbl.pack(pady=(0, 0))
        # No ocupa espacio visible — es un efecto decorativo

    # ── Título ─────────────────────────────────────────────────────────
    def _titulo_seccion(self, parent):
        f = ctk.CTkFrame(parent, fg_color="transparent")
        f.pack(fill="x", padx=24, pady=(10, 4))
        f.columnconfigure(1, weight=1)

        ctk.CTkLabel(f, text="Panel de Control Principal",
                     font=ctk.CTkFont("Segoe UI", 22, "bold"),
                     text_color=C["texto"]).grid(row=0, column=0, sticky="w")

        # Barra de búsqueda integrada al Dashboard
        search_f = ctk.CTkFrame(f, fg_color=C["input"], corner_radius=20, border_width=1, border_color=C["borde"], height=34)
        search_f.grid(row=0, column=1, sticky="ew", padx=(40, 40))
        search_f.grid_propagate(False)
        search_f.columnconfigure(1, weight=1)

        ctk.CTkLabel(search_f, text="🔍", font=ctk.CTkFont(size=14), text_color=C["texto_sec"], width=30).grid(row=0, column=0, padx=(10,0))

        self._search_var = tk.StringVar()
        self._search_entry = ctk.CTkEntry(
            search_f, textvariable=self._search_var,
            fg_color="transparent", border_width=0,
            placeholder_text="Buscar alumno para ver su rendimiento...",
            font=ctk.CTkFont("Segoe UI", 13),
            text_color=C["texto"])
        self._search_entry.grid(row=0, column=1, sticky="ew", padx=(4,10), pady=4)

        # We bind the Return key so we can perform the search
        self._search_entry.bind("<Return>", self._ejecutar_busqueda)

        fecha = datetime.datetime.now().strftime("%A %d de %B, %Y").capitalize()
        ctk.CTkLabel(f, text=fecha,
                     font=ctk.CTkFont("Segoe UI", 11),
                     text_color=C["texto_sec"]).grid(row=0, column=2, sticky="e")

        # Toast message frame for not found errors
        self._toast_lbl = ctk.CTkLabel(f, text="", text_color=C["rojo"], font=ctk.CTkFont("Segoe UI", 12, "bold"))
        self._toast_lbl.grid(row=1, column=1, pady=(2, 0))

    def _ejecutar_busqueda(self, event=None):
        texto = self._search_var.get().strip().lower()
        if not texto:
            self._current_student = None
            self._actualizar_dashboard_contextual()
            return

        encontrado = None
        try:
            if self._student_cache is None:
                import time
                start_cache = time.perf_counter()
                self._student_cache = []
                for g in self.engine.obtener_grados_activos():
                    for est in self.engine.obtener_estudiantes_completos(g):
                        self._student_cache.append({"grado": g, "nombre": est["nombre"]})
                print(f"[*] Built student cache in {time.perf_counter() - start_cache:.4f}s")

            start_search = time.perf_counter()
            for cached_est in self._student_cache:
                if texto in cached_est["nombre"].lower():
                    encontrado = cached_est
                    break
            print(f"[*] Search completed in {time.perf_counter() - start_search:.6f}s")

        except Exception:
            # Fallback mock for UI test
            if "maria" in texto:
                encontrado = {"grado": "7°A", "nombre": "Maria Gonzalez"}

        if encontrado:
            self._current_student = encontrado
            self._toast_lbl.configure(text="")
            self._actualizar_dashboard_contextual()
        else:
            self._toast_lbl.configure(text="❌ Estudiante no registrado")
            self._search_var.set("")
            self._current_student = None
            self._actualizar_dashboard_contextual()
            # Ocultar toast despues de 3 segundos
            def clear_toast():
                if self.winfo_exists() and hasattr(self, "_toast_lbl") and self._toast_lbl.winfo_exists():
                    self._toast_lbl.configure(text="")

            self.after(3000, clear_toast)

    def _actualizar_dashboard_contextual(self):
        # Refresh the cards and graph based on _current_student state
        if hasattr(self, "_cards_frame") and self._cards_frame.winfo_exists():
            self._cards_frame.destroy()
            self._tarjetas_metricas(self._panel_container)

        if hasattr(self, "_graph_frame") and self._graph_frame.winfo_exists():
            self._graph_frame.destroy()
            self._graficas(self._panel_container)

    # ── Tarjetas métricas ────────────────────────────────────────────────
    def _tarjetas_metricas(self, parent):
        self._cards_frame = ctk.CTkFrame(parent, fg_color="transparent")
        # Ensure it always stays above footer, below title. Since pack is used, we have to insert it at correct index if recreating
        # We will pack it directly, but since we are recreating it during dynamic search,
        # a better approach would be to have dedicated container frames.
        # But we'll just reconstruct the entire view or place it carefully.
        # Actually it's easier to just destroy and pack everything below the title again.

        # Let's clean the container elements below the title before rendering
        for widget in parent.winfo_children():
            if widget not in (self._cards_frame, parent.winfo_children()[0], parent.winfo_children()[1]):
                if hasattr(self, "_cards_frame") and widget == self._cards_frame: continue
                # Do not destroy water mark or title section

        self._cards_frame.pack(fill="x", padx=20, pady=(6, 10))
        self._cards_frame.columnconfigure((0, 1, 2, 3), weight=1, uniform="card")

        if self._current_student:
            # Contextual mode for individual student
            nombre_est = self._current_student["nombre"]

            datos = [
                ("👤",  "Estudiante",           "1",    nombre_est,                 self._acento),
                ("⚠️",  "Estatus de Riesgo",    "Estable", "Promedio normal",       C["verde"]),
                ("🏆",  "Cuadro de Honor",      "—",    "No aplica",                C["texto_sec"]),
                ("📊",  "Asistencia Indiv.",    "100%", "Sin ausencias",            self._acento),
            ]
        else:
            # General mode
            s = self._stats
            total    = str(s.get("total", 0))
            riesgo   = str(s.get("riesgo", 0))
            honor    = str(s.get("honor", "—"))
            asist    = str(s.get("asistencia", "0%"))

            datos = [
                ("👥",  "Total Alumnos",       total,  "Matriculados",             self._acento),
                ("⚠️",  "Alumnos en Riesgo",   riesgo, "Nota promedio < 3.0",      "#EF4444"),
                ("🏆",  "Cuadro de Honor",     "1",    honor[:18],                 C["amarillo"]),
                ("📊",  "Asistencia",          asist,  "Promedio general",         C["verde"]),
            ]

        for col, (ico, titulo, valor, sub, color) in enumerate(datos):
            self._tarjeta(self._cards_frame, col, ico, titulo, valor, sub, color)

    def _tarjeta(self, parent, col, ico, titulo, valor, sub, color):
        card = ctk.CTkFrame(parent, fg_color=C["card"],
                            border_width=1, border_color=color,
                            corner_radius=12)
        card.grid(row=0, column=col, padx=6, pady=4, sticky="nsew")

        # Icono
        ctk.CTkLabel(card, text=ico, font=ctk.CTkFont(size=20),
                     fg_color="transparent",
                     text_color=color).pack(anchor="w", padx=16, pady=(14, 0))

        # Título (con wraplength para que no desborde)
        ctk.CTkLabel(card, text=titulo,
                     font=ctk.CTkFont("Segoe UI", 12),
                     text_color=C["texto_sec"],
                     wraplength=140,
                     justify="left").pack(anchor="w", padx=16, pady=(2, 0))

        # Valor grande
        ctk.CTkLabel(card, text=valor,
                     font=ctk.CTkFont("Segoe UI", 32, "bold"),
                     text_color=color,
                     wraplength=160,
                     justify="left").pack(anchor="w", padx=16, pady=(0, 2))

        # Subtítulo
        ctk.CTkLabel(card, text=sub,
                     font=ctk.CTkFont("Segoe UI", 11),
                     text_color=C["texto_sec"],
                     wraplength=150,
                     justify="left").pack(anchor="w", padx=16, pady=(0, 14))

    # ── Gráficas ─────────────────────────────────────────────────────────
    def _graficas(self, parent):
        self._graph_frame = ctk.CTkFrame(parent, fg_color="transparent")
        self._graph_frame.pack(fill="x", padx=20, pady=6)
        self._graph_frame.columnconfigure(0, weight=6)
        self._graph_frame.columnconfigure(1, weight=4)

        self._grafica_linea(self._graph_frame)
        self._grafica_barras(self._graph_frame)

    def _grafica_linea(self, parent):
        """Gráfica de escalada — evolución de notas T1→T2→T3."""
        card = ctk.CTkFrame(parent, fg_color=C["card"],
                            border_width=1, border_color=C["borde"],
                            corner_radius=12)
        card.grid(row=0, column=0, padx=(0, 8), sticky="nsew")

        ctk.CTkLabel(card, text="Tendencia de Rendimiento Académico  (X/Y)",
                     font=ctk.CTkFont("Segoe UI", 13, "bold"),
                     text_color=C["texto"]).pack(anchor="w", padx=16, pady=(12, 0))
        ctk.CTkLabel(card, text="Escala 1–5 por trimestre",
                     font=ctk.CTkFont("Segoe UI", 10),
                     text_color=C["texto_sec"]).pack(anchor="w", padx=16)

        # Matplotlib
        fig = Figure(figsize=(5.4, 2.8), dpi=96,
                     facecolor=C["card"])
        ax = fig.add_subplot(111)
        ax.set_facecolor(C["card"])
        fig.subplots_adjust(left=0.1, right=0.97, top=0.92, bottom=0.16)

        x = [0, 1, 2, 3, 4, 5, 6, 7]
        y = [2.0, 2.5, 3.1, 3.0, 3.4, 3.7, 4.1, 4.8]

        ax.plot(x, y, color=self._acento, linewidth=2.2, label="Promedio Grupal",
                zorder=3, solid_capstyle="round")
        ax.fill_between(x, y, alpha=0.12, color=self._acento)
        ax.scatter(x, y, color=self._acento, s=55, zorder=5,
                   edgecolors=C["fondo"], linewidths=1.5)

        if getattr(self, "_current_student", None):
            # Highlight individual student point (fake data for visualization)
            y_indiv = [2.2, 2.7, 3.5, 3.3, 3.8, 4.0, 4.5, 5.0]
            ax.plot(x, y_indiv, color=C["amarillo"], linewidth=2.2, label=self._current_student["nombre"][:12],
                    zorder=4, linestyle="--")
            ax.scatter(x, y_indiv, color=C["amarillo"], s=60, zorder=6,
                       edgecolors=C["fondo"], linewidths=1.5)
            ax.legend(loc="upper left", facecolor=C["fondo"], edgecolor=C["borde"], labelcolor=C["texto"], fontsize=8)

        # Ejes
        ax.set_xlim(-0.3, 7.3)
        ax.set_ylim(1, 5.2)
        ax.set_xlabel("X  (Trimestres)", color=C["texto_sec"], fontsize=9)
        ax.set_ylabel("Y  (Notas 1–5)", color=C["texto_sec"], fontsize=9)
        ax.tick_params(colors=C["texto_sec"], labelsize=8)
        for spine in ax.spines.values():
            spine.set_edgecolor(C["borde"])
        ax.grid(True, color=C["borde"], linewidth=0.5, alpha=0.6)
        ax.set_xticks(x)

        canvas = FigureCanvasTkAgg(fig, master=card)
        canvas.draw()
        canvas.get_tk_widget().pack(fill="x", padx=10, pady=(4, 12))

    def _grafica_barras(self, parent):
        """Barras de asistencia / rendimiento por grupo."""
        card = ctk.CTkFrame(parent, fg_color=C["card"],
                            border_width=1, border_color=C["borde"],
                            corner_radius=12)
        card.grid(row=0, column=1, sticky="nsew")

        ctk.CTkLabel(card, text="Rendimiento por Grupo",
                     font=ctk.CTkFont("Segoe UI", 13, "bold"),
                     text_color=C["texto"]).pack(anchor="w", padx=16, pady=(12, 0))
        ctk.CTkLabel(card, text="Promedio final por sección",
                     font=ctk.CTkFont("Segoe UI", 10),
                     text_color=C["texto_sec"]).pack(anchor="w", padx=16)

        fig2 = Figure(figsize=(3.2, 2.8), dpi=96,
                      facecolor=C["card"])
        ax2 = fig2.add_subplot(111)
        ax2.set_facecolor(C["card"])
        fig2.subplots_adjust(left=0.14, right=0.97, top=0.9, bottom=0.18)

        grupos   = ["7°", "8°", "9°", "7°B", "8°B"]
        valores  = [3.8, 3.5, 4.2, 3.1, 3.7]
        colores  = [self._acento] * len(grupos)

        ax2.bar(grupos, valores, color=colores, width=0.6,
                edgecolor=C["fondo"], linewidth=0.8)

        ax2.set_ylim(1, 5.3)
        ax2.tick_params(colors=C["texto_sec"], labelsize=8)
        for spine in ax2.spines.values():
            spine.set_edgecolor(C["borde"])
        ax2.grid(True, axis="y", color=C["borde"],
                 linewidth=0.5, alpha=0.5)
        ax2.set_ylabel("Promedio", color=C["texto_sec"], fontsize=8)

        canvas2 = FigureCanvasTkAgg(fig2, master=card)
        canvas2.draw()
        canvas2.get_tk_widget().pack(fill="x", padx=10, pady=(4, 12))

    # ── Footer: grados activos + accesos rápidos ─────────────────────────
    def _footer_panel(self, parent):
        self._footer_frame = ctk.CTkFrame(parent, fg_color="transparent")
        ff = self._footer_frame
        ff.pack(fill="x", padx=20, pady=(4, 20))
        ff.columnconfigure((0, 1, 2), weight=1, uniform="bot")

        # Grados activos
        gc = ctk.CTkFrame(ff, fg_color=C["card"],
                          border_width=1, border_color=C["borde"],
                          corner_radius=12)
        gc.grid(row=0, column=0, padx=(0, 8), sticky="nsew", pady=4)

        ctk.CTkLabel(gc, text="Grados Activos",
                     font=ctk.CTkFont("Segoe UI", 14, "bold"),
                     text_color=C["texto"]).pack(anchor="w", padx=16, pady=(14, 8))

        try:
            grados = self.engine.obtener_grados_activos()
        except Exception:
            grados = ["7°", "8°", "9°"]

        gf = ctk.CTkFrame(gc, fg_color="transparent")
        gf.pack(padx=16, pady=(0, 14))
        for g in grados:
            pill = ctk.CTkLabel(
                gf, text=f"  {g}  ",
                fg_color=C["input"],
                corner_radius=16,
                font=ctk.CTkFont("Segoe UI", 14, "bold"),
                text_color=self._acento,
                padx=12, pady=6)
            pill.pack(side="left", padx=4)

        # Horario sincronizado
        hc = ctk.CTkFrame(ff, fg_color=C["card"],
                          border_width=1, border_color=C["borde"],
                          corner_radius=12)
        hc.grid(row=0, column=1, padx=(0, 8), sticky="nsew", pady=4)

        ctk.CTkLabel(hc, text="Horario del Día",
                     font=ctk.CTkFont("Segoe UI", 14, "bold"),
                     text_color=C["texto"]).pack(anchor="w", padx=16, pady=(14, 8))

        hf = ctk.CTkFrame(hc, fg_color="transparent")
        hf.pack(fill="x", padx=16, pady=(0, 14))

        # Obtener horario actual
        materia_actual, hora_actual = self._obtener_materia_actual()
        
        if materia_actual:
            ctk.CTkLabel(hf, text=f"📚 {materia_actual}",
                        font=ctk.CTkFont("Segoe UI", 16, "bold"),
                        text_color=self._acento).pack(anchor="w")
            ctk.CTkLabel(hf, text=f"🕐 {hora_actual}",
                        font=ctk.CTkFont("Segoe UI", 12),
                        text_color=C["texto_sec"]).pack(anchor="w", pady=(2, 0))
        else:
            ctk.CTkLabel(hf, text="📚 Fuera de horario escolar",
                        font=ctk.CTkFont("Segoe UI", 14),
                        text_color=C["texto_sec"]).pack(anchor="w")
            ctk.CTkLabel(hf, text="🕐 No hay clases programadas",
                        font=ctk.CTkFont("Segoe UI", 12),
                        text_color=C["texto_sec"]).pack(anchor="w", pady=(2, 0))

        # Accesos rápidos
        ac = ctk.CTkFrame(ff, fg_color=C["card"],
                          border_width=1, border_color=C["borde"],
                          corner_radius=12)
        ac.grid(row=0, column=2, sticky="nsew", pady=4)

        ctk.CTkLabel(ac, text="Accesos Rápidos",
                     font=ctk.CTkFont("Segoe UI", 14, "bold"),
                     text_color=C["texto"]).pack(anchor="w", padx=16, pady=(14, 8))

        af = ctk.CTkFrame(ac, fg_color="transparent")
        af.pack(fill="x", padx=16, pady=(0, 14))
        af.columnconfigure((0, 1, 2), weight=1)

        def nav_notas():
            if self.app:
                try: self.app.mostrar_notas()
                except Exception: pass

        def nav_asistencia():
            if self.app:
                try: self.app.mostrar_asistencia()
                except Exception: pass

        def nav_config():
            if self.app:
                try: self.app.mostrar_configuracion()
                except Exception: pass

        botones = [
            ("📝 Notas",      "#2563EB", nav_notas),
            ("📅 Asistencia", "#059669", nav_asistencia),
            ("⚙️ Config",     "#7C3AED", nav_config),
        ]
        for i, (txt, col, cmd) in enumerate(botones):
            ctk.CTkButton(af, text=txt, height=48,
                          font=ctk.CTkFont("Segoe UI", 13, "bold"),
                          fg_color=col, hover_color=col,
                          corner_radius=10,
                          command=cmd).grid(
                row=0, column=i, padx=5, sticky="ew")


    def _obtener_materia_actual(self):
        """Obtiene la materia actual basada en el día y hora actuales."""
        try:
            # Obtener día de la semana (0=lunes, 6=domingo)
            dia_semana = datetime.datetime.now().weekday()
            
            # Mapear a nombres del horario (solo días de semana)
            dias_horario = ["lunes", "martes", "miercoles", "jueves", "viernes"]
            
            if dia_semana >= 5:  # Sábado o domingo
                return None, None
            
            dia_actual = dias_horario[dia_semana]
            
            # Obtener hora actual
            hora_actual = datetime.datetime.now().time()
            
            # Obtener horario del engine
            horario = self.engine.obtener_horario()
            
            # Buscar el período actual
            for periodo in horario:
                if not periodo["horas"] or not periodo[dia_actual]:
                    continue
                
                # Parsear el rango de horas (ej: "7:00-7:45")
                try:
                    horas_str = periodo["horas"].strip()
                    if "-" in horas_str:
                        inicio_str, fin_str = horas_str.split("-")
                        inicio = datetime.datetime.strptime(inicio_str.strip(), "%H:%M").time()
                        fin = datetime.datetime.strptime(fin_str.strip(), "%H:%M").time()
                        
                        if inicio <= hora_actual <= fin:
                            return periodo[dia_actual], periodo["horas"]
                except (ValueError, AttributeError):
                    continue
            
            return None, None
            
        except Exception:
            return None, None


# ─── MOCK PARA PRUEBA STANDALONE ─────────────────────────────────────────────
if __name__ == "__main__":
    ctk.set_appearance_mode("dark")
    ctk.set_default_color_theme("blue")

    class MockEngine:
        def get_dashboard_stats(self):
            return {"total": 92, "riesgo": 7,
                    "honor": "SANTOS FIDEL", "honor_prom": "4.9",
                    "asistencia": "98%"}
        def obtener_grados_activos(self):
            return ["7°", "8°", "9°"]
        def obtener_estudiantes_completos(self, g):
            names = ["Maria Gonzalez", "Pedro Pinto", "Ana Santos"]
            return [{"id": i+1, "nombre": n, "cedula": ""} for i, n in enumerate(names)]
        def obtener_datos_generales(self):
            return {"docente_nombre": "Prof. Elmer Tugri",
                    "correo": "elmer.tugri7@meduca.edu.pa"}
        def obtener_horario(self):
            return [
                {"horas": "7:00-7:45", "lunes": "Matemáticas", "martes": "Español", "miercoles": "Ciencias", "jueves": "Historia", "viernes": "Inglés"},
                {"horas": "7:45-8:30", "lunes": "Español", "martes": "Matemáticas", "miercoles": "Inglés", "jueves": "Ciencias", "viernes": "Historia"},
                {"horas": "8:30-9:15", "lunes": "Ciencias", "martes": "Historia", "miercoles": "Matemáticas", "jueves": "Español", "viernes": "Ciencias"},
                {"horas": "9:15-10:00", "lunes": "Historia", "martes": "Inglés", "miercoles": "Español", "jueves": "Matemáticas", "viernes": "Inglés"},
                {"horas": "10:15-11:00", "lunes": "Inglés", "martes": "Ciencias", "miercoles": "Historia", "jueves": "Inglés", "viernes": "Matemáticas"},
                {"horas": "11:00-11:45", "lunes": "Educación Física", "martes": "Arte", "miercoles": "Música", "jueves": "Computación", "viernes": "Educación Cívica"},
                {"horas": "11:45-12:30", "lunes": "Computación", "martes": "Educación Física", "miercoles": "Arte", "jueves": "Música", "viernes": "Computación"},
                {"horas": "12:30-13:15", "lunes": "Religión", "martes": "Computación", "miercoles": "Educación Física", "jueves": "Arte", "viernes": "Música"}
            ]

    root = ctk.CTk()
    root.title("RegistroDoc Pro v3.0")
    # ── VENTANA FIJA 1280×720 ──────────────────────────────────────────
    root.geometry("1280x720")
    root.minsize(1280, 720)
    root.maxsize(1280, 720)
    root.resizable(False, False)

    DashboardFrame(root, MockEngine()).pack(fill="both", expand=True)
    root.mainloop()
