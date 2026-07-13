import customtkinter as ctk
import tkinter as tk
import matplotlib.pyplot as plt
from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import threading
from theme import FONT_BODY

class ToolTip:
    def __init__(self, widget, text):
        self.widget = widget
        self.text = text
        self.tooltip_window = None
        self.widget.bind("<Enter>", self.show_tooltip)
        self.widget.bind("<Leave>", self.hide_tooltip)

    def show_tooltip(self, event=None):
        if self.tooltip_window or not self.text:
            return
        x = self.widget.winfo_rootx() + 20
        y = self.widget.winfo_rooty() + 25
        self.tooltip_window = tw = tk.Toplevel(self.widget)
        tw.wm_overrideredirect(True)
        tw.wm_geometry(f"+{x}+{y}")
        label = tk.Label(tw, text=self.text, justify='left',
                         background="#1e293b", foreground="#f8fafc",
                         relief='solid', borderwidth=1,
                         font=("Segoe UI", 9), padx=8, pady=6)
        label.pack()

    def hide_tooltip(self, event=None):
        tw = self.tooltip_window
        self.tooltip_window = None
        if tw:
            tw.destroy()

class GraficosFrame(ctk.CTkFrame):
    def __init__(self, master, engine, **kwargs):
        super().__init__(master, fg_color="transparent", **kwargs)
        self.engine = engine

        from theme import obtener_C_res
        self.C = obtener_C_res()

        # Estructura principal
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)

        # Contenedor de gráficos con scroll
        self.scroll_canvas = ctk.CTkScrollableFrame(self, fg_color="transparent")
        self.scroll_canvas.grid(row=1, column=0, sticky="nsew", padx=10, pady=10)
        self.fig_objs = []
        self.scroll_canvas.grid_columnconfigure((0, 1), weight=1)

        self.fig_objs = [] # Para limpiar figuras después

        self.crear_filtros()
        # Los gráficos se cargan en actualizar_vista() al mostrar el frame
        self._initialized = False

    def crear_filtros(self):
        f_top = ctk.CTkFrame(self, fg_color=self.C["card"], corner_radius=10)
        f_top.grid(row=0, column=0, sticky="ew", padx=15, pady=15)
        
        ctk.CTkLabel(f_top, text="📊 Análisis Visual de Rendimiento",
                    font=("Segoe UI", 20, "bold"), text_color=self.C["cian"]).pack(side="left", padx=20, pady=15)

        # Grado
        grados = self.engine.obtener_grados_activos()
        if not grados: 
            grados = ["Sin datos"]
        else:
            grados = ["Todos los Grados"] + grados

        self.combo_grado = ctk.CTkOptionMenu(f_top, values=grados, command=self.al_cambiar_grado)
        self.combo_grado.pack(side="left", padx=5)
        self.combo_grado.set(grados[0])

        # Materia
        materias = self.engine.obtener_materias_por_grado(grados[0]) if grados[0] != "Sin datos" else ["Sin materias"]
        if materias and materias != ["Sin materias"]:
            materias = ["Todas las Materias"] + materias
        self.combo_materia = ctk.CTkOptionMenu(f_top, values=materias, command=lambda _: self.actualizar_graficos())
        self.combo_materia.pack(side="left", padx=5)
        if materias: self.combo_materia.set(materias[0])

        # Estudiante
        self.combo_estudiante = ctk.CTkOptionMenu(f_top, values=["Todos los Estudiantes"], command=lambda _: self.actualizar_graficos())
        self.combo_estudiante.pack(side="left", padx=5)

        # Trimestre (si aplica)
        self.combo_trimestre = ctk.CTkOptionMenu(f_top, values=["Trimestre 1", "Trimestre 2", "Trimestre 3", "Todos los Trimestres"], command=lambda _: self.actualizar_graficos())
        self.combo_trimestre.pack(side="left", padx=5)

        # Opciones de graficos (agrupadas)
        f_opciones = ctk.CTkFrame(f_top, fg_color="#13263f", corner_radius=8)
        f_opciones.pack(side="left", padx=12, pady=8)

        ctk.CTkLabel(
            f_opciones,
            text="Mostrar graficos",
            font=("Segoe UI", 12, "bold"),
            text_color=self.C["texto"],
        ).pack(anchor="w", padx=10, pady=(6, 2))

        checks_grid = ctk.CTkFrame(f_opciones, fg_color="transparent")
        checks_grid.pack(padx=8, pady=(0, 8))

        check_specs = [
            ("chk_pastel", "Pastel", True),
            ("chk_barras", "Barras", True),
            ("chk_tendencia", "Tendencia", True),
            ("chk_hist", "Histograma", False),
            ("chk_pareto", "Pareto", False),
            ("chk_scatter", "Dispersion", False),
            ("chk_box", "Box-plot", False),
            ("chk_lineas", "Lineas Trim", True),
            ("chk_heat", "Heatmap", False),
            ("chk_habitos", "Hábitos", True),
        ]

        descripciones_tips = {
            "chk_pastel": "Porcentajes de Aprobados, En Riesgo y Reprobados del grupo.",
            "chk_barras": "Comparación visual del promedio final de cada estudiante.",
            "chk_tendencia": "Progresión promedio grupal a lo largo de las actividades.",
            "chk_hist": "Distribución y cantidad de notas en rangos de 1.0 a 5.0.",
            "chk_pareto": "Identifica de forma ordenada el rendimiento acumulativo.",
            "chk_scatter": "Relación entre Calificaciones y el Porcentaje de Asistencia.",
            "chk_box": "Analiza la dispersión, mediana y valores atípicos de notas.",
            "chk_lineas": "Evolución comparativa del promedio entre T1, T2 y T3.",
            "chk_heat": "Visualización matricial coloreada del rendimiento por materias.",
            "chk_habitos": "Frecuencia evaluada en Hábitos y Aptitudes (S, R, X)."
        }

        grid_idx = 0
        for attr_name, label, should_grid in check_specs:
            chk = ctk.CTkCheckBox(
                checks_grid,
                text=label,
                command=self.actualizar_graficos,
                width=120,
            )
            if should_grid:
                chk.grid(row=grid_idx // 3, column=grid_idx % 3, sticky="w", padx=8, pady=2)
                chk.select()
                grid_idx += 1
            else:
                chk.deselect()
            setattr(self, attr_name, chk)
            tip_text = descripciones_tips.get(attr_name, "")
            if tip_text:
                ToolTip(chk, tip_text)

        ctk.CTkButton(f_top, text="ℹ️ Guía de Uso", fg_color="#00DDEB", text_color="#0A1628", hover_color="#00C0CD", font=("Segoe UI", 12, "bold"), command=self.mostrar_guia_graficas).pack(side="right", padx=5)
        ctk.CTkButton(f_top, text="Actualizar", fg_color="#3B82F6", command=self.actualizar_graficos).pack(side="right", padx=5)

    def actualizar_vista(self):
        """Recarga grados activos y actualiza los gráficos."""
        grados = self.engine.obtener_grados_activos()
        if not grados:
            grados = ["Sin datos"]
        else:
            grados = ["Todos los Grados"] + grados
        old_sel = self.combo_grado.get()
        self.combo_grado.configure(values=grados)
        if old_sel in grados:
            self.combo_grado.set(old_sel)
        else:
            self.combo_grado.set(grados[0])
        self.al_cambiar_grado(self.combo_grado.get())
        self._initialized = True

    def mostrar_guia_graficas(self):
        from tkinter import messagebox
        guia_texto = (
            "📊 Guía de Gráficos de Rendimiento:\n\n"
            "• Pastel: Porcentajes de aprobación, riesgo y reprobación.\n"
            "• Barras: Comparación del promedio final de cada estudiante.\n"
            "• Tendencia: Línea de progresión grupal del salón.\n"
            "• Histograma: Distribución y frecuencia de las calificaciones.\n"
            "• Pareto: Identifica la concentración de alumnos en rangos clave.\n"
            "• Dispersión: Relación entre notas obtenidas y la asistencia.\n"
            "• Box-plot: Visualiza la distribución de notas, mediana y dispersión.\n"
            "• Líneas Trim: Muestra el progreso académico del T1 al T3.\n"
            "• Heatmap: Mapa de calor de rendimiento de notas por materia.\n"
            "• Hábitos: Frecuencia de Hábitos y Aptitudes (S, R, X)."
        )
        messagebox.showinfo("Guía de Uso — Gráficos", guia_texto)

    def al_cambiar_grado(self, grado):
        if grado == "Sin datos": return
        if grado == "Todos los Grados":
            grados_activos = self.engine.obtener_grados_activos()
            todas_materias = []
            for g in grados_activos:
                todas_materias.extend(self.engine.obtener_materias_por_grado(g))
            materias = sorted(list(set(todas_materias)))
            if "Sin materias registradas" in materias:
                materias.remove("Sin materias registradas")
        else:
            materias = self.engine.obtener_materias_por_grado(grado)

        if materias:
            materias = ["Todas las Materias"] + materias
            self.combo_materia.configure(values=materias)
            self.combo_materia.set(materias[0])
        else:
            self.combo_materia.configure(values=["No hay materias"])
            self.combo_materia.set("No hay materias")

        if grado == "Todos los Grados":
            self.combo_estudiante.configure(values=["Todos los Estudiantes"])
            self.combo_estudiante.set("Todos los Estudiantes")
        else:
            estudiantes = self.engine.obtener_estudiantes_completos(grado)
            nombres = ["Todos los Estudiantes"] + [e['nombre'] for e in estudiantes]
            self.combo_estudiante.configure(values=nombres)
            self.combo_estudiante.set("Todos los Estudiantes")

        self.actualizar_graficos()


    def exportar_grafico_individual(self, fig, titulo):
        import tempfile
        import os
        from tkinter import messagebox
        try:
            from docx import Document
            from utils.docx_graphics import save_figure_as_image, add_image_to_doc
        except ImportError:
            messagebox.showerror("Error", "Falta librería python-docx o utils de gráficos.")
            return

        doc = Document()
        try:
            from utils.footer_utils import add_header_with_logo, get_school_logo_path
            logo_esc = get_school_logo_path()
            if logo_esc:
                add_header_with_logo(doc, logo_esc)
        except Exception: pass

        doc.add_heading(f"Gráfico: {titulo}", 0)
        try:
            img = save_figure_as_image(fig, "indiv_")
            add_image_to_doc(doc, img, ancho=6)

            with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp:
                doc.save(tmp.name)
                tmp_docx = tmp.name
            messagebox.showinfo("Exportar Gráfico", f"Gráfico exportado a:\n{tmp_docx}")
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo exportar el gráfico: {e}")

    def limpiar_graficos(self):
        # Cerrar figuras matplotlib para liberar memoria (evita leak de ~2-5MB por figura)
        for fig in self.fig_objs:
            try:
                plt.close(fig)
            except Exception:
                pass
        self.fig_objs.clear()
        for w in self.scroll_canvas.winfo_children():
            w.destroy()

    def actualizar_graficos(self):
        self.limpiar_graficos()

        # Mostrar indicador de carga
        self.lbl_loading = ctk.CTkLabel(
            self.scroll_canvas,
            text="🔄 Cargando datos y generando gráficos...",
            font=(FONT_BODY, 14, "bold"),
            text_color=self.C["cian"]
        )
        self.lbl_loading.pack(pady=60)

        # Leer filtros en el hilo principal
        grado = self.combo_grado.get()
        materia = self.combo_materia.get()
        trimestre = self.combo_trimestre.get()
        estudiante_sel = self.combo_estudiante.get()
        
        chk_states = {
            "pastel": self.chk_pastel.get(),
            "barras": self.chk_barras.get(),
            "tendencia": self.chk_tendencia.get(),
            "hist": self.chk_hist.get(),
            "pareto": self.chk_pareto.get(),
            "scatter": self.chk_scatter.get(),
            "box": self.chk_box.get(),
            "lineas": self.chk_lineas.get(),
            "heat": self.chk_heat.get(),
            "habitos": self.chk_habitos.get()
        }

        def bg_work():
            estudiantes = []
            try:
                if grado == "Todos los Grados":
                    for g in self.engine.obtener_grados_activos():
                        estudiantes.extend(self.engine.obtener_estudiantes_completos(g))
                else:
                    estudiantes = self.engine.obtener_estudiantes_completos(grado)
            except Exception: pass

            aprobados = 0; reprobados = 0; en_riesgo = 0

            # Obtener promedios
            promedios_por_est = {}
            obtener_promedios = getattr(self.engine, 'obtener_promedios_reales', None)
            if obtener_promedios:
                if grado == "Todos los Grados":
                    for g in self.engine.obtener_grados_activos():
                        try:
                            proms = obtener_promedios(g, materia, trimestre)
                            for k, v in proms.items():
                                promedios_por_est[f"{k} ({g})"] = v
                        except Exception: pass
                else:
                    try:
                        promedios_por_est = obtener_promedios(grado, materia, trimestre)
                    except Exception: pass

            if not promedios_por_est and estudiantes:
                # Fallback a 1.0 si no hay calificaciones válidas registradas aún
                for est in estudiantes:
                    promedios_por_est[est['nombre']] = 1.0

            for nom, promedio in promedios_por_est.items():
                if promedio >= 3.0: aprobados += 1
                elif promedio >= 2.5: en_riesgo += 1
                else: reprobados += 1

            # Para la gráfica de tendencia (promedio de cada actividad cronológica)
            tendencia_actividades = []
            if grado != "Todos los Grados" and materia not in ["Todas las Materias", "No hay materias", "Sin materias", "No hay materias registradas"]:
                trimestres_a_buscar = ["Trimestre 1", "Trimestre 2", "Trimestre 3"] if trimestre == "Todos los Trimestres" else [trimestre]
                tipos_a_buscar = ["Diaria / Parcial", "Apreciación", "Examen"]
                for t_name in trimestres_a_buscar:
                    for tipo_n in tipos_a_buscar:
                        try:
                            descs = self.engine.obtener_descripciones_notas(grado, materia, t_name, tipo_n)
                            for d in descs:
                                if not d or "Sin notas" in d or "Seleccione" in d:
                                    continue
                                res = self.engine.buscar_notas_por_descripcion_exacta(grado, materia, t_name, tipo_n, d)
                                if res and res["notas"]:
                                    notas_validas = [float(str(v).replace(',','.')) for v in res["notas"].values() if v is not None and str(v).strip() != ""]
                                    if notas_validas:
                                        prom = sum(notas_validas) / len(notas_validas)
                                        desc_corta = d if len(d) <= 12 else d[:10] + ".."
                                        prefix = "D" if "Diaria" in tipo_n else ("A" if "Apreciación" in tipo_n else "E")
                                        tendencia_actividades.append((f"{prefix}: {desc_corta}", prom))
                        except Exception: pass

            historial = []
            if estudiante_sel != "Todos los Estudiantes":
                historial = getattr(self.engine, 'obtener_historial_real', lambda g,m,e: [3.0, 3.0])(grado, materia, estudiante_sel)
                if len(historial) < 2: historial = [3.0, 3.0]

            datos = {
                "grado": grado,
                "materia": materia,
                "trimestre": trimestre,
                "estudiante_sel": estudiante_sel,
                "aprobados": aprobados,
                "reprobados": reprobados,
                "en_riesgo": en_riesgo,
                "promedios_por_est": promedios_por_est,
                "historial": historial,
                "tendencia_actividades": tendencia_actividades,
                "chk_states": chk_states
            }
            
            # Enviar al hilo principal para renderizado de widgets
            try:
                self.after(0, lambda: self._dibujar_con_datos(datos))
            except Exception:
                pass

        # Iniciar hilo secundario
        t = threading.Thread(target=bg_work, daemon=True)
        t.start()

    def _dibujar_con_datos(self, datos):
        # Destruir indicador de carga
        if hasattr(self, "lbl_loading") and self.lbl_loading.winfo_exists():
            self.lbl_loading.destroy()

        grado = datos["grado"]
        materia = datos["materia"]
        trimestre = datos["trimestre"]
        estudiante_sel = datos["estudiante_sel"]
        aprobados = datos["aprobados"]
        reprobados = datos["reprobados"]
        en_riesgo = datos["en_riesgo"]
        promedios_por_est = datos["promedios_por_est"]
        historial = datos["historial"]
        chk = datos["chk_states"]

        if estudiante_sel != "Todos los Estudiantes":
            # Vista individual
            self.dibujar_proyeccion(estudiante_sel, historial, 0, 0, colspan=2)
        else:
            # Vista general
            row = 0
            col = 0
            if chk["pastel"]:
                self.dibujar_pastel(aprobados, en_riesgo, reprobados, row, col)
                col += 1
                if col > 1:
                    col = 0
                    row += 1
            if chk["barras"]:
                self.dibujar_barras(promedios_por_est, row, col)
                col += 1
                if col > 1:
                    col = 0
                    row += 1
            if chk["tendencia"]:
                if col == 1:
                    col = 0
                    row += 1
                self.dibujar_tendencia(datos.get("tendencia_actividades", []), row, col, colspan=2)
                col = 0
                row += 1
            if chk["hist"]:
                self.dibujar_histograma(list(promedios_por_est.values()), row, col)
                col += 1
                if col > 1:
                    col = 0
                    row += 1
            if chk["pareto"]:
                self.dibujar_pareto(grado, row, col)
                col += 1
                if col > 1:
                    col = 0
                    row += 1
            if chk["scatter"]:
                self.dibujar_scatter(grado, promedios_por_est, row, col)
                col += 1
                if col > 1:
                    col = 0
                    row += 1
            if chk["box"]:
                self.dibujar_boxplot(grado, materia, promedios_por_est, row, col)
                col += 1
                if col > 1:
                    col = 0
                    row += 1
            if chk["lineas"]:
                self.dibujar_lineas_trimestrales(grado, materia, row, col)
                col += 1
                if col > 1:
                    col = 0
                    row += 1
            if chk["heat"]:
                self.dibujar_heatmap(grado, row, col, colspan=2 if col == 0 else 1)
                col += 1
                if col > 1:
                    col = 0
                    row += 1
            if chk.get("habitos", False):
                self.dibujar_habitos(grado, trimestre, row, col, colspan=2 if col == 0 else 1)

    def dibujar_proyeccion(self, nombre, historial, row, col, colspan=1):
        import numpy as np

        f_grafico = ctk.CTkFrame(self.scroll_canvas, fg_color=self.C["card"], corner_radius=10)
        f_grafico.grid(row=row, column=col, columnspan=colspan, sticky="nsew", padx=10, pady=10)

        fig = Figure(figsize=(10, 4), dpi=100)
        ax = fig.add_subplot(111)
        fig.patch.set_facecolor(self.C["card"])
        ax.set_facecolor(self.C["fondo"])
        self.fig_objs.append(fig)

        if not historial or len(historial) < 2:
            ax.text(0.5, 0.5, "Sin Datos", ha='center', va='center', color='white', fontsize=16)
        else:
            x = np.arange(1, len(historial) + 1)
            y = np.array(historial)

            # Regresion lineal sin dependencia de scipy.
            slope, intercept = np.polyfit(x, y, 1)
            x_pred = np.arange(1, len(historial) + 3) # Proyectar 2 periodos mas
            y_pred = intercept + slope * x_pred

            ax.plot(x, y, marker='o', color=self.C["cian"], label='Historial de Notas', linewidth=2, markersize=8)
            ax.plot(x_pred, y_pred, linestyle='--', color=self.C["amarillo"], label='Tendencia / Predicción', linewidth=2)

            ax.axhline(y=3.0, color=self.C["rojo"], linestyle=':', alpha=0.7, label='Límite Aprobación (3.0)')

            ax.set_ylim(1.0, 5.2)
            ax.set_xticks(x_pred)
            ax.set_xticklabels([f"T{i}" for i in x] + ["Pred 1", "Pred 2"])

            ax.legend(facecolor=self.C["fondo"], edgecolor=self.C["card"], labelcolor="white")

        ax.tick_params(axis='x', colors=self.C["texto"])
        ax.tick_params(axis='y', colors=self.C["texto"])

        ax.spines['bottom'].set_color(self.C["texto"])
        ax.spines['left'].set_color(self.C["texto"])
        ax.spines['top'].set_visible(False)
        ax.spines['right'].set_visible(False)

        ax.set_title(f"Proyección y Predicción de Rendimiento: {nombre}", color=self.C["cian"], pad=15, fontweight="bold")
        fig.tight_layout()

        canvas = FigureCanvasTkAgg(fig, master=f_grafico)
        canvas.draw()
        canvas.get_tk_widget().pack(fill="both", expand=True, padx=10, pady=10)
        ctk.CTkButton(f_grafico, text="⬇️ a Word", width=80, height=24, fg_color="#334155", hover_color="#475569", command=lambda f=fig, t=ax.get_title(): self.exportar_grafico_individual(f, t)).place(relx=0.98, rely=0.02, anchor="ne")
    def dibujar_pastel(self, aprobados, riesgo, reprobados, row, col):
        f_grafico = ctk.CTkFrame(self.scroll_canvas, fg_color=self.C["card"], corner_radius=10)
        f_grafico.grid(row=row, column=col, sticky="nsew", padx=10, pady=10)

        fig = Figure(figsize=(5, 4), dpi=100)
        ax = fig.add_subplot(111)
        fig.patch.set_facecolor(self.C["card"])
        ax.set_facecolor(self.C["card"])
        self.fig_objs.append(fig)

        etiquetas = ['Aprobados (>=3.0)', 'En Riesgo (2.5-2.9)', 'Reprobados (<2.5)']
        valores = [aprobados, riesgo, reprobados]
        colores = [self.C["verde"], self.C["amarillo"], self.C["rojo"]]

        # Filtrar ceros
        e_filt, v_filt, c_filt = [], [], []
        for e, v, c in zip(etiquetas, valores, colores):
            if v > 0:
                e_filt.append(e); v_filt.append(v); c_filt.append(c)

        if sum(valores) == 0:
            ax.text(0.5, 0.5, "Sin Datos", ha='center', va='center', color='white')
        else:
            wedges, texts, autotexts = ax.pie(v_filt, labels=e_filt, colors=c_filt, autopct='%1.1f%%',
                                            startangle=90, textprops={'color': "white", 'fontsize': 10})

            # Mejorar visual
            for w in wedges: w.set_edgecolor(self.C["card"])

        ax.set_title("Distribución General del Salón", color=self.C["cian"], pad=20, fontweight="bold")

        canvas = FigureCanvasTkAgg(fig, master=f_grafico)
        canvas.draw()
        canvas.get_tk_widget().pack(fill="both", expand=True, padx=10, pady=10)
        ctk.CTkButton(f_grafico, text="⬇️ a Word", width=80, height=24, fg_color="#334155", hover_color="#475569", command=lambda f=fig, t=ax.get_title(): self.exportar_grafico_individual(f, t)).place(relx=0.98, rely=0.02, anchor="ne")

    def dibujar_barras(self, promedios_dict, row, col):
        f_grafico = ctk.CTkFrame(self.scroll_canvas, fg_color=self.C["card"], corner_radius=10)
        f_grafico.grid(row=row, column=col, sticky="nsew", padx=10, pady=10)

        fig = Figure(figsize=(6, 4), dpi=100)
        ax = fig.add_subplot(111)
        fig.patch.set_facecolor(self.C["card"])
        ax.set_facecolor(self.C["fondo"])
        self.fig_objs.append(fig)

        # Tomar solo los primeros 10-15 para no saturar si son muchos
        nombres = list(promedios_dict.keys())[:15]
        notas = [promedios_dict[n] for n in nombres]

        # Nombres cortos para X
        nombres_cortos = [n.split(" ")[0] for n in nombres]

        # Colores dinámicos basados en la nota
        colores = [self.C["rojo"] if n < 3.0 else self.C["verde"] for n in notas]

        barras = ax.bar(nombres_cortos, notas, color=colores, edgecolor=self.C["cian"], alpha=0.8)

        # Línea de pase (3.0)
        ax.axhline(y=3.0, color=self.C["amarillo"], linestyle='--', alpha=0.7)

        ax.tick_params(axis='x', colors=self.C["texto"], rotation=45)
        ax.tick_params(axis='y', colors=self.C["texto"])
        ax.set_ylim(1.0, 5.2)
        
        ax.spines['bottom'].set_color(self.C["texto"])
        ax.spines['left'].set_color(self.C["texto"])
        ax.spines['top'].set_visible(False)
        ax.spines['right'].set_visible(False)

        ax.set_title("Top 15 Estudiantes (Promedios)", color=self.C["cian"], pad=15, fontweight="bold")

        # Ajustar margen para que los nombres rotados entren bien
        fig.tight_layout()

        canvas = FigureCanvasTkAgg(fig, master=f_grafico)
        canvas.draw()
        canvas.get_tk_widget().pack(fill="both", expand=True, padx=10, pady=10)
        ctk.CTkButton(f_grafico, text="⬇️ a Word", width=80, height=24, fg_color="#334155", hover_color="#475569", command=lambda f=fig, t=ax.get_title(): self.exportar_grafico_individual(f, t)).place(relx=0.98, rely=0.02, anchor="ne")

    def dibujar_tendencia(self, tendencia_actividades, row, col, colspan=1):
        f_grafico = ctk.CTkFrame(self.scroll_canvas, fg_color=self.C["card"], corner_radius=10)
        f_grafico.grid(row=row, column=col, columnspan=colspan, sticky="nsew", padx=10, pady=10)

        fig = Figure(figsize=(10, 3.5), dpi=100)
        ax = fig.add_subplot(111)
        fig.patch.set_facecolor(self.C["card"])
        ax.set_facecolor(self.C["fondo"])
        self.fig_objs.append(fig)

        if not tendencia_actividades:
            ax.text(0.5, 0.5, "Seleccione un Grado y Materia para ver Tendencia Temporal", ha='center', va='center', color='white', fontsize=12)
        else:
            labels = [x[0] for x in tendencia_actividades]
            valores = [x[1] for x in tendencia_actividades]
            x_ind = list(range(len(labels)))
            
            ax.plot(x_ind, valores, marker='o', color=self.C["cian"], linewidth=2, markersize=6, label='Promedio Actividad')
            ax.axhline(y=3.0, color=self.C["rojo"], linestyle=':', alpha=0.7, label='Límite de Aprobación (3.0)')
            
            ax.set_xticks(x_ind)
            ax.set_xticklabels(labels, rotation=15, ha='right', fontsize=8)
            ax.set_ylim(1.0, 5.2)
            ax.legend(facecolor=self.C["fondo"], edgecolor=self.C["card"], labelcolor="white", fontsize=8)
            ax.grid(color=self.C["borde"] if "borde" in self.C else "#334155", linestyle='--', alpha=0.3)

        ax.tick_params(axis='x', colors=self.C["texto"])
        ax.tick_params(axis='y', colors=self.C["texto"])
        ax.spines['bottom'].set_color(self.C["texto"])
        ax.spines['left'].set_color(self.C["texto"])
        ax.spines['top'].set_visible(False)
        ax.spines['right'].set_visible(False)

        ax.set_title("Evolución y Tendencia Temporal del Grupo por Actividad", color=self.C["cian"], pad=15, fontweight="bold")
        fig.tight_layout()

        canvas = FigureCanvasTkAgg(fig, master=f_grafico)
        canvas.draw()
        canvas.get_tk_widget().pack(fill="both", expand=True, padx=10, pady=10)
        ctk.CTkButton(f_grafico, text="⬇️ a Word", width=80, height=24, fg_color="#334155", hover_color="#475569", command=lambda f=fig, t=ax.get_title(): self.exportar_grafico_individual(f, t)).place(relx=0.98, rely=0.02, anchor="ne")



    def dibujar_histograma(self, notas, row, col, colspan=1):
        f_grafico = ctk.CTkFrame(self.scroll_canvas, fg_color=self.C["card"], corner_radius=10)
        f_grafico.grid(row=row, column=col, columnspan=colspan, sticky="nsew", padx=10, pady=10)

        fig = Figure(figsize=(6, 4), dpi=100)
        ax = fig.add_subplot(111)
        fig.patch.set_facecolor(self.C["card"])
        ax.set_facecolor(self.C["fondo"])
        self.fig_objs.append(fig)

        if not notas:
            ax.text(0.5, 0.5, "Sin Datos", ha='center', va='center', color='white')
        else:
            ax.hist(notas, bins=[1.0, 2.0, 2.5, 3.0, 4.0, 5.0], color=self.C["purpura"], alpha=0.7, edgecolor='white')

        ax.tick_params(axis='x', colors=self.C["texto"])
        ax.tick_params(axis='y', colors=self.C["texto"])
        ax.spines['bottom'].set_color(self.C["texto"])
        ax.spines['left'].set_color(self.C["texto"])
        ax.spines['top'].set_visible(False)
        ax.spines['right'].set_visible(False)

        ax.set_title("Histograma de Distribución de Notas", color=self.C["cian"], pad=15, fontweight="bold")
        fig.tight_layout()

        canvas = FigureCanvasTkAgg(fig, master=f_grafico)
        canvas.draw()
        canvas.get_tk_widget().pack(fill="both", expand=True, padx=10, pady=10)
        ctk.CTkButton(f_grafico, text="⬇️ a Word", width=80, height=24, fg_color="#334155", hover_color="#475569", command=lambda f=fig, t=ax.get_title(): self.exportar_grafico_individual(f, t)).place(relx=0.98, rely=0.02, anchor="ne")

    def dibujar_pareto(self, grado, row, col, colspan=1):
        f_grafico = ctk.CTkFrame(self.scroll_canvas, fg_color=self.C["card"], corner_radius=10)
        f_grafico.grid(row=row, column=col, columnspan=colspan, sticky="nsew", padx=10, pady=10)

        fig = Figure(figsize=(6, 4), dpi=100)
        ax = fig.add_subplot(111)
        ax2 = ax.twinx()
        fig.patch.set_facecolor(self.C["card"])
        ax.set_facecolor(self.C["fondo"])
        self.fig_objs.append(fig)

        if grado == "Todos los Grados":
            grados_activos = self.engine.obtener_grados_activos()
        else:
            grados_activos = [grado]

        # Obtener lista unificada de materias
        todas_materias = []
        for g in grados_activos:
            todas_materias.extend(self.engine.obtener_materias_por_grado(g))
        materias = sorted(list(set(todas_materias)))
        if "Sin materias registradas" in materias:
            materias.remove("Sin materias registradas")

        if not materias:
            ax.text(0.5, 0.5, "Sin Datos", ha='center', va='center', color='white')
        else:
            reprobados_por_materia = {}
            for mat in materias:
                reps = 0
                for g in grados_activos:
                    promedios = getattr(self.engine, 'obtener_promedios_reales', lambda gr,m,t: {})(g, mat, "Anual")
                    if not promedios:
                        # Intenta cualquier trimestre si Anual está vacío
                        promedios = getattr(self.engine, 'obtener_promedios_reales', lambda gr,m,t: {})(g, mat, "Trimestre 1")
                    reps += sum(1 for v in promedios.values() if v < 3.0)
                reprobados_por_materia[mat] = reps

            reprobados_ordenado = dict(sorted(reprobados_por_materia.items(), key=lambda item: item[1], reverse=True))
            materias_nombres = list(reprobados_ordenado.keys())[:10]
            materias_cortas = [m[:10] for m in materias_nombres]
            conteos = [reprobados_ordenado[m] for m in materias_nombres]

            if sum(conteos) == 0:
                ax.text(0.5, 0.5, "Sin Reprobados", ha='center', va='center', color='white')
            else:
                ax.bar(materias_cortas, conteos, color=self.C["rojo"], alpha=0.8)

                acumulado = []
                total = sum(conteos)
                suma = 0
                for c in conteos:
                    suma += c
                    acumulado.append(suma / total * 100)

                ax2.plot(materias_cortas, acumulado, color=self.C["amarillo"], marker="D", ms=5)
                ax2.set_ylim(0, 110)
                ax2.tick_params(axis='y', colors=self.C["amarillo"])

        ax.tick_params(axis='x', colors=self.C["texto"], rotation=45)
        ax.tick_params(axis='y', colors=self.C["texto"])
        ax.spines['bottom'].set_color(self.C["texto"])
        ax.spines['left'].set_color(self.C["texto"])
        ax.spines['top'].set_visible(False)
        ax2.spines['top'].set_visible(False)
        ax2.spines['right'].set_color(self.C["amarillo"])

        ax.set_title("Gráfico de Pareto (Reprobación)", color=self.C["cian"], pad=15, fontweight="bold")
        fig.tight_layout()

        canvas = FigureCanvasTkAgg(fig, master=f_grafico)
        canvas.draw()
        canvas.get_tk_widget().pack(fill="both", expand=True, padx=10, pady=10)
        ctk.CTkButton(f_grafico, text="⬇️ a Word", width=80, height=24, fg_color="#334155", hover_color="#475569", command=lambda f=fig, t=ax.get_title(): self.exportar_grafico_individual(f, t)).place(relx=0.98, rely=0.02, anchor="ne")

    def dibujar_scatter(self, grado, promedios_dict, row, col, colspan=1):
        f_grafico = ctk.CTkFrame(self.scroll_canvas, fg_color=self.C["card"], corner_radius=10)
        f_grafico.grid(row=row, column=col, columnspan=colspan, sticky="nsew", padx=10, pady=10)

        fig = Figure(figsize=(6, 4), dpi=100)
        ax = fig.add_subplot(111)
        fig.patch.set_facecolor(self.C["card"])
        ax.set_facecolor(self.C["fondo"])
        self.fig_objs.append(fig)

        if not promedios_dict:
            ax.text(0.5, 0.5, "Sin Datos", ha='center', va='center', color='white')
        else:
            x = []
            y = []
            grados_info = {}
            # Determinar trimestre de asistencia contextualmente
            asis_trim = "Trimestre 1"
            sel_trim = self.combo_trimestre.get()
            if sel_trim in ["Trimestre 1", "Trimestre 2", "Trimestre 3"]:
                asis_trim = sel_trim
            for n, p in promedios_dict.items():
                est_grado = grado
                est_name = n
                if "(" in n and n.endswith(")"):
                    parts = n.rsplit(" (", 1)
                    if len(parts) == 2:
                        est_name = parts[0]
                        est_grado = parts[1][:-1]
                
                if est_grado not in grados_info:
                    try:
                        fechas = self.engine.obtener_fechas_asistencia(est_grado, asis_trim)
                        estudiantes = self.engine.obtener_estudiantes_completos(est_grado)
                        name_to_id = {e["nombre"].upper().strip(): e["id"] for e in estudiantes}
                        grados_info[est_grado] = {
                            "fechas": fechas,
                            "name_to_id": name_to_id,
                            "asis_cache": {}
                        }
                    except:
                        grados_info[est_grado] = None
                
                info = grados_info[est_grado]
                asistencia = 100
                if info and info["fechas"]:
                    presentes = 0
                    total = len(info["fechas"])
                    for fecha in info["fechas"]:
                        if fecha not in info["asis_cache"]:
                            res = getattr(self.engine, 'buscar_asistencia_existente', lambda g,t,f: None)(est_grado, asis_trim, fecha)
                            info["asis_cache"][fecha] = res["asistencia"] if res else {}
                        
                        asis_dict = info["asis_cache"][fecha]
                        est_id = info["name_to_id"].get(est_name.upper().strip())
                        est_idx = int(est_id) % 100 if str(est_id).isdigit() else est_id
                        val = asis_dict.get(est_idx, "P") if est_idx is not None else "P"
                        if val in ["P", "T"]:
                            presentes += 1
                    asistencia = (presentes / total) * 100 if total > 0 else 100
                
                x.append(asistencia)
                y.append(p)

            ax.scatter(x, y, color=self.C["cian"], alpha=0.7)
            ax.set_xlabel("% Asistencia (Estimado)", color=self.C["texto"])
            ax.set_ylabel("Nota Promedio", color=self.C["texto"])

        ax.tick_params(axis='x', colors=self.C["texto"])
        ax.tick_params(axis='y', colors=self.C["texto"])
        ax.spines['bottom'].set_color(self.C["texto"])
        ax.spines['left'].set_color(self.C["texto"])
        ax.spines['top'].set_visible(False)
        ax.spines['right'].set_visible(False)

        ax.set_title("Scatter Plot: Notas vs Asistencia", color=self.C["cian"], pad=15, fontweight="bold")
        fig.tight_layout()

        canvas = FigureCanvasTkAgg(fig, master=f_grafico)
        canvas.draw()
        canvas.get_tk_widget().pack(fill="both", expand=True, padx=10, pady=10)
        ctk.CTkButton(f_grafico, text="⬇️ a Word", width=80, height=24, fg_color="#334155", hover_color="#475569", command=lambda f=fig, t=ax.get_title(): self.exportar_grafico_individual(f, t)).place(relx=0.98, rely=0.02, anchor="ne")

    def dibujar_boxplot(self, grado, materia, promedios_dict, row, col, colspan=1):
        f_grafico = ctk.CTkFrame(self.scroll_canvas, fg_color=self.C["card"], corner_radius=10)
        f_grafico.grid(row=row, column=col, columnspan=colspan, sticky="nsew", padx=10, pady=10)

        fig = Figure(figsize=(6, 4), dpi=100)
        ax = fig.add_subplot(111)
        fig.patch.set_facecolor(self.C["card"])
        ax.set_facecolor(self.C["fondo"])
        self.fig_objs.append(fig)

        notas = list(promedios_dict.values())
        if not notas:
            ax.text(0.5, 0.5, "Sin Datos", ha='center', va='center', color='white')
        else:
            bplot = ax.boxplot(notas, patch_artist=True, vert=True)
            for patch in bplot['boxes']:
                patch.set_facecolor(self.C["verde"])
            for median in bplot['medians']:
                median.set(color=self.C["card"], linewidth=2)
            ax.set_xticklabels([materia[:10] if materia != "Todas las Materias" else grado])

        ax.tick_params(axis='x', colors=self.C["texto"])
        ax.tick_params(axis='y', colors=self.C["texto"])
        ax.spines['bottom'].set_color(self.C["texto"])
        ax.spines['left'].set_color(self.C["texto"])
        ax.spines['top'].set_visible(False)
        ax.spines['right'].set_visible(False)

        ax.set_title("Box-plot Comparativo", color=self.C["cian"], pad=15, fontweight="bold")
        fig.tight_layout()

        canvas = FigureCanvasTkAgg(fig, master=f_grafico)
        canvas.draw()
        canvas.get_tk_widget().pack(fill="both", expand=True, padx=10, pady=10)
        ctk.CTkButton(f_grafico, text="⬇️ a Word", width=80, height=24, fg_color="#334155", hover_color="#475569", command=lambda f=fig, t=ax.get_title(): self.exportar_grafico_individual(f, t)).place(relx=0.98, rely=0.02, anchor="ne")

    def dibujar_lineas_trimestrales(self, grado, materia, row, col, colspan=1):
        f_grafico = ctk.CTkFrame(self.scroll_canvas, fg_color=self.C["card"], corner_radius=10)
        f_grafico.grid(row=row, column=col, columnspan=colspan, sticky="nsew", padx=10, pady=10)

        fig = Figure(figsize=(6, 4), dpi=100)
        ax = fig.add_subplot(111)
        fig.patch.set_facecolor(self.C["card"])
        ax.set_facecolor(self.C["fondo"])
        self.fig_objs.append(fig)

        if grado == "Todos los Grados":
            grados_activos = self.engine.obtener_grados_activos()
        else:
            grados_activos = [grado]

        t1_vals = []
        t2_vals = []
        t3_vals = []

        for g in grados_activos:
            p_t1 = getattr(self.engine, 'obtener_promedios_reales', lambda gr,m,t: {})(g, materia, "Trimestre 1")
            p_t2 = getattr(self.engine, 'obtener_promedios_reales', lambda gr,m,t: {})(g, materia, "Trimestre 2")
            p_t3 = getattr(self.engine, 'obtener_promedios_reales', lambda gr,m,t: {})(g, materia, "Trimestre 3")
            t1_vals.extend(p_t1.values())
            t2_vals.extend(p_t2.values())
            t3_vals.extend(p_t3.values())

        t1_avg = sum(t1_vals) / len(t1_vals) if t1_vals else 0
        t2_avg = sum(t2_vals) / len(t2_vals) if t2_vals else 0
        t3_avg = sum(t3_vals) / len(t3_vals) if t3_vals else 0

        y = [t1_avg, t2_avg, t3_avg]
        x = [1, 2, 3]

        if not any(y):
            ax.text(0.5, 0.5, "Sin Datos", ha='center', va='center', color='white')
        else:
            ax.plot(x, y, marker='o', color=self.C["amarillo"], linewidth=2, markersize=8)
            ax.set_xticks(x)
            ax.set_xticklabels(["T1", "T2", "T3"])
            ax.set_ylim(1.0, 5.2)

        ax.tick_params(axis='x', colors=self.C["texto"])
        ax.tick_params(axis='y', colors=self.C["texto"])
        ax.spines['bottom'].set_color(self.C["texto"])
        ax.spines['left'].set_color(self.C["texto"])
        ax.spines['top'].set_visible(False)
        ax.spines['right'].set_visible(False)

        ax.set_title("Evolución Trimestral", color=self.C["cian"], pad=15, fontweight="bold")
        fig.tight_layout()

        canvas = FigureCanvasTkAgg(fig, master=f_grafico)
        canvas.draw()
        canvas.get_tk_widget().pack(fill="both", expand=True, padx=10, pady=10)
        ctk.CTkButton(f_grafico, text="⬇️ a Word", width=80, height=24, fg_color="#334155", hover_color="#475569", command=lambda f=fig, t=ax.get_title(): self.exportar_grafico_individual(f, t)).place(relx=0.98, rely=0.02, anchor="ne")

    def dibujar_heatmap(self, grado, row, col, colspan=1):
        import numpy as np
        f_grafico = ctk.CTkFrame(self.scroll_canvas, fg_color=self.C["card"], corner_radius=10)
        f_grafico.grid(row=row, column=col, columnspan=colspan, sticky="nsew", padx=10, pady=10)

        fig = Figure(figsize=(10, 4), dpi=100)
        ax = fig.add_subplot(111)
        fig.patch.set_facecolor(self.C["card"])
        ax.set_facecolor(self.C["fondo"])
        self.fig_objs.append(fig)

        if grado == "Todos los Grados":
            grados_activos = self.engine.obtener_grados_activos()
            if grados_activos:
                grado = grados_activos[0]

        materias = self.engine.obtener_materias_por_grado(grado)
        estudiantes = []
        try:
            estudiantes = self.engine.obtener_estudiantes_completos(grado)
        except Exception: pass

        if not materias or materias == ["Sin materias"] or not estudiantes:
            ax.text(0.5, 0.5, "Sin Datos", ha='center', va='center', color='white')
        else:
            materias_cortas = [m[:10] for m in materias[:5]] # Limitar a 5 para no colapsar
            nombres_cortos = [e['nombre'].split(" ")[0] for e in estudiantes[:10]] # Limitar a 10

            # Optimización: Obtener mapa de promedios fuera del bucle de estudiantes
            promedios_por_materia = {}
            for m in materias[:5]:
                promedios = getattr(self.engine, 'obtener_promedios_reales', lambda g,ma,t: {})(grado, m, "Anual")
                if not promedios:
                    promedios = getattr(self.engine, 'obtener_promedios_reales', lambda g,ma,t: {})(grado, m, "Trimestre 1")
                promedios_por_materia[m] = promedios

            data = []
            for e in estudiantes[:10]:
                row_data = []
                for m in materias[:5]:
                    val = promedios_por_materia[m].get(e['nombre'], 0.0)
                    row_data.append(val)
                data.append(row_data)

            data_np = np.array(data)
            im = ax.imshow(data_np, cmap="RdYlGn", vmin=1.0, vmax=5.0, aspect='auto')

            ax.set_xticks(np.arange(len(materias_cortas)))
            ax.set_yticks(np.arange(len(nombres_cortos)))
            ax.set_xticklabels(materias_cortas, color=self.C["texto"])
            ax.set_yticklabels(nombres_cortos, color=self.C["texto"])

            for i in range(len(nombres_cortos)):
                for j in range(len(materias_cortas)):
                    val = data_np[i, j]
                    if val > 0:
                        text = ax.text(j, i, f"{val:.1f}", ha="center", va="center", color="black" if 2.5 <= val <= 4.0 else "white")

            fig.colorbar(im, ax=ax)

        ax.set_title("Heatmap: Notas x Estudiante y Materia (Top 10/5)", color=self.C["cian"], pad=15, fontweight="bold")
        fig.tight_layout()

        canvas = FigureCanvasTkAgg(fig, master=f_grafico)
        canvas.draw()
        canvas.get_tk_widget().pack(fill="both", expand=True, padx=10, pady=10)
        ctk.CTkButton(f_grafico, text="⬇️ a Word", width=80, height=24, fg_color="#334155", hover_color="#475569", command=lambda f=fig, t=ax.get_title(): self.exportar_grafico_individual(f, t)).place(relx=0.98, rely=0.02, anchor="ne")

    def mostrar_guia_graficas(self):
        win = ctk.CTkToplevel(self)
        win.title("Guía de Uso de Gráficos Académicos")
        win.geometry("750x650")
        win.configure(fg_color=self.C["fondo"])
        win.transient(self)
        win.grab_set()
        
        from config import establecer_icono_ventana
        establecer_icono_ventana(win)
        
        ctk.CTkLabel(
            win,
            text="📘 Guía de Interpretación de Gráficos",
            font=("Segoe UI", 20, "bold"),
            text_color=self.C["cian"]
        ).pack(pady=(20, 10))
        
        sf = ctk.CTkScrollableFrame(
            win,
            fg_color=self.C["card_alt"],
            scrollbar_button_color=self.C["cian"],
            scrollbar_fg_color=self.C["card_alt"]
        )
        sf.pack(fill="both", expand=True, padx=20, pady=(10, 20))
        
        guias = [
            ("📊 Gráfico de Pastel", 
             "Muestra la proporción de estudiantes que están aprobando (nota ≥ 3.0) frente a los que están reprobando (nota < 3.0).\n\n"
             "💡 Uso Pedagógico: Permite visualizar de forma instantánea la tasa general de éxito de un aula de clases."),
            
            ("📶 Gráfico de Barras", 
             "Representa el promedio individual de cada estudiante en la asignatura seleccionada.\n\n"
             "💡 Uso Pedagógico: Facilita la comparación directa entre alumnos y la identificación de casos con rendimiento destacado o rezago."),
            
            ("📈 Tendencia de Rendimiento", 
             "Muestra la evolución del promedio del grupo (o del estudiante seleccionado) a lo largo de los trimestres (T1 ➔ T2 ➔ T3).\n\n"
             "💡 Uso Pedagógico: Ideal para evaluar si las estrategias educativas implementadas están dando frutos a lo largo del año."),
            
            ("〰️ Comparativa Trimestral", 
             "Compara las curvas de calificaciones individuales de cada alumno de manera directa a través del tiempo.\n\n"
             "💡 Uso Pedagógico: Permite hacer seguimiento individualizado detallado sobre la evolución trimestral de los alumnos.")
        ]
        
        for titulo, desc in guias:
            card = ctk.CTkFrame(sf, fg_color=self.C["card"], border_width=1, border_color=self.C["cian"], corner_radius=8)
            card.pack(fill="x", padx=10, pady=8)
            
            ctk.CTkLabel(
                card, 
                text=titulo, 
                font=("Segoe UI", 14, "bold"), 
                text_color=self.C["cian"]
            ).pack(anchor="w", padx=15, pady=(10, 5))
            
            ctk.CTkLabel(
                card, 
                text=desc, 
                font=("Segoe UI", 12), 
                text_color=self.C["texto"],
                justify="left",
                wraplength=640
            ).pack(anchor="w", padx=15, pady=(0, 10))
            
        ctk.CTkButton(
            win,
            text="Entendido",
            fg_color=self.C["cian"],
            text_color=self.C["fondo"],
            hover_color=self.C["hover"],
            font=("Segoe UI", 12, "bold"),
            command=win.destroy
        ).pack(pady=(10, 15))

    def dibujar_habitos(self, grado, trimestre, row, col, colspan=1):
        import os
        f_grafico = ctk.CTkFrame(self.scroll_canvas, fg_color=self.C["card"], corner_radius=10)
        if colspan > 1:
            f_grafico.grid(row=row, column=col, columnspan=colspan, sticky="nsew", padx=10, pady=10)
        else:
            f_grafico.grid(row=row, column=col, sticky="nsew", padx=10, pady=10)

        fig = Figure(figsize=(6.5, 4), dpi=100)
        ax = fig.add_subplot(111)
        fig.patch.set_facecolor(self.C["card"])
        ax.set_facecolor(self.C["card"])
        self.fig_objs.append(fig)

        import json
        ruta_json = os.path.abspath(os.path.join(os.path.dirname(self.engine.ruta), "Expedientes_Estudiantes", "habitos_evaluaciones.json"))
        criterios_dict = {}
        try:
            if os.path.exists(ruta_json):
                with open(ruta_json, "r", encoding="utf-8") as f:
                    habitos_data = json.load(f)
                
                for key, val_entry in habitos_data.items():
                    parts = key.split("::")
                    k_grado = parts[0]
                    k_trim = parts[1] if len(parts) > 1 else ""
                    
                    if (grado == "Todos los Grados" or grado.replace("°","") in k_grado.replace("°","")) and \
                       (trimestre == "Todos los Trimestres" or k_trim.lower() == trimestre.lower()):
                        
                        est_evals = val_entry.get("estudiantes", {})
                        for est_id, crit_vals in est_evals.items():
                            for crit, score in crit_vals.items():
                                if crit not in criterios_dict:
                                    criterios_dict[crit] = {"S": 0, "R": 0, "X": 0}
                                if score in ["S", "R", "X"]:
                                    criterios_dict[crit][score] += 1
        except Exception as e:
            print(f"Error cargando habitos para grafico: {e}")

        if not criterios_dict:
            ax.text(0.5, 0.5, "Sin Datos de Hábitos", ha='center', va='center', color='white')
            ax.set_axis_off()
        else:
            criterios = list(criterios_dict.keys())
            s_vals = [criterios_dict[c]["S"] for c in criterios]
            r_vals = [criterios_dict[c]["R"] for c in criterios]
            x_vals = [criterios_dict[c]["X"] for c in criterios]

            import numpy as np
            ind = np.arange(len(criterios))
            width = 0.5

            # Stacked bars
            ax.bar(ind, s_vals, width, label="S (Satisfactorio)", color=self.C["verde"])
            ax.bar(ind, r_vals, width, bottom=s_vals, label="R (Regular)", color=self.C["amarillo"])
            bottom_x = [s + r for s, r in zip(s_vals, r_vals)]
            ax.bar(ind, x_vals, width, bottom=bottom_x, label="X (No Satisface)", color=self.C["rojo"])

            criterios_clean = [c.replace(" y ", "\ny ").replace(" en ", "\nen ").replace(" a la ", "\na la ").replace(" de la ", "\nde la ") for c in criterios]
            ax.set_xticks(ind)
            ax.set_xticklabels(criterios_clean, color="white", fontsize=7, rotation=35, ha='right')
            ax.tick_params(colors="white", labelsize=8)
            ax.legend(facecolor=self.C["card"], edgecolor=self.C["borde"], labelcolor="white", fontsize=8)
            ax.grid(axis='y', color=self.C["borde"], linestyle='--', alpha=0.5)
            
            for spine in ['top', 'right']:
                ax.spines[spine].set_visible(False)
            ax.spines['left'].set_color(self.C["borde"])
            ax.spines['bottom'].set_color(self.C["borde"])

        ax.set_title("Hábitos y Actitudes del Salón", color=self.C["cian"], pad=20, fontweight="bold")

        canvas = FigureCanvasTkAgg(fig, master=f_grafico)
        canvas.draw()
        canvas.get_tk_widget().pack(fill="both", expand=True, padx=10, pady=10)
        ctk.CTkButton(f_grafico, text="⬇️ a Word", width=80, height=24, fg_color="#334155", hover_color="#475569", command=lambda f=fig, t=ax.get_title(): self.exportar_grafico_individual(f, t)).place(relx=0.98, rely=0.02, anchor="ne")
