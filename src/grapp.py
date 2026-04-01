import customtkinter as ctk
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

class GraficosFrame(ctk.CTkFrame):
    def __init__(self, master, engine, **kwargs):
        super().__init__(master, fg_color="transparent", **kwargs)
        self.engine = engine

        # Tema de colores
        self.C = {
            "fondo":        "#0A1628",
            "card":         "#0F2744",
            "cian":         "#00DDEB",
            "verde":        "#00FF88",
            "rojo":         "#FF4444",
            "amarillo":     "#FFD700",
            "texto":        "#E2E8F0",
        }

        # Estructura principal
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)

        self.crear_filtros()

        # Contenedor de gráficos con scroll
        self.scroll_canvas = ctk.CTkScrollableFrame(self, fg_color="transparent")
        self.scroll_canvas.grid(row=1, column=0, sticky="nsew", padx=10, pady=10)
        self.fig_objs = []
        self.scroll_canvas.grid_columnconfigure((0, 1), weight=1)

        self.fig_objs = [] # Para limpiar figuras después

        # Iniciar gráficos
        self.actualizar_graficos()

    def crear_filtros(self):
        f_top = ctk.CTkFrame(self, fg_color=self.C["card"], corner_radius=10)
        f_top.grid(row=0, column=0, sticky="ew", padx=15, pady=15)
        
        ctk.CTkLabel(f_top, text="📊 Análisis Visual de Rendimiento",
                    font=("Segoe UI", 20, "bold"), text_color=self.C["cian"]).pack(side="left", padx=20, pady=15)

        # Grado
        grados = self.engine.obtener_grados_activos()
        if not grados: grados = ["Sin datos"]

        self.combo_grado = ctk.CTkOptionMenu(f_top, values=grados, command=self.al_cambiar_grado)
        self.combo_grado.pack(side="left", padx=5)
        self.combo_grado.set(grados[0])

        # Materia
        materias = self.engine.obtener_materias_por_grado(grados[0]) if grados[0] != "Sin datos" else ["Sin materias"]
        self.combo_materia = ctk.CTkOptionMenu(f_top, values=materias, command=lambda _: self.actualizar_graficos())
        self.combo_materia.pack(side="left", padx=5)
        if materias: self.combo_materia.set(materias[0])

        # Estudiante
        self.combo_estudiante = ctk.CTkOptionMenu(f_top, values=["Todos los Estudiantes"], command=lambda _: self.actualizar_graficos())
        self.combo_estudiante.pack(side="left", padx=5)

        # Trimestre (si aplica)
        self.combo_trimestre = ctk.CTkOptionMenu(f_top, values=["Trimestre 1", "Trimestre 2", "Trimestre 3"], command=lambda _: self.actualizar_graficos())
        self.combo_trimestre.pack(side="left", padx=5)

        ctk.CTkButton(f_top, text="🔄 Actualizar", fg_color="#3B82F6", command=self.actualizar_graficos).pack(side="right", padx=10)
        self.al_cambiar_grado(grados[0])

    def al_cambiar_grado(self, grado):
        if grado == "Sin datos": return
        materias = self.engine.obtener_materias_por_grado(grado)
        if materias:
            self.combo_materia.configure(values=materias)
            self.combo_materia.set(materias[0])
        else:
            self.combo_materia.configure(values=["No hay materias"])
            self.combo_materia.set("No hay materias")

        estudiantes = self.engine.obtener_estudiantes_completos(grado)
        nombres = ["Todos los Estudiantes"] + [e['nombre'] for e in estudiantes]
        self.combo_estudiante.configure(values=nombres)
        self.combo_estudiante.set("Todos los Estudiantes")

        self.actualizar_graficos()

    def limpiar_graficos(self):
        for w in self.scroll_canvas.winfo_children():
            w.destroy()
        for fig in self.fig_objs:
            plt.close(fig)
        self.fig_objs.clear()

    def actualizar_graficos(self):
        self.limpiar_graficos()

        grado = self.combo_grado.get()
        materia = self.combo_materia.get()
        trimestre = self.combo_trimestre.get()
        estudiante_sel = self.combo_estudiante.get()

        if grado == "Sin datos" or materia == "Sin materias" or materia == "No hay materias":
            ctk.CTkLabel(self.scroll_canvas, text="No hay datos suficientes para graficar.",
                        font=("Segoe UI", 16)).grid(row=0, column=0, columnspan=2, pady=50)
            return

        estudiantes = []
        try:
            estudiantes = self.engine.obtener_estudiantes_completos(grado)
        except Exception: pass

        aprobados = 0; reprobados = 0; en_riesgo = 0

        promedios_por_est = getattr(self.engine, 'obtener_promedios_reales', lambda g,m,t: {})(grado, materia, trimestre)

        if not promedios_por_est and estudiantes:
            # Fallback a 1.0 si no hay calificaciones válidas registradas aún
            for est in estudiantes:
                promedios_por_est[est['nombre']] = 1.0

        for nom, promedio in promedios_por_est.items():
            if promedio >= 3.0: aprobados += 1
            elif promedio >= 2.5: en_riesgo += 1
            else: reprobados += 1

        if estudiante_sel != "Todos los Estudiantes":
            # Vista individual
            historial = getattr(self.engine, 'obtener_historial_real', lambda g,m,e: [3.0, 3.0])(grado, materia, estudiante_sel)
            if len(historial) < 2: historial = [3.0, 3.0] # Fallback minimo de regresion
            self.dibujar_proyeccion(estudiante_sel, historial, 0, 0, colspan=2)
        else:
            # Vista general
            self.dibujar_pastel(aprobados, en_riesgo, reprobados, 0, 0)
            self.dibujar_barras(promedios_por_est, 0, 1)
            self.dibujar_tendencia(list(promedios_por_est.values()), 1, 0, colspan=2)

    def dibujar_proyeccion(self, nombre, historial, row, col, colspan=1):
        import numpy as np
        from scipy.stats import linregress

        f_grafico = ctk.CTkFrame(self.scroll_canvas, fg_color=self.C["card"], corner_radius=10)
        f_grafico.grid(row=row, column=col, columnspan=colspan, sticky="nsew", padx=10, pady=10)

        fig, ax = plt.subplots(figsize=(10, 4), dpi=100)
        fig.patch.set_facecolor(self.C["card"])
        ax.set_facecolor(self.C["fondo"])
        self.fig_objs.append(fig)

        x = np.arange(1, len(historial) + 1)
        y = np.array(historial)

        # Regresion lineal
        slope, intercept, r_value, p_value, std_err = linregress(x, y)
        x_pred = np.arange(1, len(historial) + 3) # Proyectar 2 periodos mas
        y_pred = intercept + slope * x_pred

        ax.plot(x, y, marker='o', color=self.C["cian"], label='Historial de Notas', linewidth=2, markersize=8)
        ax.plot(x_pred, y_pred, linestyle='--', color=self.C["amarillo"], label='Tendencia / Predicción', linewidth=2)

        ax.axhline(y=3.0, color=self.C["rojo"], linestyle=':', alpha=0.7, label='Límite Aprobación (3.0)')

        ax.set_ylim(1.0, 5.2)
        ax.set_xticks(x_pred)
        ax.set_xticklabels([f"T{i}" for i in x] + ["Pred 1", "Pred 2"])

        ax.tick_params(axis='x', colors=self.C["texto"])
        ax.tick_params(axis='y', colors=self.C["texto"])

        ax.spines['bottom'].set_color(self.C["texto"])
        ax.spines['left'].set_color(self.C["texto"])
        ax.spines['top'].set_visible(False)
        ax.spines['right'].set_visible(False)

        ax.set_title(f"Proyección y Predicción de Rendimiento: {nombre}", color=self.C["cian"], pad=15, fontweight="bold")
        ax.legend(facecolor=self.C["fondo"], edgecolor=self.C["card"], labelcolor="white")
        fig.tight_layout()

        canvas = FigureCanvasTkAgg(fig, master=f_grafico)
        canvas.draw()
        canvas.get_tk_widget().pack(fill="both", expand=True, padx=10, pady=10)
    def dibujar_pastel(self, aprobados, riesgo, reprobados, row, col):
        f_grafico = ctk.CTkFrame(self.scroll_canvas, fg_color=self.C["card"], corner_radius=10)
        f_grafico.grid(row=row, column=col, sticky="nsew", padx=10, pady=10)

        fig, ax = plt.subplots(figsize=(5, 4), dpi=100)
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

    def dibujar_barras(self, promedios_dict, row, col):
        f_grafico = ctk.CTkFrame(self.scroll_canvas, fg_color=self.C["card"], corner_radius=10)
        f_grafico.grid(row=row, column=col, sticky="nsew", padx=10, pady=10)

        fig, ax = plt.subplots(figsize=(6, 4), dpi=100)
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

    def dibujar_tendencia(self, notas, row, col, colspan=1):
        f_grafico = ctk.CTkFrame(self.scroll_canvas, fg_color=self.C["card"], corner_radius=10)
        f_grafico.grid(row=row, column=col, columnspan=colspan, sticky="nsew", padx=10, pady=10)

        fig, ax = plt.subplots(figsize=(10, 3.5), dpi=100)
        fig.patch.set_facecolor(self.C["card"])
        ax.set_facecolor(self.C["fondo"])
        self.fig_objs.append(fig)

        if not notas:
            ax.text(0.5, 0.5, "Sin Datos", ha='center', va='center', color='white')
        else:
            # Histograma de curva / frecuencia
            counts, bins, patches = ax.hist(notas, bins=10, range=(1.0, 5.0), color=self.C["cian"], alpha=0.6, edgecolor='white')

            # Promedio general del salón
            promedio_gen = sum(notas) / len(notas)
            ax.axvline(x=promedio_gen, color=self.C["amarillo"], linestyle='-', linewidth=2, label=f'Promedio General ({promedio_gen:.1f})')
            ax.legend(facecolor=self.C["fondo"], edgecolor=self.C["card"], labelcolor="white")

        ax.tick_params(axis='x', colors=self.C["texto"])
        ax.tick_params(axis='y', colors=self.C["texto"])
        ax.set_xticks([1.0, 2.0, 3.0, 4.0, 5.0])
        ax.set_xlabel("Calificación", color=self.C["texto_dim"] if "texto_dim" in self.C else "#64748B")
        ax.set_ylabel("Cant. Estudiantes", color=self.C["texto_dim"] if "texto_dim" in self.C else "#64748B")

        ax.spines['bottom'].set_color(self.C["texto"])
        ax.spines['left'].set_color(self.C["texto"])
        ax.spines['top'].set_visible(False)
        ax.spines['right'].set_visible(False)

        ax.set_title("Distribución y Campana de Calificaciones", color=self.C["cian"], pad=15, fontweight="bold")
        fig.tight_layout()

        canvas = FigureCanvasTkAgg(fig, master=f_grafico)
        canvas.draw()
        canvas.get_tk_widget().pack(fill="both", expand=True, padx=10, pady=10)
