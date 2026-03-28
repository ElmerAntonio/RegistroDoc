import customtkinter as ctk
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

class GraficosFrame(ctk.CTkFrame):
    def __init__(self, master, engine, **kwargs):
        super().__init__(master, **kwargs)
        self.engine = engine
        self.render_dashboard()

    def render_dashboard(self):
        # Título de Sección
        ctk.CTkLabel(self, text="Análisis de Rendimiento Académico", font=("Arial", 20, "bold")).pack(pady=10)
        
        # Frame para gráficos
        self.fig_container = ctk.CTkFrame(self, fg_color="transparent")
        self.fig_container.pack(fill="both", expand=True, padx=20)

        # Gráfico de Barras: Estudiantes en Riesgo (Notas < 3.0)
        fig, ax = plt.subplots(figsize=(5, 4), dpi=100)
        fig.patch.set_facecolor('#1a1a1a')
        ax.set_facecolor('#1a1a1a')
        
        # Simulación de datos (Claude debe conectar esto a engine.get_promedios())
        categorias = ['Aprobados', 'En Riesgo', 'Fracasos']
        valores = [25, 8, 3] 
        colores = ['#2ecc71', '#f1c40f', '#e74c3c']

        ax.bar(categorias, valores, color=colores)
        ax.tick_params(axis='x', colors='white')
        ax.tick_params(axis='y', colors='white')
        ax.set_title("Estado General del Grupo", color='white')

        canvas = FigureCanvasTkAgg(fig, master=self.fig_container)
        canvas.draw()
        canvas.get_tk_widget().pack(side="left", fill="both", expand=True)