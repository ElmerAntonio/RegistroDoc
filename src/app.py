import os
import customtkinter as ctk
from fapp import AsistenciaFrame
from rddata import DataEngine
from obsapp import ObservacionesFrame
from dapp import EstudiantesFrame  # <--- NUEVO IMPORT
from eapp import NotasFrame

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

class RegistroDocApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("RegistroDoc Pro v3.0 - MEDUCA Panamá")
        self.geometry("1100x700")

        base_dir = os.path.dirname(os.path.abspath(__file__))
        ruta_excel = os.path.abspath(os.path.join(base_dir, "..", "Registro_2026.xlsx"))
        self.engine = DataEngine(ruta_excel, "premedia")

        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        self.crear_menu_lateral()
        self.contenedor_principal = ctk.CTkFrame(self, fg_color="transparent")
        self.contenedor_principal.grid(row=0, column=1, padx=20, pady=20, sticky="nsew")

        self.mostrar_dashboard()

    def crear_menu_lateral(self):
        self.sidebar = ctk.CTkFrame(self, width=220, corner_radius=0, fg_color="#1A2638")
        self.sidebar.grid(row=0, column=0, sticky="nsew")
        self.sidebar.grid_rowconfigure(6, weight=1) 

        ctk.CTkLabel(self.sidebar, text="RegistroDoc Pro", font=("Segoe UI", 22, "bold"), text_color="#3B82F6").pack(pady=(30, 30))

        # --- BOTONES ACTUALIZADOS ---
        ctk.CTkButton(self.sidebar, text="🏠  Dashboard", fg_color="transparent", font=("Segoe UI", 15), anchor="w", command=self.mostrar_dashboard).pack(pady=5, padx=20, fill="x")
        ctk.CTkButton(self.sidebar, text="👥  Estudiantes", fg_color="transparent", font=("Segoe UI", 15), anchor="w", command=self.mostrar_estudiantes).pack(pady=5, padx=20, fill="x")
        ctk.CTkButton(self.sidebar, text="📝  Registro Notas", fg_color="transparent", font=("Segoe UI", 15), anchor="w", command=self.mostrar_notas).pack(pady=5, padx=20, fill="x")
        ctk.CTkButton(self.sidebar, text="📅  Asistencia", fg_color="transparent", font=("Segoe UI", 15), anchor="w", command=self.mostrar_asistencia).pack(pady=5, padx=20, fill="x")
        # -- AÑADIR ESTE BOTÓN NUEVO DEBAJO DEL DE ASISTENCIA --
        ctk.CTkButton(self.sidebar, text="🔍  Observaciones", fg_color="transparent", font=("Segoe UI", 15), anchor="w", command=self.mostrar_observaciones).pack(pady=5, padx=20, fill="x")

        ctk.CTkButton(self.sidebar, text="🔒 Bloquear / Salir", fg_color="#EF4444", hover_color="#B91C1C", font=("Segoe UI", 13, "bold"), command=self.quit).pack(pady=20, padx=20, side="bottom")

    def limpiar_pantalla(self):
        for widget in self.contenedor_principal.winfo_children():
            widget.destroy()

    def mostrar_dashboard(self):
        self.limpiar_pantalla()
        ctk.CTkLabel(self.contenedor_principal, text="Resumen Ejecutivo Académico", font=("Segoe UI", 24, "bold")).pack(anchor="w", pady=(0, 20))
        stats = self.engine.get_dashboard_stats()
        tarjetas = ctk.CTkFrame(self.contenedor_principal, fg_color="transparent")
        tarjetas.pack(fill="x")
        self.crear_tarjeta(tarjetas, "Total Matriculados", str(stats["total"]), "#3B82F6")
        self.crear_tarjeta(tarjetas, "⚠️ En Riesgo Académico", str(stats["riesgo"]), "#EF4444")
        self.crear_tarjeta(tarjetas, "⭐ Cuadro de Honor", stats["honor"], "#10B981")

    def crear_tarjeta(self, parent, titulo, valor, color_borde):
        card = ctk.CTkFrame(parent, fg_color="#1E2D42", border_width=2, border_color=color_borde, corner_radius=8)
        card.pack(side="left", fill="both", expand=True, padx=10)
        ctk.CTkLabel(card, text=titulo, font=("Segoe UI", 14), text_color="#94A3B8").pack(pady=(15, 5))
        ctk.CTkLabel(card, text=valor, font=("Segoe UI", 22, "bold"), text_color="white").pack(pady=(0, 15))

    # --- NUEVA FUNCIÓN PARA MOSTRAR ESTUDIANTES ---
    def mostrar_estudiantes(self):
        self.limpiar_pantalla()
        panel_est = EstudiantesFrame(self.contenedor_principal, self.engine)
        panel_est.pack(fill="both", expand=True)

    def mostrar_notas(self):
        self.limpiar_pantalla()
        panel_not = NotasFrame(self.contenedor_principal, self.engine)
        panel_not.pack(fill="both", expand=True)

    def mostrar_asistencia(self):
        self.limpiar_pantalla()
        panel_asist = AsistenciaFrame(self.contenedor_principal, self.engine)
        panel_asist.pack(fill="both", expand=True)

    def mostrar_observaciones(self):
        self.limpiar_pantalla()
        panel_obs = ObservacionesFrame(self.contenedor_principal, self.engine)
        panel_obs.pack(fill="both", expand=True)
        
    def en_construccion(self):
        self.limpiar_pantalla()
        ctk.CTkLabel(self.contenedor_principal, text="Módulo en Construcción 🚧", font=("Segoe UI", 24, "bold"), text_color="#94A3B8").pack(expand=True)

if __name__ == "__main__":
    app = RegistroDocApp()
    app.mainloop()
