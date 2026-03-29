import os
import json
import customtkinter as ctk
from dashapp import DashboardFrame
from rddata import DataEngine
from setup import SetupWizard

# Importar paneles
from dapp import EstudiantesFrame
from eapp import NotasFrame
from fapp import AsistenciaFrame
from obsapp import ObservacionesFrame
from sapp import ConfigFrame

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

CONFIG_FILE = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "perfil.json"))


class RegistroDocApp(ctk.CTk):
    def __init__(self, modalidad_inicial="premedia"):
        super().__init__()
        self.title("RegistroDoc Pro v3.0 - MEDUCA Panamá")
        self.geometry("1280x720")

        self.sidebar_expanded = True

        # Motor
        archivo = "Registro_Primaria.xlsx" if modalidad_inicial == "primaria" else "Registro_2026.xlsx"
        base_dir = os.path.dirname(os.path.abspath(__file__))
        ruta_excel = os.path.abspath(os.path.join(base_dir, "..", archivo))
        self.engine = DataEngine(ruta_excel, modalidad_inicial)

        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        self.crear_menu_lateral()
        self.contenedor_principal = ctk.CTkFrame(self, fg_color="transparent")
        self.contenedor_principal.grid(row=0, column=1, padx=20, pady=20, sticky="nsew")

        self.mostrar_dashboard()

    # ====================== MENÚ LATERAL COMPACTO ESTILO GROK ======================
        # ====================== MENÚ LATERAL COMPACTO ESTILO GROK ======================
    def crear_menu_lateral(self):
        if hasattr(self, 'sidebar'):
            self.sidebar.destroy()

        width = 240 if self.sidebar_expanded else 68
        self.sidebar = ctk.CTkFrame(self, width=width, corner_radius=0, fg_color="#1A2638")
        self.sidebar.grid(row=0, column=0, sticky="nsew")
        self.sidebar.grid_rowconfigure(8, weight=1)

        # ==================== LOGO (carga real de tu archivo) ====================
        logo_path = os.path.join(os.path.dirname(__file__), "..", "img", "logo.png")
        
        if os.path.exists(logo_path):
            try:
                from PIL import Image
                pil_image = Image.open(logo_path)
                logo_img = ctk.CTkImage(light_image=pil_image, dark_image=pil_image, size=(32, 32))
                ctk.CTkLabel(self.sidebar, image=logo_img, text="").pack(pady=(25, 15))
            except:
                # fallback si hay algún problema con la imagen
                logo_text = "RD" if not self.sidebar_expanded else "RegistroDoc Pro"
                ctk.CTkLabel(self.sidebar, text=logo_text, font=ctk.CTkFont(size=20, weight="bold"), text_color="#3B82F6").pack(pady=(30, 20))
        else:
            logo_text = "RD" if not self.sidebar_expanded else "RegistroDoc Pro"
            ctk.CTkLabel(self.sidebar, text=logo_text, font=ctk.CTkFont(size=20, weight="bold"), text_color="#3B82F6").pack(pady=(30, 20))

        # ==================== ITEMS DEL MENÚ ====================
        items = [
            ("🏠", "Dashboard", self.mostrar_dashboard),
            ("👥", "Estudiantes", self.mostrar_estudiantes),
            ("📝", "Registro Notas", self.mostrar_notas),
            ("📅", "Asistencia", self.mostrar_asistencia),
            ("🔍", "Observaciones", self.mostrar_observaciones),
            ("⚙️", "Configuración", self.mostrar_configuracion),
        ]

        for icon, text, command in items:
            btn = ctk.CTkButton(self.sidebar, 
                                text=icon if not self.sidebar_expanded else f"  {icon}  {text}",
                                fg_color="transparent", 
                                hover_color="#2A3A50",
                                font=ctk.CTkFont(size=18 if not self.sidebar_expanded else 15),
                                anchor="w", 
                                height=42, 
                                command=command)
            btn.pack(pady=2, padx=8, fill="x")

        # ==================== BOTÓN SALIR PEQUEÑO ====================
        ctk.CTkButton(self.sidebar, text="🚪 Salir", fg_color="#EF4444", hover_color="#B91C1C",
                      font=ctk.CTkFont(size=13, weight="bold"), height=38, command=self.quit).pack(pady=20, padx=12, side="bottom")

        # ==================== TOGGLE << / >> ====================
        toggle_text = ">>" if self.sidebar_expanded else "<<"
        self.toggle_btn = ctk.CTkButton(self.sidebar, text=toggle_text, width=50, height=32,
                                        fg_color="#2A3A50", hover_color="#334155",
                                        font=ctk.CTkFont(size=18, weight="bold"),
                                        command=self.toggle_sidebar)
        self.toggle_btn.pack(pady=15, padx=10, side="bottom")

    # ====================== DASHBOARD Y DEMÁS ======================
    def limpiar_pantalla(self):
        for widget in self.contenedor_principal.winfo_children():
            widget.destroy()

    def mostrar_dashboard(self):
        self.limpiar_pantalla()
        panel = DashboardFrame(self.contenedor_principal, self.engine, app_principal=self)
        panel.pack(fill="both", expand=True)

    def mostrar_estudiantes(self):
        self.limpiar_pantalla()
        panel = EstudiantesFrame(self.contenedor_principal, self.engine)
        panel.pack(fill="both", expand=True)

    def mostrar_notas(self):
        self.limpiar_pantalla()
        panel = NotasFrame(self.contenedor_principal, self.engine)
        panel.pack(fill="both", expand=True)

    def mostrar_asistencia(self):
        self.limpiar_pantalla()
        panel = AsistenciaFrame(self.contenedor_principal, self.engine)
        panel.pack(fill="both", expand=True)

    def mostrar_observaciones(self):
        self.limpiar_pantalla()
        panel = ObservacionesFrame(self.contenedor_principal, self.engine)
        panel.pack(fill="both", expand=True)

    def mostrar_configuracion(self):
        self.limpiar_pantalla()
        panel = ConfigFrame(self.contenedor_principal, self.engine, self)
        panel.pack(fill="both", expand=True)

    def reiniciar_motor(self, nueva_ruta, nueva_modalidad):
        self.engine = DataEngine(nueva_ruta, nueva_modalidad)
        with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
            data = json.load(f)
        data["modalidad"] = nueva_modalidad
        with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=4)
        self.crear_menu_lateral()
        self.mostrar_dashboard()


# =======================================================
def iniciar_programa_principal():
    with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
        config = json.load(f)
    app = RegistroDocApp(modalidad_inicial=config.get("modalidad", "premedia"))
    app.mainloop()

if __name__ == "__main__":
    if not os.path.exists(CONFIG_FILE):
        wizard = SetupWizard()
        wizard.mainloop()
        if os.path.exists(CONFIG_FILE):
            iniciar_programa_principal()
    else:
        iniciar_programa_principal()