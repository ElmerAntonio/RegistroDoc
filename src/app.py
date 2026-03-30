import os
import json
import customtkinter as ctk
from dashapp import DashboardFrame
from rddata import DataEngine

try:
    from setup import SetupWizard
except ImportError:
    SetupWizard = None

from dapp   import EstudiantesFrame
from eapp   import NotasFrame
from fapp   import AsistenciaFrame
from obsapp import ObservacionesFrame
from sapp   import ConfigFrame

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

CONFIG_FILE = os.path.abspath(
    os.path.join(os.path.dirname(__file__), "..", "perfil.json"))

ANCHO_FIJO = 1280
ALTO_FIJO  = 720


class RegistroDocApp(ctk.CTk):
    def __init__(self, modalidad_inicial="premedia"):
        super().__init__()
        self.title("RegistroDoc Pro v3.0 — MEDUCA Panamá")

        # Ventana fija — no se puede redimensionar
        self.geometry(f"{ANCHO_FIJO}x{ALTO_FIJO}")
        self.minsize(ANCHO_FIJO, ALTO_FIJO)
        self.maxsize(ANCHO_FIJO, ALTO_FIJO)
        self.resizable(False, False)

        # Icono de la barra de tareas actualizado con tus archivos reales
        for nombre_icono in [
            "icon.ico",
            "icono.jpg",
            "logo.png",
            "Gemini_Generated_Image_h7eh8rh7eh8rh7eh.png"
        ]:
            icon_path = os.path.join(
                os.path.dirname(__file__), "..", "img", nombre_icono)
            if os.path.exists(icon_path):
                try:
                    from PIL import Image, ImageTk
                    pil = Image.open(icon_path).resize((64, 64))
                    self._icono_app = ImageTk.PhotoImage(pil)
                    self.iconphoto(True, self._icono_app)
                    break
                except Exception:
                    pass

        # Motor de datos
        archivo = ("Registro_Primaria.xlsx"
                   if modalidad_inicial == "primaria"
                   else "Registro_2026.xlsx")
        base    = os.path.dirname(os.path.abspath(__file__))
        ruta    = os.path.abspath(os.path.join(base, "..", archivo))
        self.engine = DataEngine(ruta, modalidad_inicial)

        # Contenedor raíz
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)
        self.contenedor = ctk.CTkFrame(self, fg_color="#0A1628",
                                       corner_radius=0)
        self.contenedor.grid(row=0, column=0, sticky="nsew")

        self.mostrar_dashboard()

    def limpiar_pantalla(self):
        for w in self.contenedor.winfo_children():
            w.destroy()

    def mostrar_dashboard(self):
        self.limpiar_pantalla()
        DashboardFrame(self.contenedor, self.engine,
                       app_principal=self).pack(fill="both", expand=True)

    def mostrar_estudiantes(self):
        self.limpiar_pantalla()
        EstudiantesFrame(self.contenedor,
                         self.engine).pack(fill="both", expand=True)

    def mostrar_notas(self):
        self.limpiar_pantalla()
        NotasFrame(self.contenedor,
                   self.engine).pack(fill="both", expand=True)

    def mostrar_asistencia(self):
        self.limpiar_pantalla()
        AsistenciaFrame(self.contenedor,
                        self.engine).pack(fill="both", expand=True)

    def mostrar_observaciones(self):
        self.limpiar_pantalla()
        ObservacionesFrame(self.contenedor,
                           self.engine).pack(fill="both", expand=True)

    def mostrar_configuracion(self):
        self.limpiar_pantalla()
        ConfigFrame(self.contenedor,
                    self.engine, self).pack(fill="both", expand=True)

    def reiniciar_motor(self, nueva_ruta, nueva_modalidad):
        self.engine = DataEngine(nueva_ruta, nueva_modalidad)
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
            data["modalidad"] = nueva_modalidad
            with open(CONFIG_FILE, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=4)
        except Exception:
            pass
        self.mostrar_dashboard()


def iniciar_programa_principal():
    try:
        with open(CONFIG_FILE, "r", encoding="utf-8") as f:
            config = json.load(f)
    except (FileNotFoundError, json.JSONDecodeError):
        config = {"modalidad": "premedia"}

    app = RegistroDocApp(
        modalidad_inicial=config.get("modalidad", "premedia"))
    app.mainloop()


if __name__ == "__main__":
    if not os.path.exists(CONFIG_FILE):
        if SetupWizard:
            wizard = SetupWizard()
            wizard.mainloop()
            if os.path.exists(CONFIG_FILE):
                iniciar_programa_principal()
        else:
            cfg = {"modalidad": "premedia",
                   "docente_nombre": "", "ano_lectivo": "2026"}
            os.makedirs(os.path.dirname(CONFIG_FILE), exist_ok=True)
            with open(CONFIG_FILE, "w", encoding="utf-8") as f:
                json.dump(cfg, f, ensure_ascii=False, indent=4)
            iniciar_programa_principal()
    else:
        iniciar_programa_principal()