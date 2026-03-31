import os
import sys
import json
import tkinter as tk
import customtkinter as ctk

# Insert local directory into sys.path to allow local module imports
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

# Matplotlib
import matplotlib
matplotlib.use("TkAgg")

try:
    from PIL import Image, ImageTk
    PIL_OK = True
except ImportError:
    PIL_OK = False

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

# ─── PALETA PRINCIPAL ──────────────────────────────────────────────
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

class MainApplication(ctk.CTkFrame):
    def __init__(self, master, engine, app_principal, **kwargs):
        super().__init__(master, fg_color=C["fondo"], corner_radius=0, **kwargs)
        self.engine = engine
        self.app = app_principal
        self._acento = C["cian"]

        self.pack_propagate(False)
        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)   # El cuerpo ahora toma toda la pantalla

        self._cuerpo()

    # ══════════════════════════════════════════════════════════════════════
    #  CUERPO: SIDEBAR + PANEL PRINCIPAL
    # ══════════════════════════════════════════════════════════════════════
    def _cuerpo(self):
        cuerpo = ctk.CTkFrame(self, fg_color=C["fondo"], corner_radius=0)
        cuerpo.grid(row=0, column=0, sticky="nsew") # Movido a la fila 0
        cuerpo.rowconfigure(0, weight=1)
        cuerpo.columnconfigure(1, weight=1)

        self._sidebar_widget(cuerpo)

        self.main_content_frame = ctk.CTkFrame(cuerpo, fg_color=C["fondo"], corner_radius=0)
        self.main_content_frame.grid(row=0, column=1, sticky="nsew", padx=0, pady=0)
        self.main_content_frame.columnconfigure(0, weight=1)
        self.main_content_frame.rowconfigure(0, weight=1)

    # ══════════════════════════════════════════════════════════════════════
    #  SIDEBAR (FIJO Y SIN RETRAER)
    # ══════════════════════════════════════════════════════════════════════
    def _sidebar_widget(self, parent):
        self._sb = ctk.CTkFrame(parent, fg_color=C["sidebar"],
                                width=230, corner_radius=0) # Ancho fijo
        self._sb.grid(row=0, column=0, sticky="nsew")
        self._sb.grid_propagate(False)
        self._sb.rowconfigure(8, weight=1)

        self._sb_renderizar()

    def _sb_renderizar(self):
        for w in self._sb.winfo_children():
            w.destroy()

        # Cargar Logo en Sidebar (Grande y centrado)
        logo_path = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "img", "icono.png"))

        if PIL_OK and os.path.exists(logo_path):
            try:
                size = (140, 140)
                pil_logo = Image.open(logo_path).resize(size, Image.LANCZOS)
                self._sb_logo_img = ctk.CTkImage(pil_logo, size=size)
                ctk.CTkLabel(self._sb, image=self._sb_logo_img, text="").pack(pady=(30, 15))
            except Exception:
                pass
        else:
            ctk.CTkLabel(self._sb, text="📘 RegistroDoc", fg_color="transparent",
                         font=ctk.CTkFont("Segoe UI", 20, "bold"),
                         text_color=self._acento).pack(pady=(30, 15))

        ctk.CTkFrame(self._sb, fg_color=C["borde"], height=1).pack(
            fill="x", padx=15, pady=10)

        # Items del menú (Agregamos Configuración al final)
        items = [
            ("🏠", "Inicio",        self._ir_inicio,       True),
            ("👤", "Estudiantes",   self._ir_estudiantes,  False),
            ("📝", "Notas",         self._ir_notas,        False),
            ("📅", "Asistencia",    self._ir_asistencia,   False),
            ("📋", "Reportes",      self._ir_reportes,     False),
            ("⚙️", "Configuración", self._ir_configuracion, False)
        ]

        for icono, texto, cmd, activo in items:
            bg = C["activo"] if activo else "transparent"
            tc = self._acento if activo else C["texto_sec"]
            bw = 2 if activo else 0

            label_txt = f"  {icono}   {texto}"

            btn = ctk.CTkButton(
                self._sb, text=label_txt,
                fg_color=bg,
                hover_color=C["hover"],
                font=ctk.CTkFont("Segoe UI", 15),
                text_color=tc, anchor="w",
                height=45, corner_radius=6,
                border_width=bw,
                border_color=self._acento,
                command=cmd)
            
            btn.pack(fill="x", padx=12, pady=4)

        # Espacio flexible inferior
        ctk.CTkFrame(self._sb, fg_color="transparent").pack(
            fill="both", expand=True)

        ctk.CTkLabel(self._sb, text="v.Prov.22:6",
                     font=ctk.CTkFont("Segoe UI", 11),
                     text_color=C["texto_dim"]).pack(pady=(0, 15))

    # Rutas de navegación
    def _ir_inicio(self):
        if self.app:
            try: self.app.mostrar_dashboard()
            except Exception: pass

    def _ir_estudiantes(self):
        if self.app:
            try: self.app.mostrar_estudiantes()
            except Exception: pass

    def _ir_notas(self):
        if self.app:
            try: self.app.mostrar_notas()
            except Exception: pass

    def _ir_asistencia(self):
        if self.app:
            try: self.app.mostrar_asistencia()
            except Exception: pass

    def _ir_reportes(self):
        pass

    def _ir_configuracion(self):
        if self.app:
            try: self.app.mostrar_configuracion()
            except Exception: pass


class RegistroDocApp(ctk.CTk):
    def __init__(self, modalidad_inicial="premedia"):
        super().__init__()
        self.title("RegistroDoc Pro v3.0 — MEDUCA Panamá")

        # ─── VENTANA REDIMENSIONABLE Y MAXIMIZABLE ───
        self.geometry("1280x720")
        self.minsize(1024, 600)  # Tamaño mínimo para que no se deforme
        self.resizable(True, True) # Permite maximizar y achicar

        # Intentar iniciar maximizado en Windows
        try:
            self.state("zoomed")
        except Exception:
            pass

        # Iconos de la ventana
        icon_path = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "img", "icon.ico"))
        if os.path.exists(icon_path):
            try:
                self.iconbitmap(icon_path)
            except Exception:
                pass

        for nombre_icono in ["icon.ico", "icono.png"]:
            img_path = os.path.abspath(os.path.join(
                os.path.dirname(__file__), "..", "img", nombre_icono))
            if os.path.exists(img_path):
                try:
                    from PIL import Image, ImageTk
                    pil = Image.open(img_path).resize((64, 64))
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

        # Iniciar Aplicación
        self.main_app = MainApplication(self, self.engine, app_principal=self)
        self.main_app.grid(row=0, column=0, sticky="nsew")

        self.mostrar_dashboard()

    def limpiar_pantalla(self):
        for w in self.main_app.main_content_frame.winfo_children():
            w.destroy()

    def mostrar_dashboard(self):
        self.limpiar_pantalla()
        DashboardFrame(self.main_app.main_content_frame, self.engine,
                       app_principal=self).pack(fill="both", expand=True)

    def mostrar_estudiantes(self):
        self.limpiar_pantalla()
        EstudiantesFrame(self.main_app.main_content_frame,
                         self.engine).pack(fill="both", expand=True)

    def mostrar_notas(self):
        self.limpiar_pantalla()
        NotasFrame(self.main_app.main_content_frame,
                   self.engine).pack(fill="both", expand=True)

    def mostrar_asistencia(self):
        self.limpiar_pantalla()
        AsistenciaFrame(self.main_app.main_content_frame,
                        self.engine).pack(fill="both", expand=True)

    def mostrar_observaciones(self):
        self.limpiar_pantalla()
        ObservacionesFrame(self.main_app.main_content_frame,
                           self.engine).pack(fill="both", expand=True)

    def mostrar_configuracion(self):
        self.limpiar_pantalla()
        ConfigFrame(self.main_app.main_content_frame,
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
        self.main_app.engine = self.engine
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