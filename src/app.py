import os
import json
import tkinter as tk
import customtkinter as ctk

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

ANCHO_FIJO = 1280
ALTO_FIJO  = 720

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

class MainApplication(ctk.CTkFrame):
    def __init__(self, master, engine, app_principal, **kwargs):
        super().__init__(master, fg_color=C["fondo"], corner_radius=0, **kwargs)
        self.engine = engine
        self.app = app_principal
        self._acento = C["cian"]
        self._perfil_menu = None

        self.pack_propagate(False)
        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=0)   # header
        self.rowconfigure(1, weight=0)   # separator
        self.rowconfigure(2, weight=1)   # cuerpo

        self._header()

        ctk.CTkFrame(self, height=1, fg_color=C["borde"]).grid(
            row=1, column=0, sticky="ew")

        self._cuerpo()

    # ══════════════════════════════════════════════════════════════════════
    #  HEADER BAR
    # ══════════════════════════════════════════════════════════════════════
    def _header(self):
        hdr = ctk.CTkFrame(self, fg_color=C["header"], height=56, corner_radius=0)
        hdr.grid(row=0, column=0, sticky="ew")
        hdr.grid_propagate(False)
        hdr.columnconfigure(1, weight=1)

        # Izquierda: Logo texto
        left = ctk.CTkFrame(hdr, fg_color="transparent")
        left.grid(row=0, column=0, sticky="w", padx=(16, 0))

        # Intentar cargar logo pequeño
        logo_path = os.path.join(os.path.dirname(__file__), "..", "img",
                                 "icono.jpg")
        if not os.path.exists(logo_path):
            logo_path = os.path.join(os.path.dirname(__file__), "..", "img", "logo.png")

        if PIL_OK and os.path.exists(logo_path):
            try:
                pil = Image.open(logo_path).resize((36, 36), Image.LANCZOS)
                self._logo_img = ctk.CTkImage(pil, size=(36, 36))
                ctk.CTkLabel(left, image=self._logo_img, text="").pack(
                    side="left", padx=(0, 8))
            except Exception:
                pass

        titulo = ctk.CTkLabel(left, text="RegistroDoc ", fg_color="transparent",
                              font=ctk.CTkFont("Segoe UI", 18, "bold"),
                              text_color=C["texto"])
        titulo.pack(side="left")
        ctk.CTkLabel(left, text="Pro", fg_color="transparent",
                     font=ctk.CTkFont("Segoe UI", 18, "bold"),
                     text_color=self._acento).pack(side="left")

        # Centro: barra de búsqueda
        centro = ctk.CTkFrame(hdr, fg_color="transparent")
        centro.grid(row=0, column=1, sticky="ew", padx=40)
        centro.columnconfigure(0, weight=1)

        self._search_frame = ctk.CTkFrame(
            centro, fg_color=C["input"],
            border_width=1, border_color=C["borde"],
            corner_radius=20, height=34)
        self._search_frame.grid(row=0, column=0, sticky="ew")
        self._search_frame.grid_propagate(False)
        self._search_frame.columnconfigure(1, weight=1)

        ctk.CTkLabel(self._search_frame, text="🔍", fg_color="transparent",
                     font=ctk.CTkFont(size=14), text_color=C["texto_sec"],
                     width=30).grid(row=0, column=0, padx=(10, 0))

        self._search_var = tk.StringVar()
        self._search_var.trace_add("write", self._buscar_estudiante)
        self._search_entry = ctk.CTkEntry(
            self._search_frame, textvariable=self._search_var,
            fg_color="transparent", border_width=0,
            placeholder_text="Buscar alumno o documento...",
            font=ctk.CTkFont("Segoe UI", 13),
            text_color=C["texto"])
        self._search_entry.grid(row=0, column=1, sticky="ew",
                                padx=(4, 10), pady=4)

        # Panel de resultados (oculto por defecto)
        self._resultados_frame = ctk.CTkFrame(
            centro, fg_color=C["card"],
            border_width=1, border_color=self._acento,
            corner_radius=8)
        # no se muestra hasta que haya búsqueda

        # Derecha: iconos
        right = ctk.CTkFrame(hdr, fg_color="transparent")
        right.grid(row=0, column=2, sticky="e", padx=(0, 16))

        # Ícono notificaciones con badge
        notif_f = ctk.CTkFrame(right, fg_color="transparent")
        notif_f.pack(side="left", padx=3)
        self._notif_btn = ctk.CTkButton(
            notif_f, text="🔔", width=36, height=36,
            fg_color="transparent", hover_color=C["hover"],
            font=ctk.CTkFont(size=16), text_color=C["texto_sec"],
            corner_radius=18, command=self._toggle_notificaciones)
        self._notif_btn.pack()
        # badge
        badge = ctk.CTkLabel(notif_f, text="5", width=16, height=16,
                              fg_color="#EF4444", corner_radius=8,
                              font=ctk.CTkFont(size=9, weight="bold"),
                              text_color="white")
        badge.place(relx=0.6, rely=0.0)

        # Avatar circular → menú desplegable
        # Attempt to load avatar 1, 2, or 3
        av_path = None
        for i in ["", "1", "2", "3"]:
            p = os.path.join(os.path.dirname(__file__), "..", "img", f"avatar{i}.png")
            if os.path.exists(p):
                av_path = p
                break

        self._avatar_btn = ctk.CTkButton(
            right, text="ET", width=36, height=36,
            fg_color=self._acento, hover_color=C["hover"],
            font=ctk.CTkFont("Segoe UI", 13, "bold"),
            text_color=C["fondo"], corner_radius=18,
            command=self._toggle_perfil)
        if PIL_OK and av_path:
            try:
                pil_av = Image.open(av_path).resize((36, 36))
                self._av_img = ctk.CTkImage(pil_av, size=(36, 36))
                self._avatar_btn.configure(image=self._av_img, text="")
            except Exception:
                pass
        self._avatar_btn.pack(side="left", padx=(6, 0))

    def _buscar_estudiante(self, *_):
        # We only show search results if we are on Dashboard
        # It handles navigation or just displays info
        texto = self._search_var.get().strip()
        for w in self._resultados_frame.winfo_children():
            w.destroy()

        if len(texto) < 2:
            self._resultados_frame.grid_remove()
            return

        try:
            resultados = []
            for g in self.engine.obtener_grados_activos():
                for est in self.engine.obtener_estudiantes_completos(g):
                    if texto.lower() in est["nombre"].lower():
                        resultados.append((g, est["nombre"]))
        except Exception:
            resultados = [("7°A", "Maria Gonzalez"), ("8°", "Maria Gonzalez")]

        if not resultados:
            ctk.CTkLabel(self._resultados_frame,
                         text="Sin resultados", text_color=C["texto_sec"],
                         font=ctk.CTkFont(size=12)).pack(pady=8)
        else:
            for grado, nombre in resultados[:6]:
                btn = ctk.CTkButton(
                    self._resultados_frame,
                    text=f"  {nombre}   ({grado} - Estudiante)",
                    fg_color="transparent", hover_color=C["hover"],
                    anchor="w", font=ctk.CTkFont("Segoe UI", 12),
                    text_color=C["texto"], height=32,
                    command=lambda n=nombre: self._seleccionar_resultado(n))
                btn.pack(fill="x", padx=4, pady=1)

        self._resultados_frame.grid(row=1, column=0, sticky="ew",
                                    pady=(2, 0))

    def _seleccionar_resultado(self, nombre):
        self._search_var.set(nombre)
        self._resultados_frame.grid_remove()

    def _toggle_notificaciones(self):
        # Simple placeholder popup for notifications
        popup = ctk.CTkToplevel(self)
        popup.title("Notificaciones")
        popup.geometry("300x400")
        popup.resizable(False, False)
        popup.attributes("-topmost", True)
        ctk.CTkLabel(popup, text="Sin notificaciones nuevas", font=ctk.CTkFont("Segoe UI", 14)).pack(pady=20)

    # ══════════════════════════════════════════════════════════════════════
    #  MENÚ PERFIL DESPLEGABLE
    # ══════════════════════════════════════════════════════════════════════
    def _toggle_perfil(self):
        if self._perfil_menu and self._perfil_menu.winfo_exists():
            self._perfil_menu.destroy()
            self._perfil_menu = None
            return

        self._perfil_menu = ctk.CTkToplevel(self)
        self._perfil_menu.overrideredirect(True)
        self._perfil_menu.configure(fg_color=C["card"])
        self._perfil_menu.attributes("-topmost", True)

        x = self.winfo_rootx() + self.winfo_width() - 270
        y = self.winfo_rooty() + 60
        self._perfil_menu.geometry(f"250x320+{x}+{y}")

        ctk.CTkFrame(self._perfil_menu, fg_color=self._acento,
                     height=2).pack(fill="x")

        try:
            datos = self.engine.obtener_datos_generales()
            nombre_d = datos.get("docente_nombre", "Docente") or "Docente"
            correo_d = datos.get("correo", "") or "profe@meduca.edu.pa"
        except Exception:
            nombre_d = "Prof. Elmer Tugri"
            correo_d = "elmer.tugri7@meduca.edu.pa"

        info = ctk.CTkFrame(self._perfil_menu, fg_color="transparent")
        info.pack(fill="x", padx=16, pady=14)
        ctk.CTkLabel(info, text="👤", font=ctk.CTkFont(size=28)).pack(side="left")
        txt_f = ctk.CTkFrame(info, fg_color="transparent")
        txt_f.pack(side="left", padx=10)
        ctk.CTkLabel(txt_f, text=nombre_d[:22],
                     font=ctk.CTkFont("Segoe UI", 12, "bold"),
                     text_color=C["texto"]).pack(anchor="w")
        ctk.CTkLabel(txt_f, text=correo_d[:28],
                     font=ctk.CTkFont("Segoe UI", 10),
                     text_color=C["texto_sec"]).pack(anchor="w")

        ctk.CTkFrame(self._perfil_menu, fg_color=C["borde"],
                     height=1).pack(fill="x", padx=12)

        opciones = [
            ("👤  Editar Perfil",           lambda: None),
            ("⚙️  Configuración",           self._abrir_configuracion),
            ("🎓  Cambiar Año Lectivo",      lambda: None),
        ]
        for txt, cmd in opciones:
            ctk.CTkButton(self._perfil_menu, text=txt, fg_color="transparent",
                          hover_color=C["hover"], anchor="w",
                          font=ctk.CTkFont("Segoe UI", 13),
                          text_color=C["texto"], height=36,
                          command=lambda c=cmd: (self._perfil_menu.destroy(), c())
                          ).pack(fill="x", padx=8, pady=2)

        ctk.CTkFrame(self._perfil_menu, fg_color=C["borde"],
                     height=1).pack(fill="x", padx=12, pady=4)

        ctk.CTkLabel(self._perfil_menu, text="Color de acento:",
                     font=ctk.CTkFont(size=11), text_color=C["texto_sec"]).pack(
            anchor="w", padx=16)

        colores_f = ctk.CTkFrame(self._perfil_menu, fg_color="transparent")
        colores_f.pack(fill="x", padx=16, pady=6)
        colores = [
            ("Cian",       "#00DDEB"),
            ("Verde",      "#00FF88"),
            ("Rojo",       "#FF4444"),
            ("Púrpura",    "#A855F7"),
        ]
        for nombre_c, hex_c in colores:
            ctk.CTkButton(colores_f, text="", width=28, height=28,
                          fg_color=hex_c, hover_color=hex_c,
                          corner_radius=14,
                          command=lambda h=hex_c: self._cambiar_acento(h)
                          ).pack(side="left", padx=4)

        ctk.CTkFrame(self._perfil_menu, fg_color=C["borde"],
                     height=1).pack(fill="x", padx=12, pady=4)

        ctk.CTkButton(self._perfil_menu, text="🚪  Cerrar Sesión",
                      fg_color="transparent", hover_color="#7F1D1D",
                      anchor="w", font=ctk.CTkFont("Segoe UI", 13),
                      text_color="#EF4444", height=34,
                      command=lambda: self.app.quit()).pack(fill="x", padx=8)

        self._perfil_menu.bind("<FocusOut>", lambda e: self._cerrar_perfil_menu())

    def _cerrar_perfil_menu(self):
        try:
            if self._perfil_menu and self._perfil_menu.winfo_exists():
                self._perfil_menu.destroy()
                self._perfil_menu = None
        except Exception:
            pass

    def _cambiar_acento(self, hex_color):
        self._acento = hex_color
        self._cerrar_perfil_menu()

    def _abrir_configuracion(self):
        if self.app:
            try:
                self.app.mostrar_configuracion()
            except Exception:
                pass

    # ══════════════════════════════════════════════════════════════════════
    #  CUERPO: SIDEBAR + PANEL PRINCIPAL
    # ══════════════════════════════════════════════════════════════════════
    def _cuerpo(self):
        cuerpo = ctk.CTkFrame(self, fg_color=C["fondo"], corner_radius=0)
        cuerpo.grid(row=2, column=0, sticky="nsew")
        cuerpo.rowconfigure(0, weight=1)
        cuerpo.columnconfigure(1, weight=1)

        self._sidebar_widget(cuerpo)

        # main_content_frame holds the dynamic content
        self.main_content_frame = ctk.CTkFrame(cuerpo, fg_color=C["fondo"], corner_radius=0)
        self.main_content_frame.grid(row=0, column=1, sticky="nsew", padx=0, pady=0)
        self.main_content_frame.columnconfigure(0, weight=1)
        self.main_content_frame.rowconfigure(0, weight=1)

    # ══════════════════════════════════════════════════════════════════════
    #  SIDEBAR
    # ══════════════════════════════════════════════════════════════════════
    def _sidebar_widget(self, parent):
        self._sidebar_exp = True
        self._sb = ctk.CTkFrame(parent, fg_color=C["sidebar"],
                                width=200, corner_radius=0)
        self._sb.grid(row=0, column=0, sticky="nsew")
        self._sb.grid_propagate(False)
        self._sb.rowconfigure(8, weight=1)

        self._sb_renderizar()

    def _sb_renderizar(self):
        for w in self._sb.winfo_children():
            w.destroy()

        ancho = 200 if self._sidebar_exp else 45

        # Add Logo at top of Sidebar
        logo_path = os.path.join(os.path.dirname(__file__), "..", "img", "icon.ico")
        if not os.path.exists(logo_path):
            logo_path = os.path.join(os.path.dirname(__file__), "..", "img", "logo.png")

        if PIL_OK and os.path.exists(logo_path):
            try:
                size = (120, 120) if self._sidebar_exp else (32, 32)
                pil_logo = Image.open(logo_path).resize(size, Image.LANCZOS)
                self._sb_logo_img = ctk.CTkImage(pil_logo, size=size)
                ctk.CTkLabel(self._sb, image=self._sb_logo_img, text="").pack(pady=(12, 8))
            except Exception:
                pass
        else:
            if not self._sidebar_exp:
                ctk.CTkLabel(self._sb, text="≡", fg_color="transparent",
                             font=ctk.CTkFont(size=20),
                             text_color=C["texto_sec"]).pack(pady=(14, 8))

        ctk.CTkFrame(self._sb, fg_color=C["borde"], height=1).pack(
            fill="x", padx=8, pady=4)

        # Items del menú
        items = [
            ("🏠", "Inicio",        self._ir_inicio,    True),
            ("👤", "Estudiantes",   self._ir_estudiantes, False),
            ("📝", "Notas",         self._ir_notas,     False),
            ("📅", "Asistencia",    self._ir_asistencia, False),
            ("📋", "Reportes",      self._ir_reportes,  False),
        ]

        for icono, texto, cmd, activo in items:
            if self._sidebar_exp:
                label_txt = f"  {icono}   {texto}"
                anc = "w"
            else:
                label_txt = icono
                anc = "center"

            bg = C["activo"] if activo else "transparent"
            tc = self._acento if activo else C["texto_sec"]
            bw = 2 if activo else 0

            btn = ctk.CTkButton(
                self._sb, text=label_txt,
                fg_color=bg,
                hover_color=C["hover"],
                font=ctk.CTkFont("Segoe UI", 14 if self._sidebar_exp else 18),
                text_color=tc, anchor=anc,
                height=40, corner_radius=6,
                border_width=bw,
                border_color=self._acento,
                command=cmd)
            btn.pack(fill="x", padx=6, pady=2)

        # Espacio flexible
        ctk.CTkFrame(self._sb, fg_color="transparent").pack(
            fill="both", expand=True)

        if self._sidebar_exp:
            ctk.CTkLabel(self._sb, text="v.Prov.22:6",
                         font=ctk.CTkFont("Segoe UI", 10),
                         text_color=C["texto_dim"]).pack(pady=(0, 6))

        toggle_txt = "«" if self._sidebar_exp else "»"
        ctk.CTkButton(self._sb, text=toggle_txt, width=36, height=28,
                      fg_color=C["hover"], hover_color=C["borde"],
                      font=ctk.CTkFont(size=14, weight="bold"),
                      text_color=C["texto_sec"],
                      corner_radius=6,
                      command=self._toggle_sidebar_dash).pack(
            pady=(4, 10), padx=4)

    def _toggle_sidebar_dash(self):
        self._sidebar_exp = not self._sidebar_exp
        nuevo_ancho = 200 if self._sidebar_exp else 45
        self._sb.configure(width=nuevo_ancho)
        self._sb_renderizar()

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


class RegistroDocApp(ctk.CTk):
    def __init__(self, modalidad_inicial="premedia"):
        super().__init__()
        self.title("RegistroDoc Pro v3.0 — MEDUCA Panamá")

        # Ventana fija — no se puede redimensionar
        self.geometry(f"{ANCHO_FIJO}x{ALTO_FIJO}")
        self.minsize(ANCHO_FIJO, ALTO_FIJO)
        self.maxsize(ANCHO_FIJO, ALTO_FIJO)
        self.resizable(False, False)

        # Absolute Path for Title Bar Icon
        icon_path = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "img", "icon.ico"))
        if os.path.exists(icon_path):
            try:
                self.iconbitmap(icon_path)
            except Exception:
                pass

        # Fallback Taskbar Icon (for Linux/X11 mainly)
        for nombre_icono in [
            "icon.ico",
            "icono.jpg",
            "logo.png"
        ]:
            img_path = os.path.join(
                os.path.dirname(__file__), "..", "img", nombre_icono)
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

        # Instantiate the Main Application Structure (Header + Sidebar + Content Frame)
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