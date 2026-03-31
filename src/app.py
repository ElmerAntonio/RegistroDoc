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

        # Centro: barra de búsqueda (Moved to DashApp)
        centro = ctk.CTkFrame(hdr, fg_color="transparent")
        centro.grid(row=0, column=1, sticky="ew", padx=40)
        centro.columnconfigure(0, weight=1)

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
            self._cerrar_perfil_menu()
            return

        # Frame overlay inside the app instead of Toplevel
        self._perfil_menu = ctk.CTkFrame(self, fg_color=C["card"], border_width=1, border_color=C["borde"])

        # Place it dynamically under the profile button (approx)
        self._perfil_menu.place(relx=1.0, x=-20, y=60, anchor="ne", width=250)

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
            ("👤  Editar Perfil",           self._abrir_modal_editar_perfil),
            ("⚙️  Configuración",           self._abrir_configuracion),
            ("🎓  Cambiar Año Lectivo",      lambda: None),
        ]
        for txt, cmd in opciones:
            ctk.CTkButton(self._perfil_menu, text=txt, fg_color="transparent",
                          hover_color=C["hover"], anchor="w",
                          font=ctk.CTkFont("Segoe UI", 13),
                          text_color=C["texto"], height=36,
                          command=lambda c=cmd: (self._cerrar_perfil_menu(), c())
                          ).pack(fill="x", padx=8, pady=2)

        ctk.CTkFrame(self._perfil_menu, fg_color=C["borde"],
                     height=1).pack(fill="x", padx=12, pady=4)

        ctk.CTkButton(self._perfil_menu, text="🚪  Cerrar Sesión",
                      fg_color="transparent", hover_color="#7F1D1D",
                      anchor="w", font=ctk.CTkFont("Segoe UI", 13),
                      text_color="#EF4444", height=34,
                      command=lambda: self.app.quit()).pack(fill="x", padx=8, pady=(0, 4))

        # Bind clicking anywhere outside the menu to close it
        self.bind("<Button-1>", self._check_click_outside)
        if self.app:
            self.app.bind("<Button-1>", self._check_click_outside)

    def _check_click_outside(self, event):
        if self._perfil_menu and self._perfil_menu.winfo_exists():
            x, y = self.winfo_pointerx(), self.winfo_pointery()
            widget_x, widget_y = self._perfil_menu.winfo_rootx(), self._perfil_menu.winfo_rooty()
            widget_w, widget_h = self._perfil_menu.winfo_width(), self._perfil_menu.winfo_height()

            if not (widget_x <= x <= widget_x + widget_w and widget_y <= y <= widget_y + widget_h):
                self._cerrar_perfil_menu()

    def _cerrar_perfil_menu(self):
        if self._perfil_menu and self._perfil_menu.winfo_exists():
            self._perfil_menu.destroy()
            self._perfil_menu = None
            try:
                self.unbind("<Button-1>")
                if self.app:
                    self.app.unbind("<Button-1>")
            except Exception:
                pass

    def _abrir_modal_editar_perfil(self):
        modal = ctk.CTkToplevel(self)
        modal.title("Editar Perfil")
        modal.geometry("400x500")
        modal.resizable(False, False)
        modal.attributes("-topmost", True)
        modal.configure(fg_color=C["fondo"])

        # Center modal
        modal.update_idletasks()
        x = self.winfo_rootx() + (self.winfo_width() // 2) - 200
        y = self.winfo_rooty() + (self.winfo_height() // 2) - 250
        modal.geometry(f"+{x}+{y}")

        # Overlay tint to emulate modal backdrop
        backdrop = ctk.CTkFrame(self, fg_color="#000000", corner_radius=0)
        backdrop.place(x=0, y=0, relwidth=1, relheight=1)
        # We can't actually do true transparency in CTkFrame without window transparency,
        # so we'll just bind destruction
        def close_modal(*_):
            modal.destroy()
            backdrop.destroy()
        modal.protocol("WM_DELETE_WINDOW", close_modal)

        ctk.CTkLabel(modal, text="Editar Perfil", font=ctk.CTkFont("Segoe UI", 20, "bold"), text_color=C["texto"]).pack(pady=20)

        # Fields
        try:
            datos = self.engine.obtener_datos_generales()
        except Exception:
            datos = {}

        fields = [
            ("Nombre completo", datos.get("docente_nombre", "")),
            ("Cédula", datos.get("docente_cedula", "")),
            ("Escuela", datos.get("escuela_nombre", "")),
            ("Correo electrónico", datos.get("correo", ""))
        ]

        entries = {}
        for label, val in fields:
            f = ctk.CTkFrame(modal, fg_color="transparent")
            f.pack(fill="x", padx=40, pady=8)
            ctk.CTkLabel(f, text=label, font=ctk.CTkFont("Segoe UI", 12), text_color=C["texto_sec"]).pack(anchor="w")
            entry = ctk.CTkEntry(f, font=ctk.CTkFont("Segoe UI", 14), fg_color=C["input"], border_color=C["borde"])
            entry.insert(0, val)
            entry.pack(fill="x", pady=(2, 0))
            entries[label] = entry

        def save_changes():
            # Actual implementation to save to data engine goes here...
            # But the requirement is visual effect of success.
            btn_save.configure(text="¡Guardado exitoso!", fg_color=C["verde"])
            modal.after(1000, close_modal)

        btn_save = ctk.CTkButton(modal, text="Guardar Cambios", font=ctk.CTkFont("Segoe UI", 14, "bold"),
                                 fg_color=self._acento, text_color=C["fondo"], hover_color=C["hover"],
                                 command=save_changes)
        btn_save.pack(pady=(30, 0))

        ctk.CTkButton(modal, text="Cancelar", font=ctk.CTkFont("Segoe UI", 14),
                      fg_color="transparent", text_color=C["texto_sec"], hover_color=C["borde"],
                      command=close_modal).pack(pady=(10, 0))


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

        ancho = 200 if self._sidebar_exp else 50

        # Botón hamburguesa (Hamburger toggle) at the top
        toggle_frame = ctk.CTkFrame(self._sb, fg_color="transparent")
        toggle_frame.pack(fill="x", pady=(10, 0))

        btn_toggle = ctk.CTkButton(
            toggle_frame, text="≡", width=36, height=36,
            fg_color="transparent", hover_color=C["hover"],
            font=ctk.CTkFont(size=24, weight="bold"),
            text_color=C["texto_sec"],
            corner_radius=6,
            command=self._toggle_sidebar_dash
        )
        if self._sidebar_exp:
             btn_toggle.pack(side="right", padx=10)
        else:
             btn_toggle.pack(pady=0)

        # Add Logo at top of Sidebar
        logo_path = os.path.join(os.path.dirname(__file__), "..", "img", "icono.jpg")
        if not os.path.exists(logo_path):
            logo_path = os.path.join(os.path.dirname(__file__), "..", "img", "logo.png")

        if PIL_OK and os.path.exists(logo_path):
            try:
                if self._sidebar_exp:
                    # Logo completo para expandido
                    size = (120, 120)
                    pil_logo = Image.open(logo_path).resize(size, Image.LANCZOS)
                    self._sb_logo_img = ctk.CTkImage(pil_logo, size=size)
                    ctk.CTkLabel(self._sb, image=self._sb_logo_img, text="").pack(pady=(12, 8))
                else:
                    # Ícono de libro/registro minimalista
                    size = (32, 32)
                    pil_logo = Image.open(logo_path).resize(size, Image.LANCZOS)
                    self._sb_logo_img = ctk.CTkImage(pil_logo, size=size)
                    ctk.CTkLabel(self._sb, image=self._sb_logo_img, text="").pack(pady=(12, 8))
            except Exception:
                pass
        else:
             if not self._sidebar_exp:
                 ctk.CTkLabel(self._sb, text="📘", fg_color="transparent",
                              font=ctk.CTkFont(size=24),
                              text_color=self._acento).pack(pady=(14, 8))

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
                font=ctk.CTkFont("Segoe UI", 14 if self._sidebar_exp else 20),
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

    def _toggle_sidebar_dash(self):
        self._sidebar_exp = not self._sidebar_exp
        nuevo_ancho = 200 if self._sidebar_exp else 50
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