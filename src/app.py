import os
from config import BASE_DIR
import sys
import ctypes
import json
import tkinter as tk
import customtkinter as ctk

from utils.dialogs import patch_messagebox
patch_messagebox()

# Matplotlib
import matplotlib
matplotlib.use("TkAgg")

try:
    from PIL import Image, ImageTk
    PIL_OK = True
except ImportError:
    Image = None
    ImageTk = None
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
from happ   import ReportesYGraficosFrame
from impp   import ImpresionFrame
from habapp import HabitosFrame
from helpapp import HelpFrame
from tareasapp import TareasFrame
from reunionesapp import ReunionesFrame
from registrocompletoapp import RegistroCompletoFrame

# Cargar el tema visual desde la configuración cifrada
_tema = "dark"
try:
    from rdsecurity import cargar_config_segura
    _cfg = cargar_config_segura({"tema": "dark"})
    _tema = _cfg.get("tema", "dark")
except Exception:
    pass

ctk.set_appearance_mode(_tema)
ctk.set_default_color_theme("blue")



from theme import C, FONT_TITLE, FONT_BODY, FONT_SIZES

class MainApplication(ctk.CTkFrame):


    # ══════════════════════════════════════════════════════════════════════
    #  CUERPO: TOP HEADER (LOGO, TITLE, GROUPS) + SUB-MENU BAR
    # ══════════════════════════════════════════════════════════════════════
    def _cuerpo(self):
        # Header (Fila 0)
        self._header = ctk.CTkFrame(self, fg_color=C["header"], corner_radius=0, height=48)
        self._header.grid(row=0, column=0, sticky="ew")
        self._header.grid_propagate(False)
        
        # Info versión
        ctk.CTkLabel(self._header, text="v.Prov.22:6", font=ctk.CTkFont(family=FONT_BODY, size=10), text_color=C["texto_dim"]).pack(side="right", padx=15, pady=13)
        
        # Breadcrumb de ubicación
        self._breadcrumb = ctk.CTkLabel(self._header, text="📍 Inicio",
            font=ctk.CTkFont(family=FONT_BODY, size=11),
            text_color=C["texto_sec"], fg_color=C["badge_bg"],
            corner_radius=8, padx=10, pady=3)
        self._breadcrumb.pack(side="right", padx=8, pady=12)
        
        # Título del header
        ctk.CTkLabel(self._header, text="RegistroDoc Pro — Control Académico", font=ctk.CTkFont(family=FONT_TITLE, size=14, weight="bold"), text_color=C["cian"]).pack(side="left", padx=20, pady=13)

        # Contenedor Cuerpo (Fila 1) - Ocupa todo el resto del alto
        cuerpo = ctk.CTkFrame(self, fg_color=C["fondo"], corner_radius=0)
        cuerpo.grid(row=1, column=0, sticky="nsew")
        cuerpo.rowconfigure(0, weight=1)
        cuerpo.columnconfigure(0, weight=0) # Sidebar
        cuerpo.columnconfigure(1, weight=1) # Contenido principal

        self._sidebar_widget(cuerpo)

        # Contenedor de contenido principal (ocupa el resto de la pantalla)
        self.main_content_frame = ctk.CTkFrame(cuerpo, fg_color=C["fondo"], corner_radius=0)
        self.main_content_frame.grid(row=0, column=1, sticky="nsew", padx=0, pady=0)
        self.main_content_frame.columnconfigure(0, weight=1)
        self.main_content_frame.rowconfigure(0, weight=1)

    def _sidebar_widget(self, parent):
        self._sb = ctk.CTkScrollableFrame(
            parent,
            fg_color=C["sidebar"],
            width=210,
            corner_radius=0,
            scrollbar_button_color=C["borde"],
            scrollbar_fg_color=C["sidebar"]
        )
        self._sb.grid(row=0, column=0, sticky="nsew")
        self._sb_renderizar()

    def __init__(self, master, engine, app_principal, **kwargs):
        super().__init__(master, fg_color=C["fondo"], corner_radius=0, **kwargs)
        self.engine = engine
        self.app = app_principal
        self._acento = C["cian"]
        self.menu_activo = "Inicio"
        self._destroyed = False  # Bandera para prevenir operaciones después de destrucción

        self.pack_propagate(False)
        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=0)   # Header
        self.rowconfigure(1, weight=1)   # Cuerpo

        self._cuerpo()

    def _sb_renderizar(self):
        # Prevenir operaciones si la aplicación ya fue destruida
        if self._destroyed or not self.winfo_exists():
            return
        if getattr(self.app, "_destroyed", False):
            return
        if tk._default_root is None:
            return

        for w in self._sb.winfo_children():
            w.destroy()

        # Renderizar logo
        logo_path = os.path.join(BASE_DIR, "..", "img", "icono.png")
        if PIL_OK and os.path.exists(logo_path):
            try:
                size = (80, 80)
                pil_logo = Image.open(logo_path).resize(size, Image.Resampling.LANCZOS)
                self._sb_logo_img = ctk.CTkImage(pil_logo, size=size)
                ctk.CTkLabel(self._sb, image=self._sb_logo_img, text="").pack(pady=(20, 5))
            except Exception:
                pass
        else:
            ctk.CTkLabel(self._sb, text="📘", font=ctk.CTkFont(family=FONT_TITLE, size=24), fg_color="transparent").pack(pady=(20, 5))

        # Título
        ctk.CTkLabel(self._sb, text="RegistroDoc Pro", font=ctk.CTkFont(family=FONT_TITLE, size=15, weight="bold"), text_color=C["cian"]).pack(pady=(0, 2))
        ctk.CTkLabel(self._sb, text="Grupo Consejero 8° B", font=ctk.CTkFont(family=FONT_BODY, size=10), text_color=C["texto_dim"]).pack(pady=(0, 10))

        # Separador
        ctk.CTkFrame(self._sb, fg_color=C["borde"], height=1).pack(fill="x", padx=15, pady=(5, 15))

        # Botones de navegación
        items = [
            ("🏠", "Inicio",        self._ir_inicio),
            ("👤", "Estudiantes",   self._ir_estudiantes),
            ("📝", "Notas",         self._ir_notas),
            ("📅", "Asistencia",    self._ir_asistencia),
            ("📋", "Registro Completo", self._ir_registro_completo),
            ("🔍", "Observaciones", self._ir_observaciones),
            ("🧠", "Hábitos",       self._ir_habitos),
            ("📊", "Reportes y Gráficos", self._ir_reportes),
            ("🤝", "Reuniones",     self._ir_reuniones),
            ("🖨️", "Impresión",     self._ir_impresion),
            ("⚙️", "Configuración", self._ir_configuracion),
            ("❓", "Ayuda y Guía",  self._ir_ayuda)
        ]

        self.nav_buttons = {}
        for icono, texto, cmd in items:
            activo = (self.menu_activo == texto)
            bg = C["activo"] if activo else "transparent"
            tc = self._acento if activo else C["texto_sec"]
            bw = 1.5 if activo else 0

            label_txt = f"  {icono}   {texto}"

            def make_cmd(c=cmd, t=texto):
                def wrapper():
                    if self._destroyed or not self.winfo_exists():
                        return
                    if getattr(self.app, "_destroyed", False):
                        return
                    if tk._default_root is None:
                        return
                    self.menu_activo = t
                    try:
                        self._sb_renderizar()
                        c()
                    except (tk.TclError, RuntimeError):
                        return
                return wrapper

            try:
                btn = ctk.CTkButton(
                    self._sb,
                    text=label_txt,
                    fg_color=bg,
                    hover_color=C["hover"],
                    font=ctk.CTkFont(family=FONT_BODY, size=13, weight="bold"),
                    text_color=tc,
                    anchor="w",
                    height=38,
                    corner_radius=6,
                    border_width=bw,
                    border_color=self._acento,
                    command=make_cmd()
                )
                btn.pack(fill="x", padx=12, pady=3)
                self.nav_buttons[texto] = btn
            except (tk.TclError, RuntimeError):
                return

    def destroy(self):
        """Override destroy para marcar la aplicación como destruida."""
        self._destroyed = True
        super().destroy()

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
        if self.app:
            try: self.app.mostrar_reportes()
            except Exception: pass

    def _ir_graficos(self):
        if self.app:
            try: self.app.mostrar_graficos()
            except Exception: pass

    def _ir_observaciones(self):
        if self.app:
            try: self.app.mostrar_observaciones()
            except Exception: pass

    def _ir_habitos(self):
        if self.app:
            try: self.app.mostrar_habitos()
            except Exception: pass

    def _ir_impresion(self):
        if self.app:
            try: self.app.mostrar_impresion()
            except Exception: pass

    def _ir_configuracion(self):
        if self.app:
            try: self.app.mostrar_configuracion()
            except Exception: pass

    def _ir_tareas(self):
        if self.app:
            try: self.app.mostrar_tareas()
            except Exception: pass

    def _ir_reuniones(self):
        if self.app:
            try: self.app.mostrar_reuniones()
            except Exception: pass

    def _ir_ayuda(self):
        if self.app:
            try: self.app.mostrar_ayuda()
            except Exception: pass

    def _ir_registro_completo(self):
        if self.app:
            try: self.app.mostrar_registro_completo()
            except Exception: pass


class RegistroDocApp(ctk.CTk):
    def __init__(self, modalidad_inicial="premedia"):
        super().__init__()
        self._destroyed = False  # Bandera para prevenir operaciones después de destrucción
        self._frames = {}
        
        self.title("RegistroDoc Pro v3.0 — MEDUCA Panamá")

        # ─── VENTANA REDIMENSIONABLE Y MAXIMIZABLE ───
        self.geometry("1280x720")
        self.minsize(1024, 600)  # Tamaño mínimo para que no se deforme
        self.resizable(True, True) # Permite maximizar y achicar
        self.protocol("WM_DELETE_WINDOW", self.on_closing)


        # Iconos de la ventana (resolución para Windows / Barra de tareas)
        icon_path = os.path.join(BASE_DIR, "..", "img", "icon.ico")
        png_path = os.path.join(BASE_DIR, "..", "img", "icono.png")

        if os.path.exists(icon_path):
            try:
                self.iconbitmap(icon_path)
            except Exception:
                pass

        # Fallback multi-plataforma
        if os.path.exists(png_path) and getattr(self, "iconphoto", None):
            try:
                if PIL_OK:
                    pil = Image.open(png_path).resize((64, 64))
                    self._icono_app = ImageTk.PhotoImage(pil)
                    self.iconphoto(True, self._icono_app)
            except Exception:
                pass

        # Motor de datos
        archivo = ("Registro_Primaria.xlsx"
                   if modalidad_inicial == "primaria"
                   else "Registro_2026.xlsx")
        ruta    = os.path.join(BASE_DIR, "..", archivo)
        self.engine = DataEngine(ruta_excel=ruta, modalidad=modalidad_inicial)

        # Contenedor raíz
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)

        # Iniciar Aplicación
        self.main_app = MainApplication(self, self.engine, app_principal=self)
        self.main_app.grid(row=0, column=0, sticky="nsew")

        # ─── ATAJOS DE TECLADO GLOBALES ───
        self._registrar_atajos()

        # ─── AUTO-BACKUP CADA 30 MIN ───
        self._iniciar_auto_backup()

        self.mostrar_dashboard()

    def limpiar_pantalla(self):
        for w in self.main_app.main_content_frame.winfo_children():
            w.destroy()

    def mostrar_toast(self, mensaje, color="#10B981"):
        if hasattr(self, "_toast_widget"):
            try:
                if self._toast_widget.winfo_exists():
                    self._toast_widget.destroy()
            except Exception:
                pass
        toast = ctk.CTkFrame(self, fg_color=color, corner_radius=20, border_width=0)
        toast.place(relx=0.5, rely=0.92, anchor="center")
        lbl = ctk.CTkLabel(toast, text=f"  {mensaje}  ", font=(FONT_BODY, 13, "bold"),
                           text_color="white", padx=16, pady=10)
        lbl.pack(side="left")
        ctk.CTkButton(toast, text="✕", width=28, height=28, fg_color="transparent",
                      hover_color=color, text_color="white",
                      font=(FONT_BODY, 14, "bold"),
                      command=lambda: _cerrar()).pack(side="right", padx=(0, 6))
        self._toast_widget = toast
        def _cerrar():
            try:
                if toast.winfo_exists():
                    toast.destroy()
            except Exception:
                pass
        self.after(4000, _cerrar)

    def _mostrar_frame(self, frame_class, *args, **kwargs):
        # Ocultar todos los frames en la caché
        for f in self._frames.values():
            try:
                f.pack_forget()
            except Exception:
                pass
            
        # Si el frame no está en la caché, crearlo
        class_name = frame_class.__name__
        if class_name not in self._frames:
            f = frame_class(self.main_app.main_content_frame, self.engine, *args, **kwargs)
            self._frames[class_name] = f
            
        # Mostrar el frame activo
        active_frame = self._frames[class_name]
        active_frame.pack(fill="both", expand=True)
        
        # Invocar actualizar_vista si existe
        if hasattr(active_frame, "actualizar_vista"):
            try:
                active_frame.actualizar_vista()
            except Exception as e:
                print(f"[!] Error actualizando vista {class_name}: {e}")
                
        return active_frame

    def _actualizar_breadcrumb(self, seccion):
        """Actualiza el indicador de ubicación en el header."""
        nombres = {
            "DashboardFrame": "📍 Inicio",
            "EstudiantesFrame": "📍 Inicio › Estudiantes",
            "NotasFrame": "📍 Inicio › Notas",
            "AsistenciaFrame": "📍 Inicio › Asistencia",
            "ObservacionesFrame": "📍 Inicio › Observaciones",
            "HabitosFrame": "📍 Inicio › Hábitos",
            "ReportesYGraficosFrame": "📍 Inicio › Reportes y Gráficos",
            "ImpresionFrame": "📍 Inicio › Impresión",
            "ConfigFrame": "📍 Inicio › Configuración",
            "HelpFrame": "📍 Inicio › Ayuda",
            "TareasFrame": "📍 Inicio › Tareas",
            "ReunionesFrame": "📍 Inicio › Reuniones",
            "RegistroCompletoFrame": "📍 Inicio › Registro Completo",
        }
        txt = nombres.get(seccion, "📍 Inicio")
        try:
            if hasattr(self.main_app, '_breadcrumb'):
                self.main_app._breadcrumb.configure(text=txt)
        except Exception:
            pass

    def mostrar_dashboard(self):
        self.main_app.menu_activo = "Inicio"
        self.main_app._sb_renderizar()
        self._mostrar_frame(DashboardFrame, app_principal=self)
        self._actualizar_breadcrumb("DashboardFrame")

    def mostrar_registro_completo(self):
        self.main_app.menu_activo = "Registro Completo"
        self.main_app._sb_renderizar()
        self._mostrar_frame(RegistroCompletoFrame)
        self._actualizar_breadcrumb("RegistroCompletoFrame")

    def mostrar_estudiantes(self):
        self.main_app.menu_activo = "Estudiantes"
        self.main_app._sb_renderizar()
        self._mostrar_frame(EstudiantesFrame)
        self._actualizar_breadcrumb("EstudiantesFrame")

    def mostrar_notas(self):
        self.main_app.menu_activo = "Notas"
        self.main_app._sb_renderizar()
        self._mostrar_frame(NotasFrame)
        self._actualizar_breadcrumb("NotasFrame")

    def mostrar_asistencia(self):
        self.main_app.menu_activo = "Asistencia"
        self.main_app._sb_renderizar()
        self._mostrar_frame(AsistenciaFrame)
        self._actualizar_breadcrumb("AsistenciaFrame")

    def mostrar_observaciones(self):
        self.main_app.menu_activo = "Observaciones"
        self.main_app._sb_renderizar()
        self._mostrar_frame(ObservacionesFrame)
        self._actualizar_breadcrumb("ObservacionesFrame")

    def mostrar_reportes(self, tab_inicial="📋 Reportes y Descargas"):
        self.main_app.menu_activo = "Reportes y Gráficos"
        self.main_app._sb_renderizar()
        frame = self._mostrar_frame(ReportesYGraficosFrame)
        self._actualizar_breadcrumb("ReportesYGraficosFrame")
        if tab_inicial and hasattr(frame, "tabs"):
            try:
                frame.tabs.set(tab_inicial)
            except Exception:
                pass

    def mostrar_graficos(self):
        self.mostrar_reportes(tab_inicial="📈 Gráficos de Rendimiento")

    def mostrar_impresion(self):
        self.main_app.menu_activo = "Impresión"
        self.main_app._sb_renderizar()
        self._mostrar_frame(ImpresionFrame)
        self._actualizar_breadcrumb("ImpresionFrame")

    def mostrar_configuracion(self):
        self.main_app.menu_activo = "Configuración"
        self.main_app._sb_renderizar()
        self._mostrar_frame(ConfigFrame, self)
        self._actualizar_breadcrumb("ConfigFrame")

    def mostrar_habitos(self):
        self.main_app.menu_activo = "Hábitos"
        self.main_app._sb_renderizar()
        self._mostrar_frame(HabitosFrame)
        self._actualizar_breadcrumb("HabitosFrame")

    def mostrar_ayuda(self):
        self.main_app.menu_activo = "Ayuda y Guía"
        self.main_app._sb_renderizar()
        self._mostrar_frame(HelpFrame)
        self._actualizar_breadcrumb("HelpFrame")

    def mostrar_tareas(self):
        self.main_app.menu_activo = "Tareas"
        self.main_app._sb_renderizar()
        self._mostrar_frame(TareasFrame)
        self._actualizar_breadcrumb("TareasFrame")

    def mostrar_reuniones(self):
        self.main_app.menu_activo = "Reuniones"
        self.main_app._sb_renderizar()
        self._mostrar_frame(ReunionesFrame)
        self._actualizar_breadcrumb("ReunionesFrame")

    def reiniciar_motor(self, nueva_ruta, nueva_modalidad):
        self.engine = DataEngine(ruta_excel=nueva_ruta, modalidad=nueva_modalidad)
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
            data["modalidad"] = nueva_modalidad
            with open(CONFIG_FILE, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=4)
        except Exception:
            pass
        # Limpiar caché de frames
        for f in list(self._frames.values()):
            try: f.destroy()
            except Exception: pass
        self._frames.clear()

        self.main_app.engine = self.engine
        self.mostrar_dashboard()


    # ─── ATAJOS DE TECLADO ──────────────────────────────────────────
    def _registrar_atajos(self):
        """Registra atajos de teclado globales para navegación rápida."""
        atajos = {
            "<Control-Key-1>": lambda e: self.mostrar_dashboard(),
            "<Control-Key-2>": lambda e: self.mostrar_estudiantes(),
            "<Control-Key-3>": lambda e: self.mostrar_notas(),
            "<Control-Key-4>": lambda e: self.mostrar_asistencia(),
            "<Control-Key-5>": lambda e: self.mostrar_observaciones(),
            "<Control-Key-6>": lambda e: self.mostrar_habitos(),
            "<Control-Key-7>": lambda e: self.mostrar_reportes(),
            "<Control-Key-8>": lambda e: self.mostrar_graficos(),
            "<Control-Key-9>": lambda e: self.mostrar_impresion(),
            "<F1>": lambda e: self.mostrar_ayuda(),
            "<Escape>": lambda e: self.mostrar_dashboard(),
        }
        for atajo, cmd in atajos.items():
            try:
                self.bind(atajo, cmd)
            except Exception:
                pass

    # ─── AUTO-BACKUP ───────────────────────────────────────────────
    def _iniciar_auto_backup(self):
        """Programa un respaldo automático cada 30 minutos."""
        self._auto_backup_interval = 30 * 60 * 1000  # 30 min en ms
        self._programar_backup()

    def _programar_backup(self):
        if self._destroyed:
            return
        try:
            self._ejecutar_auto_backup()
        except Exception:
            pass
        try:
            self.after(self._auto_backup_interval, self._programar_backup)
        except Exception:
            pass

    def _ejecutar_auto_backup(self):
        import shutil, datetime
        try:
            ruta_original = self.engine.ruta
            if not os.path.exists(ruta_original):
                return
            dir_base = os.path.dirname(os.path.abspath(ruta_original))
            dir_respaldos = os.path.join(dir_base, "Respaldos_Auto")
            if not os.path.exists(dir_respaldos):
                os.makedirs(dir_respaldos)
            ahora = datetime.datetime.now().strftime("%Y-%m-%d_%H-%M")
            nombre_base = os.path.basename(ruta_original).replace(".xlsx", "")
            nombre = f"{nombre_base}_auto_{ahora}.xlsx"
            shutil.copy2(ruta_original, os.path.join(dir_respaldos, nombre))
            # Limpiar: máximo 10 auto-backups
            respaldos = sorted(
                [os.path.join(dir_respaldos, f) for f in os.listdir(dir_respaldos) if f.endswith(".xlsx")],
                key=os.path.getmtime
            )
            while len(respaldos) > 10:
                try: os.remove(respaldos.pop(0))
                except: pass
        except Exception:
            pass

    def on_closing(self):
        """Limpieza segura y salida sin destruir widgets en carrera."""
        if self._destroyed:
            return
        self._destroyed = True
        try:
            if hasattr(self, "main_app"):
                self.main_app._destroyed = True
        except Exception:
            pass
        try:
            self.withdraw()
        except Exception:
            pass
        try:
            self.quit()
        except Exception:
            pass
        try:
            super().destroy()
        except Exception:
            pass

    def destroy(self):
        """Override destroy para marcar la aplicación como destruida y limpiar recursos."""
        self._destroyed = True
        try:
            if hasattr(self, "main_app"):
                self.main_app._destroyed = True
        except Exception:
            pass
        try:
            super().destroy()
        except Exception:
            pass


def iniciar_programa_principal():
    if sys.platform == "win32":
        try:
            import ctypes
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("registrodoc.pro.v3")
        except Exception:
            pass
    try:
        from rdsecurity import cargar_config_segura
        config = cargar_config_segura({"modalidad": "premedia"})
    except Exception:
        config = {"modalidad": "premedia"}

    # Inicio directo sin ventana de carga.
    app = RegistroDocApp(
        modalidad_inicial=config.get("modalidad", "premedia"))
    try:
        app.lift()
        app.focus_force()
    except Exception:
        pass

    try:
        app.mainloop()
    except KeyboardInterrupt:
        # Si se interrumpe desde consola, cerrar limpio sin traceback ruidoso.
        try:
            app.on_closing()
        except Exception:
            pass


if __name__ == "__main__":
    from rdsecurity import cargar_config_segura, CONFIG_FILE, guardar_config_segura
    # Esto dispara la migración de perfil.json a config.enc automáticamente si existe
    cargar_config_segura({})

    if not os.path.exists(CONFIG_FILE):
        if SetupWizard:
            wizard = SetupWizard()
            wizard.mainloop()
            if os.path.exists(CONFIG_FILE):
                iniciar_programa_principal()
        else:
            cfg = {"modalidad": "premedia",
                   "docente_nombre": "", "ano_lectivo": "2026"}
            guardar_config_segura(cfg)
            iniciar_programa_principal()
    else:
        iniciar_programa_principal()
