import os
import sys
try:
    from anti_analysis import run_anti_analysis_checks
    run_anti_analysis_checks()
except Exception:
    pass

from config import BASE_DIR, ASSETS_DIR
import ctypes
import json
import tkinter as tk
import customtkinter as ctk
from utils.translator import tr

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
from notasasistenciaapp import NotasAsistenciaFrame
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
    def _obtener_nombre_menu_breadcrumb(self):
        nombres = {
            "Inicio": tr("Inicio"),
            "Estudiantes": tr("Inicio › Estudiantes"),
            "Notas y Asistencia": tr("Inicio › Notas y Asistencia"),
            "Observaciones": tr("Inicio › Observaciones"),
            "Reportes y Gráficos": tr("Inicio › Reportes y Gráficos"),
            "Impresión": tr("Inicio › Impresión"),
            "Hábitos": tr("Inicio › Hábitos"),
            "Ayuda y Guía": tr("Inicio › Ayuda y Guía"),
            "Tareas Programadas": tr("Inicio › Tareas Programadas"),
            "Reuniones": tr("Inicio › Reuniones"),
            "Configuración": tr("Inicio › Configuración"),
            "Registro Completo": tr("Inicio › Registro Completo")
        }
        return nombres.get(self.menu_activo, tr("Inicio"))


    def _renderizar_header(self):
        # Si ya se creó el header, solo actualizar el breadcrumb
        if hasattr(self, "_header_created") and self._header_created:
            try:
                if self._breadcrumb.winfo_exists():
                    self._breadcrumb.configure(text="📍 " + self._obtener_nombre_menu_breadcrumb())
            except Exception:
                pass
            return

        self._header = ctk.CTkFrame(self, fg_color=C["header"], corner_radius=0, height=48)
        self._header.grid(row=0, column=0, sticky="ew")
        self._header.grid_propagate(False)
        
        # Info versión
        ctk.CTkLabel(self._header, text="v.Prov.22:6", font=ctk.CTkFont(family=FONT_BODY, size=10), text_color=C["texto_dim"]).pack(side="right", padx=15, pady=13)
        
        # Breadcrumb de ubicación
        self._breadcrumb = ctk.CTkLabel(self._header, text="📍 " + self._obtener_nombre_menu_breadcrumb(),
            font=ctk.CTkFont(family=FONT_BODY, size=11),
            text_color=C["texto_sec"], fg_color=C["badge_bg"],
            corner_radius=8, padx=10, pady=3)
        self._breadcrumb.pack(side="right", padx=8, pady=12)
        
        # Insignia de Sincronización Horaria (Panamá)
        from utils.date_helpers import es_hora_sincronizada
        sync_ok = es_hora_sincronizada()
        sync_text = "🕒 Panamá OK" if sync_ok else "🕒 Hora Local"
        sync_color = "#00FF88" if sync_ok else "#94A3B8"
        self._time_sync_badge = ctk.CTkLabel(self._header, text=sync_text,
            font=ctk.CTkFont(family=FONT_BODY, size=11, weight="bold"),
            text_color=sync_color, fg_color=C["badge_bg"],
            corner_radius=8, padx=10, pady=3)
        self._time_sync_badge.pack(side="right", padx=8, pady=12)
        
        # Badge de Horario Escolar / Clase Actual
        self._clase_badge = ctk.CTkLabel(self._header, text="🕐 Cargando...",
            font=ctk.CTkFont(family=FONT_BODY, size=11, weight="bold"),
            text_color="#94A3B8", fg_color=C["badge_bg"],
            corner_radius=8, padx=10, pady=3)
        self._clase_badge.pack(side="right", padx=8, pady=12)
        # Título del header
        ctk.CTkLabel(self._header, text="RegistroDoc Pro — Control Académico", font=ctk.CTkFont(family=FONT_TITLE, size=14, weight="bold"), text_color=C["cian"]).pack(side="left", padx=20, pady=13)
        
        self._header_created = True

    def _cuerpo(self):
        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=0)   # Header
        self.rowconfigure(1, weight=1)   # Cuerpo
        
        self._renderizar_header()

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
        self._timer_horario_running = True
        self._actualizar_horario_header()

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
        logo_path = os.path.join(ASSETS_DIR, "icono.png")
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
            ("📝", "Notas y Asistencia", self._ir_notas_asistencia),
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

            label_txt = f"  {icono}   {tr(texto)}"

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
                        self._highlight_menu_button(t)
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

    def _actualizar_horario_header(self):
        if self._destroyed or not self.winfo_exists():
            return
        if getattr(self.app, "_destroyed", False):
            return
        if tk._default_root is None:
            return
            
        try:
            from tareas import obtener_clase_actual, obtener_proxima_clase
            from utils.date_helpers import es_hora_sincronizada, obtener_ahora_panama
            
            # 1. Actualizar el estado de sincronización horaria
            sync_ok = es_hora_sincronizada()
            if hasattr(self, "_time_sync_badge") and self._time_sync_badge.winfo_exists():
                sync_text = "🕒 Panamá OK" if sync_ok else "🕒 Hora Local"
                sync_color = "#00FF88" if sync_ok else "#94A3B8"
                self._time_sync_badge.configure(text=sync_text, text_color=sync_color)
            
            # 2. Actualizar el estado de la clase actual
            horario = self.engine.obtener_horario()
            materia, rango, dia = obtener_clase_actual(horario)
            
            ahora = obtener_ahora_panama()
            
            if hasattr(self, "_clase_badge") and self._clase_badge.winfo_exists():
                if materia:
                    self._clase_badge.configure(
                        text=f"🟢 Clase: {materia} ({rango})",
                        text_color="#00FF88"
                    )
                else:
                    prox_mat, prox_rango = obtener_proxima_clase(horario)
                    if prox_mat:
                        self._clase_badge.configure(
                            text=f"🕐 Próx: {prox_mat} ({prox_rango})",
                            text_color="#38BDF8"
                        )
                    else:
                        msg = "No hay más clases hoy" if ahora.weekday() < 5 else "Fin de semana"
                        self._clase_badge.configure(
                            text=f"🕐 {msg}",
                            text_color="#94A3B8"
                        )
        except Exception as e:
            print(f"Error updating header schedule: {e}")
            
        # Actualizar cada 60 segundos (los períodos de clase no cambian en 10s)
        self.after(60000, self._actualizar_horario_header)

    def _highlight_menu_button(self, active_text):
        if self._destroyed or not self.winfo_exists():
            return
        if not hasattr(self, "nav_buttons") or not self.nav_buttons:
            return
        for t, btn in self.nav_buttons.items():
            try:
                if btn.winfo_exists():
                    if t == active_text:
                        btn.configure(
                            fg_color=C["activo"],
                            text_color=self._acento,
                            border_width=1.5,
                            border_color=self._acento
                        )
                    else:
                        btn.configure(
                            fg_color="transparent",
                            text_color=C["texto_sec"],
                            border_width=0
                        )
            except Exception:
                pass

    # Rutas de navegación
    def _ir_inicio(self):
        if self.app:
            try: self.app.mostrar_dashboard()
            except Exception: pass

    def _ir_estudiantes(self):
        if self.app:
            try: self.app.mostrar_estudiantes()
            except Exception: pass

    def _ir_notas_asistencia(self):
        if self.app:
            try: self.app.mostrar_notas()
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

        # Fondo del tema en la VENTANA RAÍZ: al maximizar/redimensionar, Windows
        # rellena el área recién expuesta con este color (brush de la ventana) en
        # vez de negro, así no se ven los márgenes negros mientras el frame interno
        # se reajusta al nuevo tamaño.
        try:
            self.configure(fg_color=C["fondo"])
        except Exception:
            pass

        # Ocultar ventana principal durante la pantalla de carga
        self.withdraw()

        # Iconos de la ventana (resolución para Windows / Barra de tareas)
        icon_path = os.path.join(ASSETS_DIR, "icon_fixed.ico")
        png_path = os.path.join(ASSETS_DIR, "icono.png")

        if os.path.exists(icon_path):
            try:
                self.iconbitmap(icon_path)
            except Exception:
                pass
        self.bind("<Configure>", self._on_window_configure)
        self.bind("<Unmap>", self._on_window_unmap)
        self.bind("<Map>", self._on_window_map)
        self._is_minimized = False
        self._last_width = 1280
        self._last_height = 720

        # Definir las tareas del Splash Screen (Arranque rápido e instantáneo)
        def task_cargar_motor(sp):
            archivo = ("Registro_Primaria.xlsx"
                       if modalidad_inicial == "primaria"
                       else "Registro_Premedia.xlsx")
            ruta = os.path.join(ASSETS_DIR, "templates", archivo)
            self.engine = DataEngine(ruta_excel=ruta, modalidad=modalidad_inicial)

        def task_cargar_ui(sp):
            self.grid_columnconfigure(0, weight=1)
            self.grid_rowconfigure(0, weight=1)
            self.main_app = MainApplication(self, self.engine, app_principal=self)
            self.main_app.grid(row=0, column=0, sticky="nsew")

        def task_sincronizacion(sp):
            self._registrar_atajos()
            if not os.environ.get("PYTEST_CURRENT_TEST"):
                self._iniciar_inactividad_check()
                self._iniciar_auto_backup()
                try:
                    from utils.date_helpers import iniciar_sincronizacion_hora_panama
                    iniciar_sincronizacion_hora_panama()
                except Exception:
                    pass

        def task_finalizar(sp):
            self.mostrar_dashboard()
            if sp:
                sp.app_instance = self
            # Pre-calentar las filas de datos de las secciones navegables mientras
            # el usuario ve el Dashboard → la primera navegación a cada una es fluida.
            if not os.environ.get("PYTEST_CURRENT_TEST"):
                self.after(700, self._precalentar_datos_en_idle)

        if os.environ.get("PYTEST_CURRENT_TEST"):
            # En entorno de pruebas (pytest), inicializar todo síncronamente y omitir el Splash Screen
            task_cargar_motor(None)
            task_cargar_ui(None)
            task_sincronizacion(None)
            self._precargar_un_frame(DashboardFrame)
            self._precargar_un_frame(EstudiantesFrame)
            self._precargar_un_frame(NotasAsistenciaFrame)
            self._precargar_un_frame(ObservacionesFrame)
            self._precargar_un_frame(HabitosFrame)
            self._precargar_un_frame(ReportesYGraficosFrame)
            self._precargar_un_frame(ImpresionFrame)
            self._precargar_un_frame(TareasFrame)
            self._precargar_un_frame(ReunionesFrame)
            self._precargar_un_frame(RegistroCompletoFrame)
            task_finalizar(None)
        else:
            # Iniciar la pantalla de carga transparente en el hilo principal
            from splash import SplashScreen
            self.splash = SplashScreen(self)

            loading_tasks = [
                ("Cargando base de datos...", task_cargar_motor),
                ("Inicializando interfaz...", task_cargar_ui),
                ("Configurando sistema...", task_sincronizacion),
                ("Panel de Control...", lambda sp: self._precargar_un_frame(DashboardFrame)),
                ("Módulo de Estudiantes...", lambda sp: self._precargar_un_frame(EstudiantesFrame)),
                ("Módulo de Calificaciones...", lambda sp: self._precargar_un_frame(NotasAsistenciaFrame)),
                ("Módulo de Observaciones...", lambda sp: self._precargar_un_frame(ObservacionesFrame)),
                ("Módulo de Hábitos...", lambda sp: self._precargar_un_frame(HabitosFrame)),
                ("Reportes y Gráficos...", lambda sp: self._precargar_un_frame(ReportesYGraficosFrame)),
                ("Módulo de Impresión...", lambda sp: self._precargar_un_frame(ImpresionFrame)),
                ("Módulo de Tareas...", lambda sp: self._precargar_un_frame(TareasFrame)),
                ("Minutas y Reuniones...", lambda sp: self._precargar_un_frame(ReunionesFrame)),
                ("Registro Consolidado...", lambda sp: self._precargar_un_frame(RegistroCompletoFrame)),
                ("Abriendo aplicación...", task_finalizar),
            ]

            def on_splash_complete(app_inst):
                if app_inst:
                    try:
                        # Pre-renderizado invisible
                        app_inst.attributes("-alpha", 0.0)
                        app_inst.deiconify()

                        # Desactivar las animaciones de DWM (maximizar/minimizar/
                        # restaurar) para ESTA ventana → sin doble-imagen/ghosting.
                        app_inst._deshabilitar_animaciones_dwm()

                        # Abrir MAXIMIZADA (a pantalla completa) como una app normal,
                        # en vez de una ventana flotante pequeña sobre otras ventanas.
                        try:
                            app_inst.state("zoomed")
                        except Exception:
                            pass

                        # Cargar icono en barra de tareas al revelar la ventana
                        icon_p = os.path.join(ASSETS_DIR, "icon_fixed.ico")
                        if os.path.exists(icon_p):
                            try:
                                app_inst.iconbitmap(icon_p)
                            except Exception:
                                pass
                                
                        app_inst.update_idletasks()
                        app_inst.update()
                        app_inst.attributes("-alpha", 1.0)
                        app_inst.lift()
                        app_inst.focus_force()
                    except Exception:
                        pass

            self.splash.set_tasks(loading_tasks, on_splash_complete)
            self.splash.iniciar()

    def _precargar_un_frame(self, frame_class):
        """Crea un frame y lo oculta inmediatamente."""
        if self._destroyed:
            return
        class_name = frame_class.__name__
        if class_name not in self._frames:
            try:
                f = frame_class(self.main_app.main_content_frame, self.engine)
                f.pack_forget()
                self._frames[class_name] = f
            except Exception as e:
                print(f"[!] Error precargando {class_name}: {e}")

    def _precalentar_datos_en_idle(self):
        """Construye los datos/filas de las secciones navegables una vez, en idle
        y fuera de la vista (el usuario está en el Dashboard), para que la primera
        navegación a cada sección no tenga freeze de construcción de filas."""
        if getattr(self, "_destroyed", False):
            return
        nombres = ["EstudiantesFrame", "NotasAsistenciaFrame", "ObservacionesFrame"]
        self._cola_precalentar = [n for n in nombres if n in self._frames]
        self._calentar_siguiente_frame()

    def _calentar_siguiente_frame(self):
        if getattr(self, "_destroyed", False):
            return
        cola = getattr(self, "_cola_precalentar", [])
        if not cola:
            return
        nombre = cola.pop(0)
        frame = self._frames.get(nombre)
        try:
            if (frame and frame.winfo_exists()
                    and hasattr(frame, "actualizar_vista")
                    and not getattr(frame, "_precalentado", False)):
                frame._precalentado = True
                frame.actualizar_vista()
        except Exception as e:
            print(f"[precalentar] {nombre}: {e}")
        # Espaciar cada sección para no encadenar trabajo pesado en el hilo principal
        self.after(250, self._calentar_siguiente_frame)

    def _mostrar_app_principal(self):
        try:
            self.attributes("-alpha", 0.0)
            self.deiconify()
            icon_p = os.path.join(ASSETS_DIR, "icon_fixed.ico")
            if os.path.exists(icon_p):
                try:
                    self.iconbitmap(icon_p)
                except Exception:
                    pass
            self.update_idletasks()
            self.update()
            self.attributes("-alpha", 1.0)
            self.lift()
            self.focus_force()
        except Exception:
            pass

    def limpiar_pantalla(self):
        for w in self.main_app.main_content_frame.winfo_children():
            w.destroy()

    def mostrar_toast(self, mensaje, color="#10B981"):
        mensaje = tr(mensaje)
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
        class_name = frame_class.__name__
        ya_activo   = (getattr(self, "_frame_activo", None) == class_name)

        # ── 1. Ocultar todos los frames en caché ──────────────────────────
        for f in self._frames.values():
            try:
                f.pack_forget()
            except Exception:
                pass

        # ── 2. Crear el frame si aún no existe ────────────────────────────
        if class_name not in self._frames:
            f = frame_class(self.main_app.main_content_frame, self.engine, *args, **kwargs)
            self._frames[class_name] = f

        # ── 3. Mostrar el frame INMEDIATAMENTE (sin bloquear) ─────────────
        active_frame = self._frames[class_name]
        active_frame.pack(fill="both", expand=True)
        self._frame_activo = class_name

        # ── 4. Determinar si los datos necesitan actualizarse ─────────────
        version_actual = 0
        try:
            if hasattr(self.engine, "db_manager") and self.engine.db_manager and self.engine.db_manager.conn:
                version_actual = self.engine.db_manager.conn.total_changes
        except Exception:
            pass

        last_ver = getattr(active_frame, "last_updated_version", -1)
        necesita_update = (
            not ya_activo and
            hasattr(active_frame, "actualizar_vista") and
            (last_ver != version_actual or not hasattr(active_frame, "last_updated_version"))
        )

        if necesita_update:
            # Marcar versión antes de lanzar para no repetir si llegan
            # varias navegaciones seguidas al mismo frame
            active_frame.last_updated_version = version_actual

            # ── 5. Diferir la actualización: el frame ya es visible
            #       cuando llega este callback → el usuario ve el cambio
            #       de sección instantáneo y los datos llegan sin freeze.
            def _defer_update(frame=active_frame, cn=class_name):
                try:
                    if not frame.winfo_exists():
                        return
                    frame.actualizar_vista()
                except Exception as e:
                    print(f"[!] Error actualizando vista {cn}: {e}")

            self.after(0, _defer_update)

        return active_frame

    def _actualizar_breadcrumb(self, seccion):
        """Actualiza el indicador de ubicación en el header."""
        try:
            if hasattr(self, "main_app") and hasattr(self.main_app, "_breadcrumb") and self.main_app._breadcrumb.winfo_exists():
                txt = self.main_app._obtener_nombre_menu_breadcrumb()
                self.main_app._breadcrumb.configure(text="📍 " + txt)
        except Exception:
            pass

    def mostrar_dashboard(self):
        self.main_app.menu_activo = "Inicio"
        self.main_app._highlight_menu_button("Inicio")
        self._mostrar_frame(DashboardFrame, app_principal=self)
        self._actualizar_breadcrumb("DashboardFrame")

    def mostrar_registro_completo(self):
        self.main_app.menu_activo = "Registro Completo"
        self.main_app._highlight_menu_button("Registro Completo")
        self._mostrar_frame(RegistroCompletoFrame)
        self._actualizar_breadcrumb("RegistroCompletoFrame")

    def mostrar_estudiantes(self):
        self.main_app.menu_activo = "Estudiantes"
        self.main_app._highlight_menu_button("Estudiantes")
        self._mostrar_frame(EstudiantesFrame)
        self._actualizar_breadcrumb("EstudiantesFrame")

    def mostrar_notas(self):
        self.main_app.menu_activo = "Notas y Asistencia"
        self.main_app._highlight_menu_button("Notas y Asistencia")
        frame = self._mostrar_frame(NotasAsistenciaFrame)
        frame.tabview.set("Notas")
        self._actualizar_breadcrumb("NotasAsistenciaFrame")

    def mostrar_asistencia(self):
        self.main_app.menu_activo = "Notas y Asistencia"
        self.main_app._highlight_menu_button("Notas y Asistencia")
        frame = self._mostrar_frame(NotasAsistenciaFrame)
        frame.tabview.set("Asistencia")
        self._actualizar_breadcrumb("NotasAsistenciaFrame")

    def mostrar_observaciones(self):
        self.main_app.menu_activo = "Observaciones"
        self.main_app._highlight_menu_button("Observaciones")
        self._mostrar_frame(ObservacionesFrame)
        self._actualizar_breadcrumb("ObservacionesFrame")

    def mostrar_reportes(self, tab_inicial="📋 Reportes y Descargas"):
        self.main_app.menu_activo = "Reportes y Gráficos"
        self.main_app._highlight_menu_button("Reportes y Gráficos")
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
        self.main_app._highlight_menu_button("Impresión")
        self._mostrar_frame(ImpresionFrame)
        self._actualizar_breadcrumb("ImpresionFrame")

    def mostrar_configuracion(self):
        self.main_app.menu_activo = "Configuración"
        self.main_app._highlight_menu_button("Configuración")
        self._mostrar_frame(ConfigFrame, self)
        self._actualizar_breadcrumb("ConfigFrame")

    def mostrar_habitos(self):
        self.main_app.menu_activo = "Hábitos"
        self.main_app._highlight_menu_button("Hábitos")
        self._mostrar_frame(HabitosFrame)
        self._actualizar_breadcrumb("HabitosFrame")

    def mostrar_ayuda(self):
        self.main_app.menu_activo = "Ayuda y Guía"
        self.main_app._highlight_menu_button("Ayuda y Guía")
        self._mostrar_frame(HelpFrame)
        self._actualizar_breadcrumb("HelpFrame")

    def mostrar_tareas(self):
        self.main_app.menu_activo = "Tareas"
        self.main_app._highlight_menu_button("Tareas")
        self._mostrar_frame(TareasFrame)
        self._actualizar_breadcrumb("TareasFrame")

    def mostrar_reuniones(self):
        self.main_app.menu_activo = "Reuniones"
        self.main_app._highlight_menu_button("Reuniones")
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
            "<Control-Key-s>": lambda e: self.ejecutar_atajo_guardar(),
            "<Control-Key-S>": lambda e: self.ejecutar_atajo_guardar(),
            "<Control-Key-e>": lambda e: self.ejecutar_atajo_exportar(),
            "<Control-Key-E>": lambda e: self.ejecutar_atajo_exportar(),
        }
        for atajo, cmd in atajos.items():
            try:
                self.bind(atajo, cmd)
            except Exception:
                pass

    def ejecutar_atajo_guardar(self):
        # Encontrar frame activo (Tarea 17)
        for class_name, f in self._frames.items():
            try:
                if f.winfo_exists() and f.winfo_viewable():
                    if class_name == "NotasAsistenciaFrame":
                        active_tab = f.tabview.get()
                        if active_tab == "Notas":
                            f.frame_notas.guardar_notas()
                        elif active_tab == "Asistencia":
                            f.frame_asistencia.guardar_asistencia()
                    elif class_name == "ObservacionesFrame":
                        f.guardar_observacion()
                    elif class_name == "HabitosFrame":
                        f.guardar_habitos()
            except Exception as e:
                print(f"[!] Error ejecutando guardar por atajo: {e}")

    def ejecutar_atajo_exportar(self):
        # Enrutar a reportes (Tarea 17)
        self.mostrar_reportes()

    def _iniciar_inactividad_check(self):
        if self._destroyed:
            return
        try:
            from anti_analysis import get_inactivity_time
            inactivo_segundos = get_inactivity_time()
            # Si supera 300 segundos (5 minutos) y no está ya bloqueada, bloquear la app
            if inactivo_segundos >= 300 and not getattr(self, "_bloqueado", False):
                self._bloquear_aplicacion()
        except Exception:
            pass
        self.after(5000, self._iniciar_inactividad_check)

    def _bloquear_aplicacion(self):
        self._bloqueado = True
        
        # Capturar todos los eventos de teclado para deshabilitar atajos y navegación
        def bloquear_evento(e):
            return "break"
        self.bind_all("<Key>", bloquear_evento)
        
        # Crear frame de bloqueo superpuesto a pantalla completa
        self.frame_bloqueo = ctk.CTkFrame(self, fg_color="#0A141D")
        self.frame_bloqueo.place(relx=0, rely=0, relwidth=1, relheight=1)
        
        from theme import C, FONT_TITLE, FONT_BODY
        # Elementos del frame de bloqueo
        lbl_titulo = ctk.CTkLabel(self.frame_bloqueo, text="🔒 APLICACIÓN BLOQUEADA POR INACTIVIDAD", font=(FONT_TITLE, 20, "bold"), text_color=C["rojo"])
        lbl_titulo.pack(pady=(180, 20))
        
        lbl_instrucciones = ctk.CTkLabel(self.frame_bloqueo, text="Para reanudar su sesión de forma segura, ingrese su Cédula:", font=(FONT_BODY, 13), text_color=C["texto_sec"])
        lbl_instrucciones.pack(pady=10)
        
        self.entry_cedula_lock = ctk.CTkEntry(self.frame_bloqueo, show="*", width=250, placeholder_text="Ingrese su Cédula")
        self.entry_cedula_lock.pack(pady=10)
        self.entry_cedula_lock.focus_force()
        
        # Permitir presionar Enter para desbloquear
        self.entry_cedula_lock.bind("<Return>", lambda e: self._intentar_desbloquear())
        
        btn_desbloquear = ctk.CTkButton(self.frame_bloqueo, text="🔓 Desbloquear", fg_color=C["cian"], hover_color=C["verde"], command=self._intentar_desbloquear)
        btn_desbloquear.pack(pady=10)
        
        self.lbl_error_lock = ctk.CTkLabel(self.frame_bloqueo, text="", text_color=C["rojo"], font=(FONT_BODY, 12, "bold"))
        self.lbl_error_lock.pack(pady=5)

    def _intentar_desbloquear(self):
        cedula_ingresada = self.entry_cedula_lock.get().strip()
        
        # Cargar cédula configurada
        from rdsecurity import cargar_config_segura
        cfg = cargar_config_segura({})
        cedula_correcta = cfg.get("docente_cedula", "").strip()
        
        import re
        norm_ingresada = re.sub(r'[^a-zA-Z0-9]', '', cedula_ingresada).upper()
        norm_correcta = re.sub(r'[^a-zA-Z0-9]', '', cedula_correcta).upper()
        
        if not cedula_correcta or norm_ingresada == norm_correcta:
            # Desbloquear
            self._bloqueado = False
            self.unbind_all("<Key>")
            self._registrar_atajos() # Restaurar atajos
            self.frame_bloqueo.destroy()
        else:
            self.lbl_error_lock.configure(text="Cédula incorrecta. Intente de nuevo.")
            self.entry_cedula_lock.delete(0, "end")

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
            # 1. Respaldar base de datos SQLite cifrada (Tarea 15)
            if hasattr(self.engine, "db_manager"):
                self.engine.db_manager.respaldar_base_datos()
        except Exception:
            pass

        try:
            ruta_original = self.engine.ruta
            if not isinstance(ruta_original, str) or not os.path.exists(ruta_original):
                return
            dir_base = os.path.dirname(os.path.abspath(ruta_original))
            dir_respaldos = os.path.join(dir_base, "Respaldos_Auto")
            if not os.path.exists(dir_respaldos):
                os.makedirs(dir_respaldos)
            ahora = datetime.datetime.now().strftime("%Y-%m-%d_%H-%M")
            nombre_base = os.path.basename(ruta_original).replace(".xlsx", "")
            nombre = f"{nombre_base}_auto_{ahora}.xlsx"
            shutil.copy2(ruta_original, os.path.join(dir_respaldos, nombre))
            # Limpiar: máximo 3 auto-backups
            respaldos = sorted(
                [os.path.join(dir_respaldos, f) for f in os.listdir(dir_respaldos) if f.endswith(".xlsx")],
                key=os.path.getmtime
            )
            while len(respaldos) > 3:
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
        # Forzar guardado cifrado final antes de cerrar (flush del debounce pendiente)
        try:
            if hasattr(self, "engine") and hasattr(self.engine, "flush_save_sync"):
                self.engine.flush_save_sync()
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

    def _deshabilitar_animaciones_dwm(self):
        """Desactiva las transiciones (animaciones) de DWM SOLO para esta ventana.

        En Windows, al maximizar/minimizar/restaurar, el compositor (DWM) anima el
        cambio de tamaño cruzando la imagen ANTERIOR con la NUEVA. Como
        CustomTkinter repinta sus canvas relativamente lento, durante esa animación
        se ven ambas imágenes a la vez → el efecto de doble-imagen / ghosting.

        Al desactivar la transición por-ventana, el cambio de tamaño se aplica de
        forma limpia e instantánea, sin el fantasma. Es una API nativa de Win32
        (no pelea con el sistema ni usa alpha/overlays)."""
        if sys.platform != "win32":
            return
        try:
            hwnd = ctypes.windll.user32.GetParent(self.winfo_id())
            if not hwnd:
                hwnd = self.winfo_id()
            DWMWA_TRANSITIONS_FORCEDISABLED = 3
            valor = ctypes.c_int(1)  # 1 = desactivar animaciones
            ctypes.windll.dwmapi.DwmSetWindowAttribute(
                hwnd,
                DWMWA_TRANSITIONS_FORCEDISABLED,
                ctypes.byref(valor),
                ctypes.sizeof(valor),
            )
        except Exception:
            pass

    def _on_window_unmap(self, event):
        if event.widget == self:
            self._is_minimized = True

    def _on_window_map(self, event):
        # Restaurar desde minimizado: NO tocamos alpha ni forzamos update() aquí.
        # Esos "trucos" se ejecutaban DURANTE la animación de restauración de
        # Windows/DWM y provocaban la doble-imagen / ghosting. Dejamos que el
        # sistema operativo anime la ventana de forma nativa.
        if event.widget == self:
            self._is_minimized = False

    def _on_window_configure(self, event):
        # Solo rastreamos el tamaño actual. El redimensionado (arrastrar bordes,
        # maximizar y restaurar) lo maneja Windows de forma NATIVA: no dibujamos
        # overlay ni forzamos redibujados, que era lo que causaba el parpadeo y la
        # doble-imagen. El fondo navy de la ventana raíz cubre cualquier zona aún
        # no repintada durante la transición.
        if event.widget == self:
            try:
                if self.state() != "iconic":
                    self._last_width = event.width
                    self._last_height = event.height
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
        try:
            from setup import SetupWizard
            wizard = SetupWizard()
            wizard.mainloop()
            if os.path.exists(CONFIG_FILE):
                iniciar_programa_principal()
        except Exception as e:
            cfg = {
                "modalidad": "premedia",
                "docente_nombre": "Docente de Prueba",
                "docente_cedula": "8-000-0000",
                "escuela_nombre": "Escuela de Prueba",
                "escuela_region": "Comarca Ngäbe Buglé",
                "ano_lectivo": "2026"
            }
            guardar_config_segura(cfg)
            iniciar_programa_principal()
    else:
        iniciar_programa_principal()
