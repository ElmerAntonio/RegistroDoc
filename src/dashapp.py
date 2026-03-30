"""
RegistroDoc Pro — Dashboard Principal
Diseño basado en la imagen de referencia:
  - Header bar superior con búsqueda + avatar + notificaciones
  - Sidebar plegable con iconos
  - Panel principal con métricas, gráfica de línea y barra
  - Marca de agua sutil de Panamá
  - Ventana fija 1280x720 (no redimensionable)
  - Paleta: fondo #0A1628, acentos cian #00E5FF / verde #00FF88
"""

import customtkinter as ctk
import tkinter as tk
from tkinter import font as tkfont
import os
import datetime

# Matplotlib para las gráficas
import matplotlib
matplotlib.use("TkAgg")
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.figure import Figure

# PIL para el logo
try:
    from PIL import Image, ImageTk, ImageFilter, ImageDraw
    PIL_OK = True
except ImportError:
    PIL_OK = False

# ─── PALETA PRINCIPAL CORREGIDA ──────────────────────────────────────────────
C = {
    "fondo":        "#0A1628",
    "sidebar":      "#0D1F35",
    "header":       "#0D1F35",
    "card":         "#0F2744",
    "card_borde":   "#00E5FF",
    "cian":         "#00E5FF",
    "verde":        "#00FF88",
    "rojo":         "#FF4444",
    "amarillo":     "#FFD700",
    "purpura":      "#A855F7",
    "texto":        "#E2E8F0",
    "texto_sec":    "#64748B",
    "texto_dim":    "#334155",
    "input":        "#1E3A5F",
    "hover":        "#1A3352",
    "activo":       "#1A3352",  # Corregido: Hex de 6 dígitos
    "borde":        "#1E3A5F",
    "glow":         "#1E3A5F",  # Corregido: Hex de 6 dígitos
}

# Color de acento actual (cambiable)
ACENTO = C["cian"]


# ─── CLASE PRINCIPAL DEL DASHBOARD ─────────────────────────────────────────
class DashboardFrame(ctk.CTkFrame):
    """
    Panel dashboard completo.
    Parámetros:
        master      : contenedor padre
        engine      : DataEngine con los métodos del motor de datos
        app_principal: referencia a RegistroDocApp para navegar
    """

    def __init__(self, master, engine, app_principal=None, **kwargs):
        super().__init__(master, fg_color=C["fondo"], **kwargs)
        self.engine        = engine
        self.app           = app_principal
        self._acento       = ACENTO
        self._busqueda_vis = False
        self._perfil_menu  = None

        # Stats una sola vez
        self._stats = self._cargar_stats()

        self._construir()

    # ══════════════════════════════════════════════════════════════════════
    #  DATOS
    # ══════════════════════════════════════════════════════════════════════
    def _cargar_stats(self) -> dict:
        try:
            return self.engine.get_dashboard_stats()
        except Exception:
            return {
                "total": 92, "riesgo": 7, "honor": "SANTOS FIDEL",
                "honor_prom": "4.9", "asistencia": "98%",
                "grados": ["7°", "8°", "9°"],
            }

    # ══════════════════════════════════════════════════════════════════════
    #  CONSTRUCCIÓN PRINCIPAL
    # ══════════════════════════════════════════════════════════════════════
    def _construir(self):
        self.pack_propagate(False)
        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=0)   # header
        self.rowconfigure(1, weight=1)   # cuerpo

        # 1. Header
        self._header()
        # 2. Separador
        ctk.CTkFrame(self, height=1, fg_color=C["borde"]).grid(
            row=1, column=0, sticky="ew")
        # 3. Cuerpo (sidebar + panel)
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

        # Ícono mensajes
        ctk.CTkButton(right, text="✉", width=36, height=36,
                      fg_color="transparent", hover_color=C["hover"],
                      font=ctk.CTkFont(size=16), text_color=C["texto_sec"],
                      corner_radius=18, command=lambda: None).pack(
            side="left", padx=3)

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
        av_path = os.path.join(os.path.dirname(__file__), "..", "img", "avatar.png")
        self._avatar_btn = ctk.CTkButton(
            right, text="ET", width=36, height=36,
            fg_color=self._acento, hover_color=C["hover"],
            font=ctk.CTkFont("Segoe UI", 13, "bold"),
            text_color=C["fondo"], corner_radius=18,
            command=self._toggle_perfil)
        if PIL_OK and os.path.exists(av_path):
            try:
                pil_av = Image.open(av_path).resize((36, 36))
                self._av_img = ctk.CTkImage(pil_av, size=(36, 36))
                self._avatar_btn.configure(image=self._av_img, text="")
            except Exception:
                pass
        self._avatar_btn.pack(side="left", padx=(6, 0))

    # ══════════════════════════════════════════════════════════════════════
    #  BÚSQUEDA INTELIGENTE
    # ══════════════════════════════════════════════════════════════════════
    def _buscar_estudiante(self, *_):
        texto = self._search_var.get().strip()
        for w in self._resultados_frame.winfo_children():
            w.destroy()

        if len(texto) < 2:
            self._resultados_frame.grid_remove()
            return

        # Buscar en el engine
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

        # Posición bajo el avatar
        x = self.winfo_rootx() + self.winfo_width() - 270
        y = self.winfo_rooty() + 60
        self._perfil_menu.geometry(f"250x320+{x}+{y}")

        # Borde
        ctk.CTkFrame(self._perfil_menu, fg_color=self._acento,
                     height=2).pack(fill="x")

        # Info del docente
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

        # Opciones
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

        # Color de acento
        ctk.CTkLabel(self._perfil_menu, text="Color de acento:",
                     font=ctk.CTkFont(size=11), text_color=C["texto_sec"]).pack(
            anchor="w", padx=16)

        colores_f = ctk.CTkFrame(self._perfil_menu, fg_color="transparent")
        colores_f.pack(fill="x", padx=16, pady=6)
        colores = [
            ("Cian",       "#00E5FF"),
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
                      command=lambda: None).pack(fill="x", padx=8)

        # Cerrar al hacer clic fuera
        self._perfil_menu.bind("<FocusOut>",
                               lambda e: self._cerrar_perfil_menu())

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
        # Redibujar (simple: solo actualizar el título "Pro")
        # Para efecto completo se requeriría reconstruir toda la UI

    def _toggle_notificaciones(self):
        pass  # Placeholder para notificaciones

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

        # Sidebar
        self._sidebar_widget(cuerpo)
        # Panel principal
        self._panel_principal(cuerpo)

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

        ancho = 200 if self._sidebar_exp else 56

        # Botón "Software" desplegable superior
        if self._sidebar_exp:
            top_btn = ctk.CTkButton(
                self._sb, text="  Software  ▾",
                fg_color=C["hover"], hover_color=C["borde"],
                font=ctk.CTkFont("Segoe UI", 13),
                text_color=C["texto_sec"],
                anchor="w", height=38, corner_radius=6,
                command=lambda: None)
            top_btn.pack(fill="x", padx=8, pady=(12, 8))
        else:
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
                border_color=self._acento, # <--- SOLUCIÓN: Color fijo, el grosor 0 lo oculta
                command=cmd)
            btn.pack(fill="x", padx=6, pady=2)

        # Espacio flexible
        ctk.CTkFrame(self._sb, fg_color="transparent").pack(
            fill="both", expand=True)

        # Versión al fondo
        if self._sidebar_exp:
            ctk.CTkLabel(self._sb, text="v.Prov.22:6",
                         font=ctk.CTkFont("Segoe UI", 10),
                         text_color=C["texto_dim"]).pack(pady=(0, 6))

        # Toggle
        toggle_txt = "«" if self._sidebar_exp else "»"
        ctk.CTkButton(self._sb, text=toggle_txt, width=36, height=28,
                      fg_color=C["hover"], hover_color=C["borde"],
                      font=ctk.CTkFont(size=14, weight="bold"),
                      text_color=C["texto_sec"],
                      corner_radius=6,
                      command=self._toggle_sidebar_dash).pack(
            pady=(4, 10), padx=8)

    def _toggle_sidebar_dash(self):
        self._sidebar_exp = not self._sidebar_exp
        nuevo_ancho = 200 if self._sidebar_exp else 56
        self._sb.configure(width=nuevo_ancho)
        self._sb_renderizar()

    def _ir_inicio(self):
        pass  # ya estamos aquí

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

    # ══════════════════════════════════════════════════════════════════════
    #  PANEL PRINCIPAL
    # ══════════════════════════════════════════════════════════════════════
    def _panel_principal(self, parent):
        panel = ctk.CTkScrollableFrame(parent, fg_color=C["fondo"],
                                       corner_radius=0,
                                       scrollbar_fg_color=C["fondo"],
                                       scrollbar_button_color=C["borde"])
        panel.grid(row=0, column=1, sticky="nsew", padx=0, pady=0)
        panel.columnconfigure(0, weight=1)

        # Marca de agua Panamá (canvas con texto muy tenue)
        self._marca_agua(panel)

        # Título y fecha
        self._titulo_seccion(panel)

        # Tarjetas métricas
        self._tarjetas_metricas(panel)

        # Gráficas
        self._graficas(panel)

        # Grados activos y accesos rápidos
        self._footer_panel(panel)

    # ── Marca de agua ──────────────────────────────────────────────────
    def _marca_agua(self, parent):
        """Canvas con texto PANAMÁ muy tenue como marca de agua."""
        cv = tk.Canvas(parent, bg=C["fondo"], height=0,
                       highlightthickness=0, bd=0)
        cv.pack(fill="x")
        # Se dibuja sobre toda la pantalla con place desde el panel principal
        # (técnica: se pone un label gigante transparente atrás)
        fondo_lbl = ctk.CTkLabel(
            parent,
            text="P  A  N  A  M  Á",
            font=ctk.CTkFont("Segoe UI", 72, "bold"),
            text_color="#0D2040",        # casi invisible sobre fondo
            fg_color="transparent")
        fondo_lbl.pack(pady=(0, 0))
        # No ocupa espacio visible — es un efecto decorativo

    # ── Título ─────────────────────────────────────────────────────────
    def _titulo_seccion(self, parent):
        f = ctk.CTkFrame(parent, fg_color="transparent")
        f.pack(fill="x", padx=24, pady=(10, 4))
        ctk.CTkLabel(f, text="Panel de Control Principal",
                     font=ctk.CTkFont("Segoe UI", 22, "bold"),
                     text_color=C["texto"]).pack(side="left")
        fecha = datetime.datetime.now().strftime("%A %d de %B, %Y").capitalize()
        ctk.CTkLabel(f, text=fecha,
                     font=ctk.CTkFont("Segoe UI", 11),
                     text_color=C["texto_sec"]).pack(side="right")

    # ── Tarjetas métricas ────────────────────────────────────────────────
    def _tarjetas_metricas(self, parent):
        s = self._stats
        total    = str(s.get("total", 0))
        riesgo   = str(s.get("riesgo", 0))
        honor    = str(s.get("honor", "—"))
        asist    = str(s.get("asistencia", "0%"))

        frame = ctk.CTkFrame(parent, fg_color="transparent")
        frame.pack(fill="x", padx=20, pady=(6, 10))
        frame.columnconfigure((0, 1, 2, 3), weight=1, uniform="card")

        datos = [
            ("👥",  "Total Alumnos",       total,  "Matriculados",             self._acento),
            ("⚠️",  "Alumnos en Riesgo",   riesgo, "Nota promedio < 3.0",      "#EF4444"),
            ("🏆",  "Cuadro de Honor",     "1",    honor[:18],                 C["amarillo"]),
            ("📊",  "Asistencia",          asist,  "Promedio general",         C["verde"]),
        ]

        for col, (ico, titulo, valor, sub, color) in enumerate(datos):
            self._tarjeta(frame, col, ico, titulo, valor, sub, color)

    def _tarjeta(self, parent, col, ico, titulo, valor, sub, color):
        card = ctk.CTkFrame(parent, fg_color=C["card"],
                            border_width=1, border_color=color,
                            corner_radius=12)
        card.grid(row=0, column=col, padx=6, pady=4, sticky="nsew")

        # Icono
        ctk.CTkLabel(card, text=ico, font=ctk.CTkFont(size=20),
                     fg_color="transparent",
                     text_color=color).pack(anchor="w", padx=16, pady=(14, 0))

        # Título (con wraplength para que no desborde)
        ctk.CTkLabel(card, text=titulo,
                     font=ctk.CTkFont("Segoe UI", 12),
                     text_color=C["texto_sec"],
                     wraplength=140,
                     justify="left").pack(anchor="w", padx=16, pady=(2, 0))

        # Valor grande
        ctk.CTkLabel(card, text=valor,
                     font=ctk.CTkFont("Segoe UI", 32, "bold"),
                     text_color=color,
                     wraplength=160,
                     justify="left").pack(anchor="w", padx=16, pady=(0, 2))

        # Subtítulo
        ctk.CTkLabel(card, text=sub,
                     font=ctk.CTkFont("Segoe UI", 11),
                     text_color=C["texto_sec"],
                     wraplength=150,
                     justify="left").pack(anchor="w", padx=16, pady=(0, 14))

    # ── Gráficas ─────────────────────────────────────────────────────────
    def _graficas(self, parent):
        gf = ctk.CTkFrame(parent, fg_color="transparent")
        gf.pack(fill="x", padx=20, pady=6)
        gf.columnconfigure(0, weight=6)
        gf.columnconfigure(1, weight=4)

        self._grafica_linea(gf)
        self._grafica_barras(gf)

    def _grafica_linea(self, parent):
        """Gráfica de escalada — evolución de notas T1→T2→T3."""
        card = ctk.CTkFrame(parent, fg_color=C["card"],
                            border_width=1, border_color=C["borde"],
                            corner_radius=12)
        card.grid(row=0, column=0, padx=(0, 8), sticky="nsew")

        ctk.CTkLabel(card, text="Tendencia de Rendimiento Académico  (X/Y)",
                     font=ctk.CTkFont("Segoe UI", 13, "bold"),
                     text_color=C["texto"]).pack(anchor="w", padx=16, pady=(12, 0))
        ctk.CTkLabel(card, text="Escala 1–5 por trimestre",
                     font=ctk.CTkFont("Segoe UI", 10),
                     text_color=C["texto_sec"]).pack(anchor="w", padx=16)

        # Matplotlib
        fig = Figure(figsize=(5.4, 2.8), dpi=96,
                     facecolor=C["card"])
        ax = fig.add_subplot(111)
        ax.set_facecolor(C["card"])
        fig.subplots_adjust(left=0.1, right=0.97, top=0.92, bottom=0.16)

        x = [0, 1, 2, 3, 4, 5, 6, 7]
        y = [2.0, 2.5, 3.1, 3.0, 3.4, 3.7, 4.1, 4.8]

        ax.plot(x, y, color=self._acento, linewidth=2.2,
                zorder=3, solid_capstyle="round")
        # Relleno bajo la línea
        ax.fill_between(x, y, alpha=0.12, color=self._acento)
        # Nodos circulares brillantes
        ax.scatter(x, y, color=self._acento, s=55, zorder=5,
                   edgecolors=C["fondo"], linewidths=1.5)

        # Ejes
        ax.set_xlim(-0.3, 7.3)
        ax.set_ylim(1, 5.2)
        ax.set_xlabel("X  (Trimestres)", color=C["texto_sec"], fontsize=9)
        ax.set_ylabel("Y  (Notas 1–5)", color=C["texto_sec"], fontsize=9)
        ax.tick_params(colors=C["texto_sec"], labelsize=8)
        for spine in ax.spines.values():
            spine.set_edgecolor(C["borde"])
        ax.grid(True, color=C["borde"], linewidth=0.5, alpha=0.6)
        ax.set_xticks(x)

        canvas = FigureCanvasTkAgg(fig, master=card)
        canvas.draw()
        canvas.get_tk_widget().pack(fill="x", padx=10, pady=(4, 12))

    def _grafica_barras(self, parent):
        """Barras de asistencia / rendimiento por grupo."""
        card = ctk.CTkFrame(parent, fg_color=C["card"],
                            border_width=1, border_color=C["borde"],
                            corner_radius=12)
        card.grid(row=0, column=1, sticky="nsew")

        ctk.CTkLabel(card, text="Rendimiento por Grupo",
                     font=ctk.CTkFont("Segoe UI", 13, "bold"),
                     text_color=C["texto"]).pack(anchor="w", padx=16, pady=(12, 0))
        ctk.CTkLabel(card, text="Promedio final por sección",
                     font=ctk.CTkFont("Segoe UI", 10),
                     text_color=C["texto_sec"]).pack(anchor="w", padx=16)

        fig2 = Figure(figsize=(3.2, 2.8), dpi=96,
                      facecolor=C["card"])
        ax2 = fig2.add_subplot(111)
        ax2.set_facecolor(C["card"])
        fig2.subplots_adjust(left=0.14, right=0.97, top=0.9, bottom=0.18)

        grupos   = ["7°", "8°", "9°", "7°B", "8°B"]
        valores  = [3.8, 3.5, 4.2, 3.1, 3.7]
        colores  = [self._acento] * len(grupos)

        bars = ax2.bar(grupos, valores, color=colores, width=0.6,
                       edgecolor=C["fondo"], linewidth=0.8)

        ax2.set_ylim(1, 5.3)
        ax2.tick_params(colors=C["texto_sec"], labelsize=8)
        for spine in ax2.spines.values():
            spine.set_edgecolor(C["borde"])
        ax2.grid(True, axis="y", color=C["borde"],
                 linewidth=0.5, alpha=0.5)
        ax2.set_ylabel("Promedio", color=C["texto_sec"], fontsize=8)

        canvas2 = FigureCanvasTkAgg(fig2, master=card)
        canvas2.draw()
        canvas2.get_tk_widget().pack(fill="x", padx=10, pady=(4, 12))

    # ── Footer: grados activos + accesos rápidos ─────────────────────────
    def _footer_panel(self, parent):
        ff = ctk.CTkFrame(parent, fg_color="transparent")
        ff.pack(fill="x", padx=20, pady=(4, 20))
        ff.columnconfigure((0, 1), weight=1, uniform="bot")

        # Grados activos
        gc = ctk.CTkFrame(ff, fg_color=C["card"],
                          border_width=1, border_color=C["borde"],
                          corner_radius=12)
        gc.grid(row=0, column=0, padx=(0, 8), sticky="nsew", pady=4)

        ctk.CTkLabel(gc, text="Grados Activos",
                     font=ctk.CTkFont("Segoe UI", 14, "bold"),
                     text_color=C["texto"]).pack(anchor="w", padx=16, pady=(14, 8))

        try:
            grados = self.engine.obtener_grados_activos()
        except Exception:
            grados = ["7°", "8°", "9°"]

        gf = ctk.CTkFrame(gc, fg_color="transparent")
        gf.pack(padx=16, pady=(0, 14))
        for g in grados:
            pill = ctk.CTkLabel(
                gf, text=f"  {g}  ",
                fg_color=C["input"],
                corner_radius=16,
                font=ctk.CTkFont("Segoe UI", 14, "bold"),
                text_color=self._acento,
                padx=12, pady=6)
            pill.pack(side="left", padx=4)

        # Accesos rápidos
        ac = ctk.CTkFrame(ff, fg_color=C["card"],
                          border_width=1, border_color=C["borde"],
                          corner_radius=12)
        ac.grid(row=0, column=1, sticky="nsew", pady=4)

        ctk.CTkLabel(ac, text="Accesos Rápidos",
                     font=ctk.CTkFont("Segoe UI", 14, "bold"),
                     text_color=C["texto"]).pack(anchor="w", padx=16, pady=(14, 8))

        af = ctk.CTkFrame(ac, fg_color="transparent")
        af.pack(fill="x", padx=16, pady=(0, 14))
        af.columnconfigure((0, 1, 2), weight=1)

        botones = [
            ("📝 Notas",      "#2563EB", self._ir_notas),
            ("📅 Asistencia", "#059669", self._ir_asistencia),
            ("⚙️ Config",     "#7C3AED", self._abrir_configuracion),
        ]
        for i, (txt, col, cmd) in enumerate(botones):
            ctk.CTkButton(af, text=txt, height=48,
                          font=ctk.CTkFont("Segoe UI", 13, "bold"),
                          fg_color=col, hover_color=col,
                          corner_radius=10,
                          command=cmd).grid(
                row=0, column=i, padx=5, sticky="ew")


# ─── MOCK PARA PRUEBA STANDALONE ─────────────────────────────────────────────
if __name__ == "__main__":
    ctk.set_appearance_mode("dark")
    ctk.set_default_color_theme("blue")

    class MockEngine:
        modalidad = "premedia"
        def get_dashboard_stats(self):
            return {"total": 92, "riesgo": 7,
                    "honor": "SANTOS FIDEL", "honor_prom": "4.9",
                    "asistencia": "98%"}
        def obtener_grados_activos(self):
            return ["7°", "8°", "9°"]
        def obtener_estudiantes_completos(self, g):
            names = ["Maria Gonzalez", "Pedro Pinto", "Ana Santos"]
            return [{"id": i+1, "nombre": n, "cedula": ""} for i, n in enumerate(names)]
        def obtener_datos_generales(self):
            return {"docente_nombre": "Prof. Elmer Tugri",
                    "correo": "elmer.tugri7@meduca.edu.pa"}

    root = ctk.CTk()
    root.title("RegistroDoc Pro v3.0")
    # ── VENTANA FIJA 1280×720 ──────────────────────────────────────────
    root.geometry("1280x720")
    root.minsize(1280, 720)
    root.maxsize(1280, 720)
    root.resizable(False, False)

    DashboardFrame(root, MockEngine()).pack(fill="both", expand=True)
    root.mainloop()