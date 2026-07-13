import os
import shutil
import sqlite3
import tkinter as tk
from tkinter import messagebox, filedialog
import customtkinter as ctk
from PIL import Image, ImageTk
from config import BASE_DIR, ASSETS_DIR
from rdsecurity import guardar_config_segura
from rdsql import SQLDatabaseManager, obtener_dir_datos_usuario
from rddata import DataEngine
from utils.translator import tr, set_current_lang

# ══════════════════════════════════════════════════════════════
#  DISEÑO PREMIUM — PALETA DE COLORES (COHERENTE CON DASHBOARD)
# ══════════════════════════════════════════════════════════════
C = {
    "fondo": "#040D1A",
    "card": "#091524",
    "borde": "#1E2E42",
    "texto": "#F8FAFC",
    "texto_sec": "#94A3B8",
    "azul": "#3B82F6",
    "cian": "#06B6D4",
    "verde": "#10B981",
    "hover": "#2563EB",
    "rojo": "#EF4444",
}

class SetupWizard(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("RegistroDoc Pro - Asistente de Configuración Inicial")
        
        # Configurar AppUserModelID para icono en barra de tareas
        import sys
        if sys.platform == "win32":
            try:
                import ctypes
                ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("registrodoc.pro.v3")
            except Exception:
                pass

        # Cargar iconos
        icon_path = os.path.join(ASSETS_DIR, "icon_fixed.ico")
        png_path = os.path.join(ASSETS_DIR, "icono.png")
        if os.path.exists(icon_path):
            try:
                self.iconbitmap(icon_path)
            except Exception:
                pass
        if os.path.exists(png_path):
            try:
                pil = Image.open(png_path).resize((64, 64))
                self._icono_app = ImageTk.PhotoImage(pil)
                self.iconphoto(True, self._icono_app)
            except Exception:
                pass
        
        # Tamaño de la ventana
        width = 850
        height = 680
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        x = (screen_width - width) // 2
        y = (screen_height - height) // 2
        self.geometry(f"{width}x{height}+{x}+{y}")
        self.configure(fg_color=C["fondo"])
        self.resizable(False, False)

        # Datos acumulativos del Wizard
        self.datos = {
            "docente_nombre": "",
            "docente_cedula": "",
            "escuela_nombre": "",
            "escuela_region": "Comarca Ngäbe Buglé",
            "ano_lectivo": "2026",
            "modalidad": "premedia",
            "logo_escuela_path": "",
            "ruta_exportacion": os.path.join(os.path.expanduser("~"), "Documents", "RegistroDoc"),
            "idioma": "es",
            "jornada": "Matutina",
            "director_nombre": "",
            "subdirector_nombre": "",
            "grados_seleccionados": {},  # {"7°": "A", "8°": "B"}
            "materias_mapeadas": {},     # {"Español": ["7° A", "8° B"]}
            "horario_slots": []          # [{"dia": "lunes", "bloque": "...", "materia_grado": "..."}]
        }
        self.logo_preview_img = None
        self.paso_actual = 1

        self.crear_interfaz()

    def crear_interfaz(self):
        # Frame superior para el header de progreso
        self.progress_frame = ctk.CTkFrame(self, fg_color=C["card"], height=70, corner_radius=0, border_width=1, border_color=C["borde"])
        self.progress_frame.pack(fill="x", side="top")
        self.dibujar_progreso()

        # Contenedor principal
        self.main_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.main_frame.pack(fill="both", expand=True, padx=40, pady=25)

        # Frame de botones inferiores
        self.nav_frame = ctk.CTkFrame(self, fg_color="transparent", height=60)
        self.nav_frame.pack(fill="x", side="bottom", padx=40, pady=20)

        self.btn_atras = ctk.CTkButton(self.nav_frame, text=tr("⬅️ Atrás"), font=("Segoe UI", 12, "bold"), fg_color="#334155", hover_color="#475569", width=120, command=self.retroceder)
        self.btn_atras.pack(side="left")

        self.btn_omitir = ctk.CTkButton(self.nav_frame, text=tr("Omitir y Usar Demo ❌"), font=("Segoe UI", 11), fg_color="#1E293B", hover_color="#334155", text_color=C["texto_sec"], width=150, command=self.omitir_asistente)
        self.btn_omitir.pack(side="left", padx=20)

        self.btn_siguiente = ctk.CTkButton(self.nav_frame, text=tr("Siguiente ➡️"), font=("Segoe UI", 12, "bold"), fg_color=C["azul"], hover_color=C["hover"], width=140, command=self.avanzar)
        self.btn_siguiente.pack(side="right")

        self.mostrar_paso()

    def dibujar_progreso(self):
        for w in self.progress_frame.winfo_children():
            w.destroy()
        
        # Título del asistente
        lbl_title = ctk.CTkLabel(self.progress_frame, text="RegistroDoc Pro Setup", font=("Segoe UI", 15, "bold"), text_color=C["cian"])
        lbl_title.pack(side="left", padx=25)

        steps = [tr("Perfil"), tr("Grados"), tr("Materias"), tr("Horario"), tr("Finalizar")]
        steps_frame = ctk.CTkFrame(self.progress_frame, fg_color="transparent")
        steps_frame.pack(side="right", padx=25)

        for i, name in enumerate(steps, 1):
            color = C["verde"] if i < self.paso_actual else (C["cian"] if i == self.paso_actual else C["texto_sec"])
            bold_font = ("Segoe UI", 11, "bold") if i == self.paso_actual else ("Segoe UI", 10)
            lbl = ctk.CTkLabel(steps_frame, text=f" {i}. {name} ", font=bold_font, text_color=color)
            lbl.pack(side="left", padx=5)
            if i < len(steps):
                ctk.CTkLabel(steps_frame, text="•", text_color=C["borde"]).pack(side="left")


    def limpiar_main_frame(self):
        for w in self.main_frame.winfo_children():
            w.destroy()

    def mostrar_paso(self):
        self.dibujar_progreso()
        self.limpiar_main_frame()
        self.btn_atras.configure(state="normal" if self.paso_actual > 1 else "disabled")
        
        if self.paso_actual == 1:
            self.mostrar_paso_1()
        elif self.paso_actual == 2:
            self.mostrar_paso_2()
        elif self.paso_actual == 3:
            self.mostrar_paso_3()
        elif self.paso_actual == 4:
            self.mostrar_paso_4()
        elif self.paso_actual == 5:
            self.mostrar_paso_5()

    def avanzar(self):
        if self.paso_actual == 1:
            if not self.guardar_paso_1(): return
        elif self.paso_actual == 2:
            if not self.guardar_paso_2(): return
        elif self.paso_actual == 3:
            if not self.guardar_paso_3(): return
        elif self.paso_actual == 4:
            if not self.guardar_paso_4(): return
        
        self.paso_actual += 1
        self.mostrar_paso()

    def retroceder(self):
        if self.paso_actual > 1:
            self.paso_actual -= 1
            self.mostrar_paso()

    def omitir_asistente(self):
        confirm = messagebox.askyesno(
            "Omitir Configuración",
            "¿Deseas omitir el asistente e inicializar el programa con un perfil y datos de demostración?\n(Podrás cambiar estos datos más tarde)."
        )
        if confirm:
            demo_cfg = {
                "modalidad": "premedia",
                "docente_nombre": "Docente de Prueba",
                "docente_cedula": "8-000-0000",
                "escuela_nombre": "Escuela de Prueba",
                "escuela_region": "Comarca Ngäbe Buglé",
                "ano_lectivo": "2026",
                "logo_escuela_path": "",
                "ruta_exportacion": os.path.join(os.path.expanduser("~"), "Documents", "RegistroDoc"),
                "idioma": "es",
                "jornada": "Matutina",
                "director_nombre": "",
                "subdirector_nombre": ""
            }
            guardar_config_segura(demo_cfg)
            
            # Inicializar base de datos vacía
            db = SQLDatabaseManager()
            conn = db.conectar()
            cursor = conn.cursor()
            for k, v in demo_cfg.items():
                cursor.execute("INSERT OR REPLACE INTO configuracion (clave, valor) VALUES (?, ?);", (k, v))
            
            # Agregar grado de prueba
            cursor.execute("INSERT OR IGNORE INTO grados (nombre, seccion, modalidad) VALUES (?, ?, ?);", ("7", "A", "premedia"))
            conn.commit()
            db.cerrar()
            
            # Copiar plantilla
            from config import ASSETS_DIR
            plantilla = os.path.join(ASSETS_DIR, "templates", "Registro_Premedia.xlsx")
            user_dir = obtener_dir_datos_usuario()
            active_excel = os.path.join(user_dir, "temp", "Registro_Premedia.xlsx")
            os.makedirs(os.path.dirname(active_excel), exist_ok=True)
            if os.path.exists(plantilla) and not os.path.exists(active_excel):
                shutil.copy2(plantilla, active_excel)
                
            self.destroy()

    # ══════════════════════════════════════════════════════════════
    #  PASO 1: DATOS DOCENTE Y LOGO ESCOLAR
    # ══════════════════════════════════════════════════════════════
    def mostrar_paso_1(self):
        # Contenedor de dos columnas
        self.main_frame.columnconfigure(0, weight=6)
        self.main_frame.columnconfigure(1, weight=4)

        # Columna Izquierda: Scrollable Frame para acomodar los nuevos campos cómodamente
        col_izq = ctk.CTkScrollableFrame(self.main_frame, fg_color="transparent")
        col_izq.grid(row=0, column=0, sticky="nsew", padx=(0, 20))

        col_der = ctk.CTkFrame(self.main_frame, fg_color=C["card"], border_width=1, border_color=C["borde"], corner_radius=12)
        col_der.grid(row=0, column=1, sticky="nsew")

        # Títulos
        ctk.CTkLabel(col_izq, text=tr("Paso 1: Perfil Profesional"), font=("Segoe UI", 20, "bold"), text_color=C["texto"]).pack(anchor="w", pady=(0, 5))
        ctk.CTkLabel(col_izq, text=tr("Ingresa los datos oficiales de tu libreta docente."), font=("Segoe UI", 12), text_color=C["texto_sec"]).pack(anchor="w", pady=(0, 20))

        # Docente nombre (Requerido)
        ctk.CTkLabel(col_izq, text=tr("Nombre del Docente *:"), font=("Segoe UI", 12, "bold"), text_color=C["texto"]).pack(anchor="w")
        self.ent_nom = ctk.CTkEntry(col_izq, width=380, placeholder_text="Ej: Prof. Antonio Pérez")
        self.ent_nom.pack(anchor="w", pady=(4, 12))
        self.ent_nom.insert(0, self.datos["docente_nombre"])

        # Cédula (Requerido)
        ctk.CTkLabel(col_izq, text=tr("Cédula del Docente *:"), font=("Segoe UI", 12, "bold"), text_color=C["texto"]).pack(anchor="w")
        self.ent_ced = ctk.CTkEntry(col_izq, width=380, placeholder_text="Ej: 4-750-1234")
        self.ent_ced.pack(anchor="w", pady=(4, 12))
        self.ent_ced.insert(0, self.datos["docente_cedula"])

        # Escuela nombre (Requerido)
        ctk.CTkLabel(col_izq, text=tr("Nombre de la Escuela *:"), font=("Segoe UI", 12, "bold"), text_color=C["texto"]).pack(anchor="w")
        self.ent_esc = ctk.CTkEntry(col_izq, width=380, placeholder_text="Ej: C.E.B.G. Quebrada de Guabo")
        self.ent_esc.pack(anchor="w", pady=(4, 12))
        self.ent_esc.insert(0, self.datos["escuela_nombre"])

        # Región (Requerido)
        ctk.CTkLabel(col_izq, text=tr("Región Educativa (Panamá):"), font=("Segoe UI", 12, "bold"), text_color=C["texto"]).pack(anchor="w")
        regiones = [
            "Bocas del Toro", "Coclé", "Colón", "Chiriquí", "Darién", "Herrera", "Los Santos",
            "Panamá Centro", "Panamá Este", "Panamá Norte", "Panamá Oeste", "San Miguelito", 
            "Veraguas", "Comarca Ngäbe Buglé", "Comarca Emberá-Wounaan", "Comarca Guna Yala"
        ]
        self.cmb_reg = ctk.CTkOptionMenu(col_izq, values=regiones, width=380, fg_color=C["card"], button_color=C["borde"])
        self.cmb_reg.pack(anchor="w", pady=(4, 12))
        self.cmb_reg.set(self.datos["escuela_region"])

        # Año lectivo (Requerido)
        ctk.CTkLabel(col_izq, text=tr("Año Lectivo *:"), font=("Segoe UI", 12, "bold"), text_color=C["texto"]).pack(anchor="w")
        self.ent_ano = ctk.CTkEntry(col_izq, width=380)
        self.ent_ano.pack(anchor="w", pady=(4, 12))
        self.ent_ano.insert(0, self.datos["ano_lectivo"])

        # Director(a) (Opcional)
        ctk.CTkLabel(col_izq, text=tr("Director(a) de la Escuela:"), font=("Segoe UI", 12, "bold"), text_color=C["texto"]).pack(anchor="w")
        self.ent_dir = ctk.CTkEntry(col_izq, width=380, placeholder_text="Ej: Prof. María Martínez")
        self.ent_dir.pack(anchor="w", pady=(4, 12))
        self.ent_dir.insert(0, self.datos["director_nombre"])

        # Subdirector(a) (Opcional)
        ctk.CTkLabel(col_izq, text=tr("Subdirector(a) de la Escuela:"), font=("Segoe UI", 12, "bold"), text_color=C["texto"]).pack(anchor="w")
        self.ent_sub = ctk.CTkEntry(col_izq, width=380, placeholder_text="Ej: Prof. Juan Gómez")
        self.ent_sub.pack(anchor="w", pady=(4, 12))
        self.ent_sub.insert(0, self.datos["subdirector_nombre"])

        # Jornada
        ctk.CTkLabel(col_izq, text=tr("Jornada Escolar:"), font=("Segoe UI", 12, "bold"), text_color=C["texto"]).pack(anchor="w")
        jornadas = ["Matutina", "Vespertina", "Nocturna", "Extendida"]
        self.cmb_jor = ctk.CTkOptionMenu(col_izq, values=jornadas, width=380, fg_color=C["card"], button_color=C["borde"])
        self.cmb_jor.pack(anchor="w", pady=(4, 12))
        self.cmb_jor.set(self.datos["jornada"])

        # Idioma
        ctk.CTkLabel(col_izq, text=tr("Idioma de la Aplicación:"), font=("Segoe UI", 12, "bold"), text_color=C["texto"]).pack(anchor="w")
        idiomas = ["Español", "English"]
        self.cmb_idioma = ctk.CTkOptionMenu(col_izq, values=idiomas, width=380, fg_color=C["card"], button_color=C["borde"], command=self.cambiar_idioma_wizard)
        self.cmb_idioma.pack(anchor="w", pady=(4, 12))
        self.cmb_idioma.set("Español" if self.datos["idioma"] == "es" else "English")

        # Ruta de exportación (Requerido)
        ctk.CTkLabel(col_izq, text=tr("Ubicación de Exportación de Archivos *:"), font=("Segoe UI", 12, "bold"), text_color=C["texto"]).pack(anchor="w")
        f_exp = ctk.CTkFrame(col_izq, fg_color="transparent")
        f_exp.pack(anchor="w", fill="x", pady=(4, 12))
        self.ent_exp = ctk.CTkEntry(f_exp, width=270)
        self.ent_exp.pack(side="left")
        self.ent_exp.insert(0, self.datos["ruta_exportacion"])
        
        btn_sel_exp = ctk.CTkButton(f_exp, text=tr("Explorar 📂"), width=100, fg_color=C["borde"], hover_color=C["fondo"], command=self.seleccionar_ruta_exportacion)
        btn_sel_exp.pack(side="left", padx=(10, 0))

        # Columna Derecha: Carga de Logo (Opcional)
        ctk.CTkLabel(col_der, text=tr("Logo de la Escuela (Opcional)"), font=("Segoe UI", 14, "bold"), text_color=C["cian"]).pack(pady=15)
        
        self.logo_lbl = ctk.CTkLabel(col_der, text=tr("🚫 Sin Escudo\nCargado"), font=("Segoe UI", 11), text_color=C["texto_sec"], width=120, height=120, fg_color=C["fondo"], corner_radius=8)
        self.logo_lbl.pack(pady=10)

        # Actualizar preview si ya se cargó anteriormente
        if self.datos["logo_escuela_path"]:
            self.mostrar_preview_logo(self.datos["logo_escuela_path"])

        btn_logo = ctk.CTkButton(col_der, text=tr("📁 Cargar Logo"), font=("Segoe UI", 11, "bold"), fg_color=C["borde"], hover_color=C["fondo"], command=self.seleccionar_logo)
        btn_logo.pack(pady=15)
        
        ctk.CTkLabel(col_der, text=tr("Recomendado: Imagen PNG\ncon fondo transparente."), font=("Segoe UI", 10, "italic"), text_color=C["texto_sec"]).pack()


    def seleccionar_ruta_exportacion(self):
        f = filedialog.askdirectory(title="Seleccionar Carpeta de Exportación")
        if f:
            f_norm = os.path.normpath(f)
            self.ent_exp.delete(0, "end")
            self.ent_exp.insert(0, f_norm)

    def seleccionar_logo(self):
        f = filedialog.askopenfilename(filetypes=[("Imágenes", "*.png;*.jpg;*.jpeg")])
        if f:
            user_dir = obtener_dir_datos_usuario()
            dest = os.path.join(user_dir, "data", "escuela_logo.png")
            try:
                shutil.copy2(f, dest)
                self.datos["logo_escuela_path"] = dest
                self.mostrar_preview_logo(dest)
            except Exception as e:
                messagebox.showerror("Error", f"No se pudo copiar la imagen: {e}")

    def mostrar_preview_logo(self, path):
        try:
            img = Image.open(path)
            img.thumbnail((110, 110))
            self.logo_preview_img = ImageTk.PhotoImage(img)
            self.logo_lbl.configure(image=self.logo_preview_img, text="")
        except Exception:
            self.logo_lbl.configure(text=tr("⚠️ Error al\ncargar logo"))

    def cambiar_idioma_wizard(self, opcion):
        nuevo_idioma = "es" if opcion in ["Español", "Spanish"] else "en"
        self.datos["idioma"] = nuevo_idioma
        set_current_lang(nuevo_idioma)
        
        # Guardar valores ingresados actualmente para no perderlos al reconstruir Paso 1
        if hasattr(self, "ent_nom") and self.ent_nom.winfo_exists():
            self.datos["docente_nombre"] = self.ent_nom.get().strip()
        if hasattr(self, "ent_ced") and self.ent_ced.winfo_exists():
            self.datos["docente_cedula"] = self.ent_ced.get().strip()
        if hasattr(self, "ent_esc") and self.ent_esc.winfo_exists():
            self.datos["escuela_nombre"] = self.ent_esc.get().strip()
        if hasattr(self, "cmb_reg") and self.cmb_reg.winfo_exists():
            self.datos["escuela_region"] = self.cmb_reg.get()
        if hasattr(self, "ent_ano") and self.ent_ano.winfo_exists():
            self.datos["ano_lectivo"] = self.ent_ano.get().strip()
        if hasattr(self, "ent_dir") and self.ent_dir.winfo_exists():
            self.datos["director_nombre"] = self.ent_dir.get().strip()
        if hasattr(self, "ent_sub") and self.ent_sub.winfo_exists():
            self.datos["subdirector_nombre"] = self.ent_sub.get().strip()
        if hasattr(self, "cmb_jor") and self.cmb_jor.winfo_exists():
            self.datos["jornada"] = self.cmb_jor.get()
        if hasattr(self, "ent_exp") and self.ent_exp.winfo_exists():
            self.datos["ruta_exportacion"] = self.ent_exp.get().strip()
            
        # Re-dibujar el paso actual y actualizar los botones de navegación
        self.mostrar_paso()
        self.btn_atras.configure(text=tr("⬅️ Atrás"))
        self.btn_omitir.configure(text=tr("Omitir y Usar Demo ❌"))
        self.btn_siguiente.configure(text=tr("Siguiente ➡️"))


    def guardar_paso_1(self) -> bool:
        nom = self.ent_nom.get().strip()
        ced = self.ent_ced.get().strip()
        esc = self.ent_esc.get().strip()
        reg = self.cmb_reg.get()
        ano = self.ent_ano.get().strip()
        dir_n = self.ent_dir.get().strip()
        sub_n = self.ent_sub.get().strip()
        jor = self.cmb_jor.get()
        lang_sel = self.cmb_idioma.get()
        exp = self.ent_exp.get().strip()

        if not nom or not ced or not esc or not ano or not exp:
            messagebox.showwarning("Atención", "Por favor completa todos los campos requeridos (*).")
            return False

        # Validar ruta de exportación
        try:
            os.makedirs(exp, exist_ok=True)
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo crear o verificar la ruta de exportación: {e}")
            return False

        self.datos["docente_nombre"] = nom
        self.datos["docente_cedula"] = ced
        self.datos["escuela_nombre"] = esc
        self.datos["escuela_region"] = reg
        self.datos["ano_lectivo"] = ano
        self.datos["director_nombre"] = dir_n
        self.datos["subdirector_nombre"] = sub_n
        self.datos["jornada"] = jor
        self.datos["idioma"] = "es" if lang_sel in ["Español", "Spanish"] else "en"
        self.datos["ruta_exportacion"] = os.path.normpath(exp)
        
        # Aplicar idioma en caliente para el resto del setup
        set_current_lang(self.datos["idioma"])
        
        return True

    # ══════════════════════════════════════════════════════════════
    #  PASO 2: GRADOS Y SECCIONES
    # ══════════════════════════════════════════════════════════════
    def mostrar_paso_2(self):
        ctk.CTkLabel(self.main_frame, text=tr("Paso 2: Grados y Secciones"), font=("Segoe UI", 20, "bold"), text_color=C["texto"]).pack(anchor="w", pady=(0, 5))
        ctk.CTkLabel(self.main_frame, text=tr("Selecciona la modalidad y los grupos que vas a atender."), font=("Segoe UI", 12), text_color=C["texto_sec"]).pack(anchor="w", pady=(0, 20))

        # Modalidad Selector
        mod_frame = ctk.CTkFrame(self.main_frame, fg_color="transparent")
        mod_frame.pack(fill="x", pady=(0, 15))
        ctk.CTkLabel(mod_frame, text=tr("Modalidad de Enseñanza:"), font=("Segoe UI", 12, "bold")).pack(side="left", padx=(0, 15))

        self.seg_modalidad = ctk.CTkSegmentedButton(mod_frame, values=[tr("Premedia"), tr("Primaria")], selected_color=C["cian"], selected_hover_color=C["hover"], command=self.cambio_modalidad)
        self.seg_modalidad.pack(side="left")
        self.seg_modalidad.set(tr("Premedia") if self.datos["modalidad"] == "premedia" else tr("Primaria"))

        # Contenedor de Grados
        self.grados_list_frame = ctk.CTkScrollableFrame(self.main_frame, fg_color=C["card"], border_width=1, border_color=C["borde"], height=280, corner_radius=12)
        self.grados_list_frame.pack(fill="both", expand=True)

        self.construir_lista_grados()

    def cambio_modalidad(self, val):
        from utils.translator import REVERSE_TRANSLATIONS
        original_val = REVERSE_TRANSLATIONS.get(val, val)
        self.datos["modalidad"] = original_val.lower()
        self.construir_lista_grados()

    def construir_lista_grados(self):
        for w in self.grados_list_frame.winfo_children():
            w.destroy()

        mod = self.datos["modalidad"]
        grados_disponibles = ["7°", "8°", "9°"] if mod == "premedia" else ["1°", "2°", "3°", "4°", "5°", "6°"]

        # Encabezados de la grilla
        header = ctk.CTkFrame(self.grados_list_frame, fg_color="transparent")
        header.pack(fill="x", pady=5)
        ctk.CTkLabel(header, text=tr("Habilitar"), font=("Segoe UI", 11, "bold"), text_color=C["cian"], width=100, anchor="w").pack(side="left", padx=20)
        ctk.CTkLabel(header, text=tr("Grado / Nivel"), font=("Segoe UI", 11, "bold"), text_color=C["cian"], width=120, anchor="w").pack(side="left")
        ctk.CTkLabel(header, text=tr("Sección / Letra (ej: A, B, C)"), font=("Segoe UI", 11, "bold"), text_color=C["cian"], width=200, anchor="w").pack(side="left", padx=20)


        self.grado_checkboxes = {}
        self.grado_entries = {}

        for g in grados_disponibles:
            row = ctk.CTkFrame(self.grados_list_frame, fg_color="transparent")
            row.pack(fill="x", pady=4)

            # Checkbox para habilitar
            chk_var = tk.BooleanVar(value=(g in self.datos["grados_seleccionados"]))
            chk = ctk.CTkCheckBox(row, text="", variable=chk_var, width=100, fg_color=C["verde"])
            chk.pack(side="left", padx=20)
            self.grado_checkboxes[g] = chk_var

            # Nombre Grado
            ctk.CTkLabel(row, text=g, font=("Segoe UI", 13, "bold"), text_color=C["texto"], width=120, anchor="w").pack(side="left")

            # Entrada de Sección
            entry = ctk.CTkEntry(row, width=120, placeholder_text="A")
            entry.pack(side="left", padx=20)
            self.grado_entries[g] = entry

            # Pre-cargar si ya existía
            if g in self.datos["grados_seleccionados"]:
                entry.insert(0, self.datos["grados_seleccionados"][g])
            else:
                entry.insert(0, "A")

    def guardar_paso_2(self) -> bool:
        seleccionados = {}
        for g, var in self.grado_checkboxes.items():
            if var.get():
                seccion = self.grado_entries[g].get().strip().upper()
                if not seccion:
                    seccion = "A"
                seleccionados[g] = seccion

        if not seleccionados:
            messagebox.showwarning("Atención", "Debes activar al menos un grado para continuar.")
            return False

        self.datos["grados_seleccionados"] = seleccionados
        return True

    # ══════════════════════════════════════════════════════════════
    #  PASO 3: MATERIAS Y GRADOS
    # ══════════════════════════════════════════════════════════════
    def mostrar_paso_3(self):
        self.main_frame.columnconfigure(0, weight=5)
        self.main_frame.columnconfigure(1, weight=5)

        col_izq = ctk.CTkFrame(self.main_frame, fg_color="transparent")
        col_izq.grid(row=0, column=0, sticky="nsew", padx=(0, 15))

        col_der = ctk.CTkFrame(self.main_frame, fg_color=C["card"], border_width=1, border_color=C["borde"], corner_radius=12)
        col_der.grid(row=0, column=1, sticky="nsew", padx=(15, 0))

        # Columna Izquierda: Agregar materias
        ctk.CTkLabel(col_izq, text=tr("Paso 3: Materias"), font=("Segoe UI", 20, "bold"), text_color=C["texto"]).pack(anchor="w", pady=(0, 5))
        ctk.CTkLabel(col_izq, text=tr("Agrega tus materias y asígnalas a los grados correspondientes."), font=("Segoe UI", 12), text_color=C["texto_sec"]).pack(anchor="w", pady=(0, 15))

        entry_frame = ctk.CTkFrame(col_izq, fg_color="transparent")
        entry_frame.pack(fill="x", pady=5)
        
        self.ent_materia = ctk.CTkEntry(entry_frame, placeholder_text=tr("Escribe materia (ej: Matemáticas)"), width=240)
        self.ent_materia.pack(side="left")
        self.ent_materia.bind("<Return>", lambda e: self.agregar_materia_lista())

        btn_add = ctk.CTkButton(entry_frame, text=tr("➕ Agregar"), font=("Segoe UI", 12, "bold"), fg_color=C["verde"], hover_color="#059669", width=100, command=self.agregar_materia_lista)
        btn_add.pack(side="left", padx=10)

        # Scrollable list de materias
        self.materias_scroll = ctk.CTkScrollableFrame(col_izq, fg_color=C["card"], border_width=1, border_color=C["borde"], height=240, corner_radius=8)
        self.materias_scroll.pack(fill="both", expand=True, pady=10)

        # Columna Derecha: Mapeo de grados para la materia seleccionada
        self.lbl_mat_select = ctk.CTkLabel(col_der, text=tr("Selecciona una materia a la izquierda"), font=("Segoe UI", 13, "bold"), text_color=C["cian"])
        self.lbl_mat_select.pack(pady=15, padx=20)

        self.grados_map_frame = ctk.CTkFrame(col_der, fg_color="transparent")
        self.grados_map_frame.pack(fill="both", expand=True, padx=25, pady=5)

        self.materia_actual = None
        self.materia_chk_vars = {}
        self.dibujar_lista_materias()

    def agregar_materia_lista(self):
        mat = self.ent_materia.get().strip()
        if not mat: return
        if mat in self.datos["materias_mapeadas"]:
            messagebox.showinfo(tr("Información"), tr("Esta materia ya está en la lista."))
            return

        self.datos["materias_mapeadas"][mat] = []
        self.ent_materia.delete(0, "end")
        self.dibujar_lista_materias()
        self.seleccionar_materia_edicion(mat)

    def eliminar_materia_lista(self, mat):
        if mat in self.datos["materias_mapeadas"]:
            del self.datos["materias_mapeadas"][mat]
        if self.materia_actual == mat:
            self.materia_actual = None
        self.dibujar_lista_materias()
        self.cargar_grados_mapeo()

    def dibujar_lista_materias(self):
        for w in self.materias_scroll.winfo_children():
            w.destroy()

        if not self.datos["materias_mapeadas"]:
            ctk.CTkLabel(self.materias_scroll, text=tr("Sin materias agregadas."), font=("Segoe UI", 11, "italic"), text_color=C["texto_sec"]).pack(pady=40)
            return

        for mat in sorted(self.datos["materias_mapeadas"].keys()):
            row = ctk.CTkFrame(self.materias_scroll, fg_color="transparent")
            row.pack(fill="x", pady=2)

            btn_sel = ctk.CTkButton(row, text=mat, font=("Segoe UI", 12), anchor="w", fg_color="transparent", text_color=C["texto"], hover_color=C["borde"], height=32, command=lambda m=mat: self.seleccionar_materia_edicion(m))
            btn_sel.pack(side="left", fill="x", expand=True)

            btn_del = ctk.CTkButton(row, text="🗑️", fg_color="transparent", text_color=C["rojo"], hover_color=C["borde"], width=30, height=32, command=lambda m=mat: self.eliminar_materia_lista(m))
            btn_del.pack(side="right", padx=5)

    def seleccionar_materia_edicion(self, mat):
        # Guardar cambios del mapeo anterior si corresponde
        self.guardar_mapeo_materia_actual()

        self.materia_actual = mat
        self.lbl_mat_select.configure(text=f"¿Para qué grados es {mat}?")
        self.cargar_grados_mapeo()

    def cargar_grados_mapeo(self):
        for w in self.grados_map_frame.winfo_children():
            w.destroy()

        if not self.materia_actual:
            ctk.CTkLabel(self.grados_map_frame, text="Selecciona una materia\npara configurarle sus grados.", font=("Segoe UI", 11, "italic"), text_color=C["texto_sec"]).pack(pady=60)
            return

        self.materia_chk_vars = {}
        grados_activos = self.datos["grados_seleccionados"]

        ctk.CTkLabel(self.grados_map_frame, text="Grados Disponibles:", font=("Segoe UI", 11, "bold"), text_color=C["texto_sec"]).pack(anchor="w", pady=(0, 10))

        for g, sec in grados_activos.items():
            g_lbl = f"{g} {sec}"
            chk_var = tk.BooleanVar(value=(g_lbl in self.datos["materias_mapeadas"][self.materia_actual]))
            
            chk = ctk.CTkCheckBox(self.grados_map_frame, text=g_lbl, variable=chk_var, font=("Segoe UI", 12), fg_color=C["verde"], pady=6)
            chk.pack(anchor="w", pady=4)
            self.materia_chk_vars[g_lbl] = chk_var

    def guardar_mapeo_materia_actual(self):
        if self.materia_actual and self.materia_chk_vars:
            seleccionados = []
            for g_lbl, var in self.materia_chk_vars.items():
                if var.get():
                    seleccionados.append(g_lbl)
            self.datos["materias_mapeadas"][self.materia_actual] = seleccionados

    def guardar_paso_3(self) -> bool:
        self.guardar_mapeo_materia_actual()

        if not self.datos["materias_mapeadas"]:
            messagebox.showwarning(tr("Atención"), tr("Debes agregar al menos una materia."))
            return False

        # Validar que todas las materias tengan al menos un grado asignado
        for mat, lista in self.datos["materias_mapeadas"].items():
            if not lista:
                messagebox.showwarning(tr("Atención"), tr("La materia '{mat}' no tiene ningún grado seleccionado.").format(mat=mat))
                return False
        return True

    # ══════════════════════════════════════════════════════════════
    #  PASO 4: HORARIO DE CLASES
    # ══════════════════════════════════════════════════════════════
    def mostrar_paso_4(self):
        ctk.CTkLabel(self.main_frame, text=tr("Paso 4: Horario de Clases"), font=("Segoe UI", 20, "bold"), text_color=C["texto"]).pack(anchor="w", pady=(0, 5))
        ctk.CTkLabel(self.main_frame, text=tr("Asigna tus materias a los bloques correspondientes de la semana (puedes dejar celdas vacías)."), font=("Segoe UI", 11), text_color=C["texto_sec"]).pack(anchor="w", pady=(0, 12))

        # Crear contenedor del Horario
        grid_container = ctk.CTkScrollableFrame(self.main_frame, fg_color=C["card"], border_width=1, border_color=C["borde"], height=320, corner_radius=12)
        grid_container.pack(fill="both", expand=True)

        dias = [tr("Hora / Bloque"), tr("Lunes"), tr("Martes"), tr("Miércoles"), tr("Jueves"), tr("Viernes")]
        
        # Configure columns
        for c in range(6):
            grid_container.columnconfigure(c, weight=1, minsize=110)

        # Dibujar headers
        for c, d in enumerate(dias):
            header_lbl = ctk.CTkLabel(grid_container, text=d, font=("Segoe UI", 11, "bold"), text_color=C["cian"], fg_color=C["borde"], height=28)
            header_lbl.grid(row=0, column=c, padx=1, pady=1, sticky="nsew")


        # Horas por bloque
        bloques = [
            "07:00 - 07:45", "07:45 - 08:30", "08:30 - 09:15", "09:15 - 10:00",
            "10:15 - 11:00", "11:00 - 11:45", "11:45 - 12:30", "12:30 - 13:15"
        ]

        # Recopilar lista de opciones validas de materias
        # Ej: "Español - 7° A", "Matemáticas - 8° B"
        opciones_validas = [tr("Libre")]
        for mat, grados_asig in self.datos["materias_mapeadas"].items():
            for g_lbl in grados_asig:
                opciones_validas.append(f"{mat} - {g_lbl}")

        self.horario_dropdowns = {} # key: (bloque, dia_idx)

        # Dibujar grilla de bloques
        dias_ingles = ["lunes", "martes", "miercoles", "jueves", "viernes"]
        
        # Pre-cargar horario previo si existe
        previos = { (s["bloque"], s["dia"]): s["materia_grado"] for s in self.datos["horario_slots"] }

        for r_idx, blq in enumerate(bloques, 1):
            lbl_blq = ctk.CTkLabel(grid_container, text=blq, font=("Segoe UI", 10, "bold"), text_color=C["texto"], fg_color="#0D1E32", height=38)
            lbl_blq.grid(row=r_idx, column=0, padx=1, pady=1, sticky="nsew")

            for c_idx, dia in enumerate(dias_ingles, 1):
                val_inicial = previos.get((blq, dia), "Libre")
                if val_inicial == "Libre":
                    val_inicial = tr("Libre")
                if val_inicial not in opciones_validas:
                    val_inicial = tr("Libre")

                # Dropdown para selección rápida
                cmb = ctk.CTkOptionMenu(grid_container, values=opciones_validas, font=("Segoe UI", 9), fg_color="#0A1628", button_color=C["borde"], height=32)
                cmb.grid(row=r_idx, column=c_idx, padx=1, pady=1, sticky="nsew")
                cmb.set(val_inicial)
                
                self.horario_dropdowns[(blq, dia)] = cmb

    def guardar_paso_4(self) -> bool:
        slots = []
        for (blq, dia), cmb in self.horario_dropdowns.items():
            val = cmb.get()
            if val != tr("Libre"):
                slots.append({
                    "bloque": blq,
                    "dia": dia,
                    "materia_grado": val
                })
        self.datos["horario_slots"] = slots
        return True


    # ══════════════════════════════════════════════════════════════
    #  PASO 5: RESUMEN Y FINALIZAR
    # ══════════════════════════════════════════════════════════════
    def mostrar_paso_5(self):
        ctk.CTkLabel(self.main_frame, text=tr("Paso 5: Resumen y Confirmación"), font=("Segoe UI", 20, "bold"), text_color=C["texto"]).pack(anchor="w", pady=(0, 5))
        ctk.CTkLabel(self.main_frame, text=tr("Verifica que los datos sean correctos antes de construir tu libreta digital."), font=("Segoe UI", 12), text_color=C["texto_sec"]).pack(anchor="w", pady=(0, 20))

        # Panel resumen
        resumen_box = ctk.CTkFrame(self.main_frame, fg_color=C["card"], border_width=1, border_color=C["borde"], corner_radius=12)
        resumen_box.pack(fill="both", expand=True, padx=10, pady=5)

        # Contenido resumen
        txt_resumen = f"{tr('👤 Docente: ')}{self.datos['docente_nombre']}\n"
        txt_resumen += f"{tr('🪪 Cédula: ')}{self.datos['docente_cedula']}\n"
        txt_resumen += f"{tr('🏫 Escuela: ')}{self.datos['escuela_nombre']} ({self.datos['escuela_region']})\n"
        txt_resumen += f"{tr('🗓️ Año Lectivo: ')}{self.datos['ano_lectivo']}\n"
        txt_resumen += f"{tr('📈 Nivel: ')}{self.datos['modalidad'].capitalize()}\n"
        txt_resumen += f"{tr('🖼️ Logo Escolar: ')}{tr('Cargado') if self.datos['logo_escuela_path'] else tr('Por defecto')}\n\n"

        txt_resumen += tr("👥 Grupos Registrados:\n")
        for g, sec in self.datos["grados_seleccionados"].items():
            txt_resumen += f"{tr('  - Grado ')}{g}{tr(' Sección ')}{sec}\n"

        txt_resumen += tr("\n📚 Asignaturas:\n")
        for mat, gs in self.datos["materias_mapeadas"].items():
            txt_resumen += f"  - {mat}{tr(' dictada en: ')}{', '.join(gs)}\n"

        txt_resumen += f"{tr('\n📅 Bloques de Horario Asignados: ')}{len(self.datos['horario_slots'])}{tr(' períodos escolares.')}"

        # TextBox no editable
        tb = ctk.CTkTextbox(resumen_box, fg_color="transparent", font=("Segoe UI", 12), text_color=C["texto"], border_width=0)
        tb.pack(fill="both", expand=True, padx=25, pady=25)
        tb.insert("0.0", txt_resumen)
        tb.configure(state="disabled")

        # Modificar botón siguiente a finalizar
        self.btn_siguiente.configure(text=tr("✅ FINALIZAR CONFIGURACIÓN"), fg_color=C["verde"], hover_color="#059669", width=220, command=self.ejecutar_inicializacion_completa)

    def ejecutar_inicializacion_completa(self):
        self.btn_siguiente.configure(text=tr("⏳ Generando Libreta..."), state="disabled")

        self.btn_atras.configure(state="disabled")
        self.btn_omitir.configure(state="disabled")
        self.update()

        try:
            # 1. Guardar configuración cifrada en config.enc
            guardar_config_segura({
                "modalidad": self.datos["modalidad"],
                "docente_nombre": self.datos["docente_nombre"],
                "docente_cedula": self.datos["docente_cedula"],
                "escuela_nombre": self.datos["escuela_nombre"],
                "escuela_region": self.datos["escuela_region"],
                "ano_lectivo": self.datos["ano_lectivo"],
                "logo_escuela_path": self.datos["logo_escuela_path"],
                "ruta_exportacion": self.datos["ruta_exportacion"],
                "idioma": self.datos["idioma"],
                "jornada": self.datos["jornada"],
                "director_nombre": self.datos["director_nombre"],
                "subdirector_nombre": self.datos["subdirector_nombre"]
            })

            # Guardar la cédula en el registro de Windows para verificación de desinstalación segura
            import winreg
            try:
                key = winreg.CreateKey(winreg.HKEY_CURRENT_USER, r"Software\RegistroDocPro")
                winreg.SetValueEx(key, "Cedula", 0, winreg.REG_SZ, self.datos["docente_cedula"])
                winreg.CloseKey(key)
            except Exception:
                pass

            # 2. Inicializar Base de Datos SQLite
            db = SQLDatabaseManager()
            conn = db.conectar()
            cursor = conn.cursor()

            # Insertar configuraciones
            config_items = [
                ("docente_nombre", self.datos["docente_nombre"]),
                ("docente_cedula", self.datos["docente_cedula"]),
                ("escuela_nombre", self.datos["escuela_nombre"]),
                ("escuela_region", self.datos["escuela_region"]),
                ("ano_lectivo", self.datos["ano_lectivo"]),
                ("logo_escuela_path", self.datos["logo_escuela_path"]),
                ("modalidad_activa", self.datos["modalidad"]),
                ("ruta_exportacion", self.datos["ruta_exportacion"]),
                ("idioma", self.datos["idioma"]),
                ("jornada", self.datos["jornada"]),
                ("director_nombre", self.datos["director_nombre"]),
                ("subdirector_nombre", self.datos["subdirector_nombre"])
            ]
            cursor.executemany("INSERT OR REPLACE INTO configuracion (clave, valor) VALUES (?, ?);", config_items)

            # Insertar Grados
            grado_ids = {} # key: "7° A" -> id
            for g, sec in self.datos["grados_seleccionados"].items():
                cursor.execute(
                    "INSERT OR IGNORE INTO grados (nombre, seccion, modalidad) VALUES (?, ?, ?);",
                    (g, sec, self.datos["modalidad"])
                )
                conn.commit()
                
                cursor.execute(
                    "SELECT id FROM grados WHERE nombre = ? AND seccion = ? AND modalidad = ?;",
                    (g, sec, self.datos["modalidad"])
                )
                g_id = cursor.fetchone()[0]
                grado_ids[f"{g} {sec}"] = g_id

            # Insertar Materias
            materia_ids = {} # key: "Español::7° A" -> id
            for mat, gs in self.datos["materias_mapeadas"].items():
                for g_lbl in gs:
                    g_id = grado_ids[g_lbl]
                    cursor.execute(
                        "INSERT OR IGNORE INTO materias (nombre, grado_id) VALUES (?, ?);",
                        (mat, g_id)
                    )
                    conn.commit()

                    cursor.execute(
                        "SELECT id FROM materias WHERE nombre = ? AND grado_id = ?;",
                        (mat, g_id)
                    )
                    m_id = cursor.fetchone()[0]
                    materia_ids[f"{mat}::{g_lbl}"] = m_id

            # Insertar Horario
            for slot in self.datos["horario_slots"]:
                # slot: {"bloque": "...", "dia": "...", "materia_grado": "Español - 7° A"}
                parts = slot["materia_grado"].split(" - ")
                if len(parts) == 2:
                    mat_name, g_lbl = parts
                    g_id = grado_ids.get(g_lbl)
                    m_id = materia_ids.get(f"{mat_name}::{g_lbl}")
                    
                    cursor.execute(
                        """INSERT INTO horario (dia, bloque_hora, materia_id, grado_id) 
                        VALUES (?, ?, ?, ?);""",
                        (slot["dia"], slot["bloque"], m_id, g_id)
                    )

            conn.commit()
            db.cerrar()

            # 3. Copiar y poblar el Excel limpio en AppData
            archivo = "Registro_Primaria.xlsx" if self.datos["modalidad"] == "primaria" else "Registro_Premedia.xlsx"
            plantilla = os.path.join(ASSETS_DIR, "templates", archivo)
            
            user_dir = obtener_dir_datos_usuario()
            active_excel = os.path.join(user_dir, "temp", archivo)
            os.makedirs(os.path.dirname(active_excel), exist_ok=True)
            
            # Copiar plantilla limpia
            if os.path.exists(plantilla):
                shutil.copy2(plantilla, active_excel)
                
                # Sincronizar datos generales iniciales en el Excel
                engine = DataEngine(ruta_excel=active_excel, modalidad=self.datos["modalidad"])
                engine.actualizar_datos_generales(self.datos["docente_nombre"], self.datos["ano_lectivo"])
                
                # Agregar grados y consejeros iniciales
                for g, sec in self.datos["grados_seleccionados"].items():
                    engine.agregar_grado(g, self.datos["docente_nombre"], "Jornada Regular")
                    
                # Clonar materias
                for mat, gs in self.datos["materias_mapeadas"].items():
                    for g_lbl in gs:
                        g_name = g_lbl.split(" ")[0]
                        # Asegurarse de que el excel tenga la materia
                        engine.clonar_materia(g_name, "General", mat, "Regular")
                
                # Configurar el horario escolar en el excel
                # Mapear horario de slots
                horario_excel_data = []
                # Estructura del horario en excel: lista de filas con {hora: "...", lunes: "...", martes: "...", ...}
                bloques = [
                    "07:00 - 07:45", "07:45 - 08:30", "08:30 - 09:15", "09:15 - 10:00",
                    "10:15 - 11:00", "11:00 - 11:45", "11:45 - 12:30", "12:30 - 13:15"
                ]
                for blq in bloques:
                    row_data = {"hora": blq, "lunes": "", "martes": "", "miercoles": "", "jueves": "", "viernes": ""}
                    for slot in self.datos["horario_slots"]:
                        if slot["bloque"] == blq:
                            row_data[slot["dia"]] = slot["materia_grado"]
                    horario_excel_data.append(row_data)
                
                engine.guardar_horario(horario_excel_data)
                
            else:
                messagebox.showerror("Error", f"No se encontró la plantilla de Excel limpia en:\n{plantilla}")
            
            # Todo listo. Cerrar
            messagebox.showinfo("RegistroDoc Pro", "¡Configuración inicial completada con éxito!\nLa libreta digital escolar ha sido generada.")
            self.destroy()

        except Exception as e:
            messagebox.showerror("Error de Inicialización", f"Ocurrió un error al configurar la base de datos:\n{str(e)}")
            self.btn_siguiente.configure(text="✅ REINTENTAR", state="normal")
            self.btn_atras.configure(state="normal")
            self.btn_omitir.configure(state="normal")