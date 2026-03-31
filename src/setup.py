import customtkinter as ctk
from tkinter import messagebox
import json
import os
from config import BASE_DIR, CONFIG_FILE
from rddata import DataEngine 


class SetupWizard(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("RegistroDoc Pro - Configuración Inicial")
        self.geometry("700x500")
        self.eval('tk::PlaceWindow . center') 
        self.datos = {}

        self.paso_actual = 1
        self.crear_interfaz()

    def crear_interfaz(self):
        self.main_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.main_frame.pack(fill="both", expand=True, padx=40, pady=40)
        self.mostrar_paso_1()

    def limpiar_frame(self):
        for widget in self.main_frame.winfo_children():
            widget.destroy()

    # ================= PASO 1: DATOS PERSONALES =================
    def mostrar_paso_1(self):
        self.limpiar_frame()
        ctk.CTkLabel(self.main_frame, text="Paso 1 de 3: Datos del Docente", font=("Segoe UI", 24, "bold"), text_color="#3B82F6").pack(pady=(0, 20))
        ctk.CTkLabel(self.main_frame, text="Estos datos aparecerán en las portadas oficiales del MEDUCA.", text_color="#94A3B8").pack(pady=(0, 20))

        ctk.CTkLabel(self.main_frame, text="Nombre Completo:", font=("Segoe UI", 14)).pack(anchor="w")
        self.entry_nombre = ctk.CTkEntry(self.main_frame, width=400, placeholder_text="Ej: Elmer Tugri")
        self.entry_nombre.pack(pady=(5, 15))

        ctk.CTkLabel(self.main_frame, text="Cédula:", font=("Segoe UI", 14)).pack(anchor="w")
        self.entry_cedula = ctk.CTkEntry(self.main_frame, width=400, placeholder_text="Ej: 4-785-823")
        self.entry_cedula.pack(pady=(5, 15))

        ctk.CTkButton(self.main_frame, text="Siguiente ➡️", font=("Segoe UI", 14, "bold"), height=40, command=self.guardar_paso_1).pack(pady=30)

    def guardar_paso_1(self):
        nom = self.entry_nombre.get().strip()
        ced = self.entry_cedula.get().strip()
        if not nom or not ced:
            messagebox.showwarning("Atención", "Por favor completa todos los campos.")
            return
        self.datos["docente_nombre"] = nom
        self.datos["docente_cedula"] = ced
        self.mostrar_paso_2()

    # ================= PASO 2: DATOS DE LA ESCUELA =================
    def mostrar_paso_2(self):
        self.limpiar_frame()
        ctk.CTkLabel(self.main_frame, text="Paso 2 de 3: Institución Educativa", font=("Segoe UI", 24, "bold"), text_color="#3B82F6").pack(pady=(0, 20))

        ctk.CTkLabel(self.main_frame, text="Nombre de la Escuela:", font=("Segoe UI", 14)).pack(anchor="w")
        self.entry_escuela = ctk.CTkEntry(self.main_frame, width=400, placeholder_text="Ej: Escuela Cerro Cacicón")
        self.entry_escuela.pack(pady=(5, 15))

        ctk.CTkLabel(self.main_frame, text="Región Educativa:", font=("Segoe UI", 14)).pack(anchor="w")
        regiones = [
            "Bocas del Toro", "Coclé", "Colón", "Chiriquí", "Darién", "Herrera", "Los Santos",
            "Panamá Centro", "Panamá Este", "Panamá Norte", "Panamá Oeste", "San Miguelito", 
            "Veraguas", "Comarca Ngäbe Buglé", "Comarca Emberá-Wounaan", "Comarca Guna Yala"
        ]
        self.combo_region = ctk.CTkOptionMenu(self.main_frame, values=regiones, width=400)
        self.combo_region.pack(pady=(5, 15))
        self.combo_region.set("Comarca Ngäbe Buglé")

        cont_botones = ctk.CTkFrame(self.main_frame, fg_color="transparent")
        cont_botones.pack(fill="x", pady=30)
        ctk.CTkButton(cont_botones, text="⬅️ Atrás", fg_color="#64748B", hover_color="#475569", command=self.mostrar_paso_1).pack(side="left", padx=10)
        ctk.CTkButton(cont_botones, text="Siguiente ➡️", font=("Segoe UI", 14, "bold"), command=self.guardar_paso_2).pack(side="right", padx=10)

    def guardar_paso_2(self):
        esc = self.entry_escuela.get().strip()
        reg = self.combo_region.get()
        if not esc:
            messagebox.showwarning("Atención", "Escribe el nombre de la escuela.")
            return
        self.datos["escuela_nombre"] = esc
        self.datos["escuela_region"] = reg
        self.mostrar_paso_3()

    # ================= PASO 3: MODALIDAD Y AÑO =================
    def mostrar_paso_3(self):
        self.limpiar_frame()
        ctk.CTkLabel(self.main_frame, text="Paso 3 de 3: Entorno de Trabajo", font=("Segoe UI", 24, "bold"), text_color="#3B82F6").pack(pady=(0, 20))

        ctk.CTkLabel(self.main_frame, text="Año Lectivo:", font=("Segoe UI", 14)).pack(anchor="w")
        self.entry_ano = ctk.CTkEntry(self.main_frame, width=400)
        self.entry_ano.insert(0, "2026")
        self.entry_ano.pack(pady=(5, 15))

        ctk.CTkLabel(self.main_frame, text="¿Qué nivel impartes?", font=("Segoe UI", 14)).pack(anchor="w")
        self.var_modalidad = ctk.StringVar(value="premedia")
        
        rad_frame = ctk.CTkFrame(self.main_frame, fg_color="transparent")
        rad_frame.pack(fill="x", pady=(5, 20))
        ctk.CTkRadioButton(rad_frame, text="Premedia (7°, 8°, 9°)", variable=self.var_modalidad, value="premedia").pack(side="left", padx=20)
        ctk.CTkRadioButton(rad_frame, text="Primaria", variable=self.var_modalidad, value="primaria").pack(side="left", padx=20)

        ctk.CTkLabel(self.main_frame, text="¡El programa configurará y limpiará tu libreta base automáticamente!", text_color="#10B981", font=("Segoe UI", 12, "italic")).pack(pady=(10,0))

        cont_botones = ctk.CTkFrame(self.main_frame, fg_color="transparent")
        cont_botones.pack(fill="x", pady=20)
        ctk.CTkButton(cont_botones, text="⬅️ Atrás", fg_color="#64748B", hover_color="#475569", command=self.mostrar_paso_2).pack(side="left", padx=10)
        
        self.btn_final = ctk.CTkButton(cont_botones, text="✅ FINALIZAR Y LIMPIAR", fg_color="#10B981", hover_color="#059669", font=("Segoe UI", 14, "bold"), command=self.finalizar_setup)
        self.btn_final.pack(side="right", padx=10)

    def finalizar_setup(self):
        ano = self.entry_ano.get().strip()
        if not ano: return
        self.datos["ano_lectivo"] = ano
        self.datos["modalidad"] = self.var_modalidad.get()

        self.btn_final.configure(text="⏳ Limpiando Libreta...", state="disabled")
        self.update() 
        
        # Pausa de 500ms para permitir que termine la animación visual
        self.after(500, self.ejecutar_limpieza_y_cerrar)

    def ejecutar_limpieza_y_cerrar(self):
        # 1. Guardar Configuración JSON
        with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
            json.dump(self.datos, f, ensure_ascii=False, indent=4)

        # 2. Inyectar datos y limpiar el Excel
        archivo = "Registro_Primaria.xlsx" if self.datos["modalidad"] == "primaria" else "Registro_2026.xlsx"
        ruta_excel = os.path.abspath(os.path.join(BASE_DIR, "..", archivo))
        
        if os.path.exists(ruta_excel):
            try:
                engine = DataEngine(ruta_excel, self.datos["modalidad"])
                engine.actualizar_datos_generales(self.datos["docente_nombre"], self.datos["ano_lectivo"])
                engine.reiniciar_libreta() 
            except Exception as e:
                messagebox.showerror("Error en Limpieza", f"Ocurrió un error al limpiar el Excel:\n{e}")
        else:
            messagebox.showerror("Archivo no encontrado", f"No se encontró el archivo: {archivo}")
        
        # 3. Destruir la ventana de forma directa y segura
        self.destroy()