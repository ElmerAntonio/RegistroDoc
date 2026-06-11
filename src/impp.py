import os
import tkinter as tk
from tkinter import messagebox
import customtkinter as ctk
import openpyxl

from rdprint import abrir_para_imprimir, imprimir_hoja_directo

class ImpresionFrame(ctk.CTkFrame):
    def __init__(self, master, engine, **kwargs):
        super().__init__(master, fg_color="transparent", **kwargs)
        self.engine = engine

        # Paleta de colores
        self.C = {
            "azul_osc": "#1E293B",
            "azul_med": "#3B82F6",
            "verde": "#10B981",
            "card_bg": "#FFFFFF" if ctk.get_appearance_mode() == "Light" else "#1E293B",
            "texto": "#0F172A" if ctk.get_appearance_mode() == "Light" else "#F8FAFC",
            "texto_sec": "#64748B" if ctk.get_appearance_mode() == "Light" else "#94A3B8"
        }

        # Título Principal
        title_lbl = ctk.CTkLabel(
            self, text="🖨️ Centro de Impresión y Exportación",
            font=("Outfit", 24, "bold"), text_color=self.C["texto"]
        )
        title_lbl.pack(anchor="w", padx=30, pady=(25, 5))

        desc_lbl = ctk.CTkLabel(
            self, text="Seleccione qué reporte desea visualizar, imprimir o abrir directamente en Microsoft Excel.",
            font=("Inter", 13), text_color=self.C["texto_sec"]
        )
        desc_lbl.pack(anchor="w", padx=30, pady=(0, 20))

        # Contenedor Grid (2 Columnas)
        grid_frame = ctk.CTkFrame(self, fg_color="transparent")
        grid_frame.pack(fill="both", expand=True, padx=30, pady=5)
        grid_frame.columnconfigure(0, weight=1, uniform="grid")
        grid_frame.columnconfigure(1, weight=1, uniform="grid")
        grid_frame.rowconfigure(0, weight=1)

        # Panel Izquierdo: Configuración
        config_card = ctk.CTkFrame(
            grid_frame, fg_color=self.C["card_bg"], corner_radius=15, 
            border_width=1, border_color="#E2E8F0" if ctk.get_appearance_mode() == "Light" else "#334155"
        )
        config_card.grid(row=0, column=0, sticky="nsew", padx=(0, 10), pady=10)

        # Panel Derecho: Acciones y Resumen
        actions_card = ctk.CTkFrame(
            grid_frame, fg_color=self.C["card_bg"], corner_radius=15, 
            border_width=1, border_color="#E2E8F0" if ctk.get_appearance_mode() == "Light" else "#334155"
        )
        actions_card.grid(row=0, column=1, sticky="nsew", padx=(10, 0), pady=10)

        # --- PANEL IZQUIERDO: CONFIGURACIÓN ---
        ctk.CTkLabel(
            config_card, text="⚙️ Parámetros del Reporte", 
            font=("Outfit", 16, "bold"), text_color=self.C["texto"]
        ).pack(anchor="w", padx=25, pady=(20, 15))

        # 1. Tipo de Reporte
        ctk.CTkLabel(
            config_card, text="1. Seleccione el Tipo de Reporte:", 
            font=("Inter", 12, "bold"), text_color=self.C["texto"]
        ).pack(anchor="w", padx=25, pady=(5, 5))
        
        self.report_type_var = ctk.StringVar(value="Portada")
        self.report_types = [
            ("📋 Portada Oficial", "Portada"),
            ("📄 Carátula del Registro", "Caratula"),
            ("🕐 Horario Escolar", "Horarios"),
            ("📊 Planilla de Calificaciones", "Planilla"),
            ("📅 Asistencia de Estudiantes", "Asistencia"),
            ("📝 Resumen de Calificaciones", "Resumen"),
            ("📚 Libro de Registro Completo", "Libro Completo"),
            ("📝 Plantilla en Blanco (Calificar a Mano)", "Plantilla Blanco")
        ]

        self.report_menu = ctk.CTkOptionMenu(
            config_card, values=[t[0] for t in self.report_types],
            command=self._on_report_type_change,
            fg_color="#3B82F6", button_color="#2563EB", button_hover_color="#1D4ED8"
        )
        self.report_menu.pack(fill="x", padx=25, pady=(0, 15))

        # 2. Grado (Condicional)
        self.grado_label = ctk.CTkLabel(
            config_card, text="2. Seleccione el Grado:", 
            font=("Inter", 12, "bold"), text_color=self.C["texto"]
        )
        self.grado_label.pack(anchor="w", padx=25, pady=(5, 5))
        
        self.grados_activos = self.engine.obtener_grados_activos()
        self.grado_menu = ctk.CTkOptionMenu(
            config_card, values=self.grados_activos if self.grados_activos else ["No hay grados"],
            command=self._on_grado_change
        )
        self.grado_menu.pack(fill="x", padx=25, pady=(0, 15))

        # 3. Asignatura (Condicional)
        self.materia_label = ctk.CTkLabel(
            config_card, text="3. Seleccione la Asignatura:", 
            font=("Inter", 12, "bold"), text_color=self.C["texto"]
        )
        self.materia_label.pack(anchor="w", padx=25, pady=(5, 5))

        self.materia_menu = ctk.CTkOptionMenu(config_card, values=["General"])
        self.materia_menu.pack(fill="x", padx=25, pady=(0, 15))

        # 4. Periodo (Condicional)
        self.periodo_label = ctk.CTkLabel(
            config_card, text="4. Seleccione el Periodo:", 
            font=("Inter", 12, "bold"), text_color=self.C["texto"]
        )
        self.periodo_label.pack(anchor="w", padx=25, pady=(5, 5))
        self.periodo_menu = ctk.CTkOptionMenu(
            config_card, values=["Trimestre 1", "Trimestre 2", "Trimestre 3", "Anual"],
            command=self._actualizar_info
        )
        self.periodo_menu.pack(fill="x", padx=25, pady=(0, 20))

        # --- PANEL DERECHO: ACCIONES ---
        ctk.CTkLabel(
            actions_card, text="🚀 Acciones de Impresión", 
            font=("Outfit", 16, "bold"), text_color=self.C["texto"]
        ).pack(anchor="w", padx=25, pady=(20, 15))

        # Info Box
        info_frame = ctk.CTkFrame(
            actions_card, fg_color="#EFF6FF" if ctk.get_appearance_mode() == "Light" else "#1E293B", 
            corner_radius=10
        )
        info_frame.pack(fill="x", padx=25, pady=(0, 25))
        
        self.info_lbl = ctk.CTkLabel(
            info_frame, text="Se enviará a imprimir la hoja correspondiente al reporte seleccionado.",
            font=("Inter", 12), text_color="#1E40AF" if ctk.get_appearance_mode() == "Light" else "#93C5FD",
            wraplength=250, justify="left"
        )
        self.info_lbl.pack(padx=15, pady=15)

        # Botones de Acción
        self.btn_open = ctk.CTkButton(
            actions_card, text="📂 Abrir en Microsoft Excel", font=("Inter", 13, "bold"),
            fg_color="#3B82F6", hover_color="#2563EB", text_color="#FFFFFF",
            height=45, command=self._abrir_excel
        )
        self.btn_open.pack(fill="x", padx=25, pady=8)

        self.btn_print = ctk.CTkButton(
            actions_card, text="🖨️ Enviar a Impresora", font=("Inter", 13, "bold"),
            fg_color="#10B981", hover_color="#059669", text_color="#FFFFFF",
            height=45, command=self._imprimir_directo
        )
        self.btn_print.pack(fill="x", padx=25, pady=8)

        self._on_report_type_change()

    def _on_report_type_change(self, *args):
        label = self.report_menu.get()
        tipo = "Portada"
        for t in self.report_types:
            if t[0] == label:
                tipo = t[1]
                break
        
        self.report_type_var.set(tipo)

        if tipo in ["Portada", "Caratula", "Horarios", "Libro Completo"]:
            self.grado_label.pack_forget()
            self.grado_menu.pack_forget()
            self.materia_label.pack_forget()
            self.materia_menu.pack_forget()
            self.periodo_label.pack_forget()
            self.periodo_menu.pack_forget()
        elif tipo == "Asistencia":
            self.grado_label.pack(anchor="w", padx=25, pady=(5, 5))
            self.grado_menu.pack(fill="x", padx=25, pady=(0, 15))
            self.materia_label.pack_forget()
            self.materia_menu.pack_forget()
            self.periodo_label.pack(anchor="w", padx=25, pady=(5, 5))
            self.periodo_menu.pack(fill="x", padx=25, pady=(0, 20))
            self.periodo_menu.configure(values=["Trimestre 1", "Trimestre 2", "Trimestre 3"])
            self.periodo_menu.set("Trimestre 1")
        elif tipo == "Resumen":
            self.grado_label.pack(anchor="w", padx=25, pady=(5, 5))
            self.grado_menu.pack(fill="x", padx=25, pady=(0, 15))
            self.materia_label.pack_forget()
            self.materia_menu.pack_forget()
            self.periodo_label.pack(anchor="w", padx=25, pady=(5, 5))
            self.periodo_menu.pack(fill="x", padx=25, pady=(0, 20))
            self.periodo_menu.configure(values=["Trimestre 1", "Anual"])
            self.periodo_menu.set("Trimestre 1")
        elif tipo == "Planilla":
            self.grado_label.pack(anchor="w", padx=25, pady=(5, 5))
            self.grado_menu.pack(fill="x", padx=25, pady=(0, 15))
            self.materia_label.pack(anchor="w", padx=25, pady=(5, 5))
            self.materia_menu.pack(fill="x", padx=25, pady=(0, 15))
            self.periodo_label.pack_forget()
            self.periodo_menu.pack_forget()
            self._on_grado_change()
        elif tipo == "Plantilla Blanco":
            self.grado_label.pack(anchor="w", padx=25, pady=(5, 5))
            self.grado_menu.pack(fill="x", padx=25, pady=(0, 15))
            self.materia_label.pack_forget()
            self.materia_menu.pack_forget()
            self.periodo_label.pack_forget()
            self.periodo_menu.pack_forget()

        self._actualizar_info()

    def _on_grado_change(self, *args):
        tipo = self.report_type_var.get()
        if tipo == "Planilla":
            grado = self.grado_menu.get()
            materias = self.engine.obtener_materias_por_grado(grado)
            self.materia_menu.configure(values=materias if materias else ["General"])
            if materias:
                self.materia_menu.set(materias[0])
            else:
                self.materia_menu.set("General")
        self._actualizar_info()

    def _actualizar_info(self, *args):
        tipo = self.report_type_var.get()
        if tipo == "Libro Completo":
            msg = "Se abrirá o imprimirá el libro de calificaciones completo con todas sus hojas."
        elif tipo in ["Portada", "Caratula", "Horarios"]:
            msg = f"Se procesará la hoja general de '{tipo}' del registro."
        elif tipo == "Plantilla Blanco":
            grado = self.grado_menu.get()
            msg = f"Se generará una plantilla en blanco con los nombres del grado {grado} para calificar a mano."
        else:
            grado = self.grado_menu.get()
            periodo = self.periodo_menu.get() if tipo != "Planilla" else "N/A"
            msg = f"Reporte: {tipo}\nGrado: {grado}\n"
            if tipo == "Planilla":
                msg += f"Asignatura: {self.materia_menu.get()}"
            else:
                msg += f"Periodo: {periodo}"
        self.info_lbl.configure(text=msg)

    def _resolver_nombre_hoja(self):
        tipo = self.report_type_var.get()
        if tipo in ["Portada", "Caratula", "Horarios"]:
            return tipo

        if tipo == "Libro Completo":
            return None

        grado = self.grado_menu.get()
        materia = self.materia_menu.get() if tipo == "Planilla" else None
        
        if not os.path.exists(self.engine.ruta):
            return None
        
        wb = openpyxl.load_workbook(self.engine.ruta, read_only=True)
        from rdprint import encontrar_hoja_impresion
        nombre_hoja = encontrar_hoja_impresion(wb, tipo, grado, materia)
        wb.close()
        return nombre_hoja

    def _abrir_excel(self):
        tipo = self.report_type_var.get()
        if tipo == "Plantilla Blanco":
            self._generar_y_abrir_plantilla_blanco()
            return

        hoja = self._resolver_nombre_hoja()
        ok, msg = abrir_para_imprimir(hoja)
        if ok:
            messagebox.showinfo("✓ Éxito", msg)
        else:
            messagebox.showerror("Error", msg)

    def _imprimir_directo(self):
        tipo = self.report_type_var.get()
        if tipo == "Plantilla Blanco":
            self._generar_y_abrir_plantilla_blanco()
            return

        hoja = self._resolver_nombre_hoja()
        if not hoja:
            if self.report_type_var.get() == "Libro Completo":
                ok, msg = abrir_para_imprimir(None)
                if ok:
                    messagebox.showinfo("✓ Éxito", msg)
                else:
                    messagebox.showerror("Error", msg)
                return
            messagebox.showerror("Error", "No se encontró la hoja correspondiente en el archivo Excel.")
            return

        ok, msg = imprimir_hoja_directo(hoja)
        if ok:
            messagebox.showinfo("✓ Enviado a Impresora", msg)
        else:
            messagebox.showerror("Error", msg)

    def _generar_y_abrir_plantilla_blanco(self):
        import openpyxl
        from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
        import tempfile
        
        grado = self.grado_menu.get()
        estudiantes = self.engine.obtener_estudiantes_completos(grado)
        if not estudiantes:
            messagebox.showwarning("Atención", "No se encontraron estudiantes para el grado seleccionado.")
            return

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Plantilla Auxiliar"
        
        # Grid lines visible
        ws.views.sheetView[0].showGridLines = True
        
        # Styles
        font_title = Font(name="Calibri", size=16, bold=True, color="1B365D")
        font_subtitle = Font(name="Calibri", size=11, bold=True, color="475569")
        font_header = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
        font_body = Font(name="Calibri", size=11)
        font_bold = Font(name="Calibri", size=11, bold=True)
        
        fill_header = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid") # Dark blue
        fill_zebra = PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid") # Sutil grey
        
        thin_border = Border(
            left=Side(style='thin', color='CBD5E1'),
            right=Side(style='thin', color='CBD5E1'),
            top=Side(style='thin', color='CBD5E1'),
            bottom=Side(style='thin', color='CBD5E1')
        )
        
        # Headers
        ws.merge_cells("A1:N1")
        ws["A1"] = "MINISTERIO DE EDUCACIÓN — REGISTRO AUXILIAR"
        ws["A1"].font = font_title
        ws["A1"].alignment = Alignment(horizontal="center")
        
        ws.merge_cells("A2:N2")
        ws["A2"] = f"Grado: {grado}  |  Trimestre: _______  |  Asignatura: ___________________________"
        ws["A2"].font = font_subtitle
        ws["A2"].alignment = Alignment(horizontal="center")
        
        headers = ["N°", "Estudiante (Apellido, Nombre)", "C1", "C2", "C3", "C4", "C5", "C6", "C7", "C8", "C9", "C10", "Promedio", "Observaciones"]
        ws.row_dimensions[4].height = 28
        
        for col_idx, text in enumerate(headers, 1):
            cell = ws.cell(row=4, column=col_idx)
            cell.value = text
            cell.font = font_header
            cell.fill = fill_header
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = thin_border
            
        # Add students
        current_row = 5
        for idx, est in enumerate(estudiantes, 1):
            ws.row_dimensions[current_row].height = 20
            
            # Row number
            cell_num = ws.cell(row=current_row, column=1, value=idx)
            cell_num.font = font_bold
            cell_num.alignment = Alignment(horizontal="center")
            cell_num.border = thin_border
            
            # Name
            cell_name = ws.cell(row=current_row, column=2, value=est["nombre"])
            cell_name.font = font_body
            cell_name.alignment = Alignment(horizontal="left", vertical="center")
            cell_name.border = thin_border
            
            # Zebra striping
            if idx % 2 == 0:
                cell_num.fill = fill_zebra
                cell_name.fill = fill_zebra
                
            # Grade columns (3 to 14)
            for col_idx in range(3, 15):
                cell_g = ws.cell(row=current_row, column=col_idx)
                cell_g.border = thin_border
                if idx % 2 == 0:
                    cell_g.fill = fill_zebra
                    
            current_row += 1
            
        # Auto-adjust column widths
        ws.column_dimensions["A"].width = 6
        ws.column_dimensions["B"].width = 35
        for col_idx in range(3, 13):
            col_letter = openpyxl.utils.get_column_letter(col_idx)
            ws.column_dimensions[col_letter].width = 6
        ws.column_dimensions["M"].width = 10
        ws.column_dimensions["N"].width = 25
        
        # Save temp file
        temp_dir = tempfile.gettempdir()
        file_path = os.path.join(temp_dir, f"Plantilla_Limpia_{grado.replace('°','')}.xlsx")
        try:
            wb.save(file_path)
            # Open file
            if os.name == 'nt':
                os.startfile(file_path)
            else:
                import subprocess
                subprocess.call(['open', file_path])
            messagebox.showinfo("✓ Éxito", f"Se generó la plantilla en blanco con éxito para el grado {grado}.\n\nSe abrirá en Excel para que la pueda imprimir.")
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo guardar o abrir el archivo: {e}")
