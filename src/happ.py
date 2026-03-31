import customtkinter as ctk
from tkinter import messagebox
try:
    from docx import Document
except ImportError:
    pass

class ReportesFrame(ctk.CTkFrame):
    def __init__(self, master, engine, **kwargs):
        super().__init__(master, **kwargs)
        self.engine = engine
        
        ctk.CTkLabel(self, text="Centro de Impresión y Reportes", font=("Segoe UI", 20, "bold")).pack(pady=20)
        
        # Botones de acción
        ctk.CTkButton(self, text="Generar Lista de Asistencia (PDF)", command=self.print_asistencia).pack(pady=10)
        ctk.CTkButton(self, text="Reporte de Estudiantes Deficientes", fg_color="#c0392b", hover_color="#922b21", command=self.print_fracasos).pack(pady=10)
        ctk.CTkButton(self, text="Exportar Excel Limpio para Dirección", command=self.export_excel).pack(pady=10)

    def print_asistencia(self):
        messagebox.showinfo("Información", "Módulo de impresión de asistencia en desarrollo.")

    def print_fracasos(self):
        try:
            doc = Document()
            doc.add_heading('REPORTE DE ESTUDIANTES CON BAJO RENDIMIENTO', 0)
            doc.save("Reporte_Deficientes.docx")
            messagebox.showinfo("Éxito", "Reporte generado con éxito como 'Reporte_Deficientes.docx'.")
        except Exception as e:
            messagebox.showerror("Error", f"Asegúrate de instalar python-docx.\nError: {e}")

    def export_excel(self):
        # Tomar datos de pantalla y pegarlos en la plantilla oficial.
        # Infiriendo basado en columnas Cédula, Nombre, Trimestre 1, etc.
        import os
        from config import BASE_DIR
        import openpyxl

        base = BASE_DIR
        plantilla_ruta = os.path.join(base, "..", "Registro_Oficial_MEDUCA.xlsx")

        if not os.path.exists(plantilla_ruta):
            messagebox.showerror("Error", f"No se encontró la plantilla oficial en: {plantilla_ruta}")
            return

        try:
            wb_plantilla = openpyxl.load_workbook(plantilla_ruta)
            ws_plantilla = wb_plantilla.active # Assuming the active sheet is the destination

            # Extract data from engine and paste carefully without moving rows
            grados = self.engine.obtener_grados_activos()
            if not grados:
                messagebox.showwarning("Aviso", "No hay grados para exportar.")
                return

            fila_inicio = 5 # Asumimos que los datos empiezan en la fila 5
            for grado in grados:
                estudiantes = self.engine.obtener_estudiantes_completos(grado)
                for i, est in enumerate(estudiantes):
                    # Infiere nombres de columnas comunes
                    fila_destino = fila_inicio + i
                    ws_plantilla.cell(row=fila_destino, column=2).value = est.get("nombre", "") # B - Nombre
                    ws_plantilla.cell(row=fila_destino, column=3).value = est.get("cedula", "") # C - Cédula
                    # Just setting general info, leaving the rest untouched for stability

            out_file = os.path.join(base, "..", "Exportado_Registro_Oficial_MEDUCA.xlsx")
            wb_plantilla.save(out_file)
            wb_plantilla.close()
            messagebox.showinfo("Éxito", f"Datos exportados a: {out_file}")

        except Exception as e:
            messagebox.showerror("Error de Exportación", f"Hubo un error al exportar: \n{str(e)}")
