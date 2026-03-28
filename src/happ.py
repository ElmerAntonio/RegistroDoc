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
        messagebox.showinfo("Información", "Módulo de exportación Excel en desarrollo.")