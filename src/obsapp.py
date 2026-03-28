import customtkinter as ctk
from tkinter import messagebox
import datetime
import os
try:
    from docx import Document
    from docx.shared import Pt, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement
    DOCX_DISPONIBLE = True
except ImportError:
    DOCX_DISPONIBLE = False

class ObservacionesFrame(ctk.CTkFrame):
    def __init__(self, master, engine, **kwargs):
        super().__init__(master, fg_color="transparent", **kwargs)
        self.engine = engine
        self.estudiante_seleccionado = None
        self.grado_actual = ""

        # Layout
        self.grid_columnconfigure(0, weight=4) # Lista Izquierda
        self.grid_columnconfigure(1, weight=6) # Panel Derecha
        self.grid_rowconfigure(0, weight=1)

        self.crear_panel_izquierdo()
        self.crear_panel_derecho()
        
        self.al_cambiar_grado(self.combo_grado.get())

    def crear_panel_izquierdo(self):
        frame_izq = ctk.CTkFrame(self, fg_color="#1A2638", corner_radius=10)
        frame_izq.grid(row=0, column=0, sticky="nsew", padx=(0, 10))
        
        # Controles
        top = ctk.CTkFrame(frame_izq, fg_color="transparent")
        top.pack(fill="x", padx=15, pady=15)
        
        ctk.CTkLabel(top, text="Grado:", font=("Segoe UI", 14, "bold")).pack(side="left")
        grados = ["7°", "8°", "9°"] if self.engine.modalidad == "premedia" else ["7°"]
        
        self.combo_grado = ctk.CTkOptionMenu(top, values=grados, command=self.al_cambiar_grado, width=80)
        self.combo_grado.pack(side="left", padx=10)

        # Buscador en tiempo real
        self.entry_buscar = ctk.CTkEntry(top, placeholder_text="Buscar por nombre...", width=180)
        self.entry_buscar.pack(side="left", padx=5)
        self.entry_buscar.bind("<KeyRelease>", self.filtrar_lista)

        # Lista
        self.scroll_estudiantes = ctk.CTkScrollableFrame(frame_izq, fg_color="transparent")
        self.scroll_estudiantes.pack(fill="both", expand=True, padx=10, pady=5)
        self.lista_completa = []

    def crear_panel_derecho(self):
        self.frame_der = ctk.CTkFrame(self, fg_color="#1E2D42", corner_radius=10)
        self.frame_der.grid(row=0, column=1, sticky="nsew")

        ctk.CTkLabel(self.frame_der, text="Expediente y Observaciones", font=("Segoe UI", 20, "bold"), text_color="#3B82F6").pack(pady=(20, 10))
        
        # Etiqueta dinámica del alumno
        self.lbl_alumno_sel = ctk.CTkLabel(self.frame_der, text="Seleccione un estudiante en la lista 👈", font=("Segoe UI", 16, "italic"), text_color="#94A3B8")
        self.lbl_alumno_sel.pack(pady=10)

        # Controles de observación
        panel_obs = ctk.CTkFrame(self.frame_der, fg_color="#1A2638")
        panel_obs.pack(fill="both", expand=True, padx=30, pady=10)

        ctk.CTkLabel(panel_obs, text="Fecha del reporte (DD-MM-AAAA):", font=("Segoe UI", 12)).pack(anchor="w", padx=20, pady=(15,5))
        self.entry_fecha = ctk.CTkEntry(panel_obs)
        self.entry_fecha.insert(0, datetime.datetime.now().strftime("%d-%m-%Y"))
        self.entry_fecha.pack(fill="x", padx=20)

        ctk.CTkLabel(panel_obs, text="Categoría:", font=("Segoe UI", 12)).pack(anchor="w", padx=20, pady=(15,5))
        self.combo_categoria = ctk.CTkOptionMenu(panel_obs, values=["Conducta", "Académico", "Citación a Acudiente", "Mención Honorífica", "Otro"])
        self.combo_categoria.pack(fill="x", padx=20)

        ctk.CTkLabel(panel_obs, text="Redacte la observación:", font=("Segoe UI", 12)).pack(anchor="w", padx=20, pady=(15,5))
        self.texto_obs = ctk.CTkTextbox(panel_obs, height=150)
        self.texto_obs.pack(fill="both", expand=True, padx=20, pady=(0, 20))

        self.btn_guardar = ctk.CTkButton(self.frame_der, text="📝 GUARDAR EN EXPEDIENTE OFICIAL", fg_color="#10B981", hover_color="#059669", 
                      font=("Segoe UI", 14, "bold"), height=45, command=self.guardar_observacion, state="disabled")
        self.btn_guardar.pack(pady=20, padx=30, fill="x")

    def al_cambiar_grado(self, grado):
        self.grado_actual = grado
        self.lista_completa = self.engine.obtener_estudiantes_completos(grado)
        self.entry_buscar.delete(0, 'end')
        self.mostrar_lista(self.lista_completa)
        
        # Reset panel derecho
        self.estudiante_seleccionado = None
        self.lbl_alumno_sel.configure(text="Seleccione un estudiante en la lista 👈", text_color="#94A3B8", font=("Segoe UI", 16, "italic"))
        self.btn_guardar.configure(state="disabled")

    def filtrar_lista(self, event=None):
        texto = self.entry_buscar.get().lower()
        filtrados = [est for est in self.lista_completa if texto in est["nombre"].lower()]
        self.mostrar_lista(filtrados)

    def mostrar_lista(self, lista):
        for w in self.scroll_estudiantes.winfo_children(): w.destroy()
        if not lista:
            ctk.CTkLabel(self.scroll_estudiantes, text="No hay coincidencias.", text_color="#94A3B8").pack(pady=20)
            return

        for est in lista:
            btn = ctk.CTkButton(self.scroll_estudiantes, text=est['nombre'], fg_color="transparent", 
                                anchor="w", text_color="white", hover_color="#3B82F6",
                                command=lambda e=est: self.seleccionar_estudiante(e))
            btn.pack(fill="x", pady=2)

    def seleccionar_estudiante(self, estudiante):
        self.estudiante_seleccionado = estudiante
        self.lbl_alumno_sel.configure(text=f"👨‍🎓 Estudiante: {estudiante['nombre']}", text_color="white", font=("Segoe UI", 18, "bold"))
        self.btn_guardar.configure(state="normal")
        self.texto_obs.delete("1.0", "end")

    def guardar_observacion(self):
        if not self.estudiante_seleccionado: return
        
        fecha = self.entry_fecha.get().strip()
        categoria = self.combo_categoria.get()
        observacion = self.texto_obs.get("1.0", "end").strip()

        if not observacion:
            messagebox.showwarning("Falta información", "Debe redactar la observación antes de guardar.")
            return

        # Guarda en Word usando el mismo generador
        carpeta = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "Expedientes_Estudiantes"))
        if not os.path.exists(carpeta): os.makedirs(carpeta)

        nombre_est = self.estudiante_seleccionado['nombre']
        
        try:
            self._actualizar_o_crear_word(carpeta, nombre_est, self.grado_actual, fecha, categoria, observacion)
            messagebox.showinfo("Éxito", f"Observación agregada al expediente de {nombre_est}.")
            self.texto_obs.delete("1.0", "end") # Limpiamos para la siguiente
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo guardar el archivo: {e}")

    # ==========================================
    # LÓGICA DE WORD (Plantilla Oficial)
    # ==========================================
    def _actualizar_o_crear_word(self, carpeta, nombre_est, grado, fecha, tipo, motivo):
        if not DOCX_DISPONIBLE:
            raise Exception("Librería python-docx no instalada. (pip install python-docx)")

        nombre_archivo = f"{nombre_est} - {grado.replace('°','')}.docx".replace("/", "-")
        ruta_archivo = os.path.join(carpeta, nombre_archivo)

        if os.path.exists(ruta_archivo):
            doc = Document(ruta_archivo)
            if doc.tables:
                tabla = doc.tables[0]
                num_reg = str(len(tabla.rows)) 
                fila = tabla.add_row()
                fila.cells[0].text = num_reg
                fila.cells[1].text = fecha
                fila.cells[2].text = tipo
                fila.cells[3].text = motivo
            doc.save(ruta_archivo)
        else:
            doc = Document()
            p_head = doc.add_paragraph()
            p_head.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run1 = p_head.add_run("MINISTERIO DE EDUCACIÓN\n")
            run1.bold = True
            run1.font.size = Pt(14)
            
            run2 = p_head.add_run("ESCUELA CERRO CACICÓN\n")
            run2.bold = True
            run2.font.size = Pt(12)
            
            run3 = p_head.add_run("DIRECCIÓN REGIONAL COMARCA NGÄBE BUGLÉ")
            run3.font.size = Pt(11)
            
            doc.add_paragraph("\n") 

            p_info = doc.add_paragraph()
            p_info.add_run("Docente: ").bold = True
            p_info.add_run("Elmer Tugri\t\t\t")
            p_info.add_run("Estudiante: ").bold = True
            p_info.add_run(f"{nombre_est}\n")
            
            p_info.add_run("Grado: ").bold = True
            p_info.add_run(f"{grado}\t\t\t\t")
            p_info.add_run("Año Lectivo: ").bold = True
            p_info.add_run("2026")
            
            doc.add_paragraph("\n")

            tabla = doc.add_table(rows=1, cols=4)
            tabla.style = 'Table Grid'
            
            for cell in tabla.rows[0].cells:
                shading_elm = OxmlElement('w:shd')
                shading_elm.set(qn('w:fill'), 'D9E2F3')
                cell._tc.get_or_add_tcPr().append(shading_elm)

            hdr_cells = tabla.rows[0].cells
            hdr_cells[0].text = 'Registro N.º'
            hdr_cells[1].text = 'Fecha'
            hdr_cells[2].text = 'Tipo de registro'
            hdr_cells[3].text = 'Descripción (Motivo u observación)'
            
            for celda in hdr_cells:
                for p in celda.paragraphs:
                    for r in p.runs:
                        r.font.bold = True

            fila = tabla.add_row()
            fila.cells[0].text = "1"
            fila.cells[1].text = fecha
            fila.cells[2].text = tipo
            fila.cells[3].text = motivo

            doc.add_paragraph("\n\n\n\n\n")
            p_firma = doc.add_paragraph("_________________________________\nFirma del Docente")
            p_firma.alignment = WD_ALIGN_PARAGRAPH.CENTER

            section = doc.sections[0]
            footer = section.footer
            p_foot = footer.paragraphs[0]
            p_foot.text = "Documento Oficial de Seguimiento Estudiantil - RegistroDoc Pro"
            p_foot.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p_foot.runs[0].font.size = Pt(8)
            p_foot.runs[0].font.color.rgb = RGBColor(128, 128, 128)

            doc.save(ruta_archivo)