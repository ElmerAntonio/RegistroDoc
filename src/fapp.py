import customtkinter as ctk
from tkinter import messagebox
import datetime
import threading
import os
from config import BASE_DIR

try:
    from docx import Document
    from docx.shared import Pt, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement
    DOCX_DISPONIBLE = True
except ImportError:
    DOCX_DISPONIBLE = False

# DICCIONARIOS DE TRADUCCIÓN PANTALLA <-> EXCEL
UI_A_EXCEL = {"P": ".", "A": "-", "T": "T"}
EXCEL_A_UI = {".": "P", "-": "A", "T": "T", None: "P"}

class AsistenciaFrame(ctk.CTkFrame):
    def __init__(self, master, engine, **kwargs):
        super().__init__(master, fg_color="transparent", **kwargs)
        self.engine = engine
        self.entradas_asistencia = {}
        self.col_a_modificar = None

        self.grid_columnconfigure(0, weight=7)
        self.grid_columnconfigure(1, weight=3)
        self.grid_rowconfigure(0, weight=1)

        self.crear_panel_izquierdo()
        self.crear_panel_derecho()
        
        self.al_cambiar_grado(self.combo_grado.get())

    def crear_panel_izquierdo(self):
        frame_izq = ctk.CTkFrame(self, fg_color="#1A2638", corner_radius=10)
        frame_izq.grid(row=0, column=0, sticky="nsew", padx=(0, 10))
        
        top = ctk.CTkFrame(frame_izq, fg_color="transparent")
        top.pack(fill="x", padx=20, pady=15)
        
        ctk.CTkLabel(top, text="Grado:", font=("Segoe UI", 16, "bold")).pack(side="left")
        grados = ["7°", "8°", "9°"] if self.engine.modalidad == "premedia" else ["7°"]
        
        self.combo_grado = ctk.CTkOptionMenu(top, values=grados, command=self.al_cambiar_grado)
        self.combo_grado.pack(side="left", padx=10)

        header = ctk.CTkFrame(frame_izq, fg_color="#253650", corner_radius=5)
        header.pack(fill="x", padx=15, pady=(5, 0), ipady=5)
        ctk.CTkLabel(header, text="N°", width=30, font=("Segoe UI", 13, "bold")).pack(side="left", padx=5)
        ctk.CTkLabel(header, text="ESTUDIANTE", width=220, anchor="w", font=("Segoe UI", 13, "bold")).pack(side="left", padx=5)
        ctk.CTkLabel(header, text="ESTADO", width=120, font=("Segoe UI", 13, "bold")).pack(side="left", padx=10)
        ctk.CTkLabel(header, text="JUSTIFICACIÓN DE AUSENCIA/TARDANZA", anchor="w", font=("Segoe UI", 13, "bold")).pack(side="left", padx=10)

        self.scroll_estudiantes = ctk.CTkScrollableFrame(frame_izq, fg_color="transparent")
        self.scroll_estudiantes.pack(fill="both", expand=True, padx=15, pady=5)

    def crear_panel_derecho(self):
        frame_der = ctk.CTkFrame(self, fg_color="#1E2D42", corner_radius=10)
        frame_der.grid(row=0, column=1, sticky="nsew")

        ctk.CTkLabel(frame_der, text="Registro Diario", font=("Segoe UI", 18, "bold"), text_color="#3B82F6").pack(pady=(20, 10))
        ctk.CTkLabel(frame_der, text="Leyenda: P(.) A(-) T(T)", text_color="#94A3B8", font=("Segoe UI", 11)).pack()

        self.tabs = ctk.CTkTabview(frame_der)
        self.tabs.pack(fill="both", expand=True, padx=10, pady=10)
        
        tab_nueva = self.tabs.add("Pasar Lista")
        tab_mod = self.tabs.add("Modificar")

        # ====== TAB: PASAR LISTA NUEVA ======
        ctk.CTkLabel(tab_nueva, text="Trimestre:", font=("Segoe UI", 12)).pack(anchor="w", padx=10, pady=(10,0))
        self.combo_trimestre = ctk.CTkOptionMenu(tab_nueva, values=["Trimestre 1", "Trimestre 2", "Trimestre 3"])
        self.combo_trimestre.pack(fill="x", padx=10, pady=5)

        ctk.CTkLabel(tab_nueva, text="Fecha (DD-MM):", font=("Segoe UI", 12)).pack(anchor="w", padx=10, pady=(15,0))
        self.entry_fecha = ctk.CTkEntry(tab_nueva)
        self.entry_fecha.insert(0, datetime.datetime.now().strftime("%d-%m"))
        self.entry_fecha.pack(fill="x", padx=10, pady=5)

        self.btn_guardar_nueva = ctk.CTkButton(tab_nueva, text="💾 GUARDAR ASISTENCIA", fg_color="#10B981", hover_color="#059669", 
                      font=("Segoe UI", 14, "bold"), height=40, command=self.guardar_asistencia)
        self.btn_guardar_nueva.pack(pady=30, padx=10, fill="x")

        # ====== TAB: MODIFICAR ASISTENCIA ======
        ctk.CTkLabel(tab_mod, text="Trimestre a corregir:", font=("Segoe UI", 12)).pack(anchor="w", padx=10, pady=(10,0))
        self.combo_trimestre_mod = ctk.CTkOptionMenu(tab_mod, values=["Trimestre 1", "Trimestre 2", "Trimestre 3"], command=self.cargar_fechas)
        self.combo_trimestre_mod.pack(fill="x", padx=10, pady=5)

        ctk.CTkLabel(tab_mod, text="Seleccione fecha existente:", font=("Segoe UI", 12, "bold")).pack(anchor="w", padx=10, pady=(15,0))
        self.combo_fechas_mod = ctk.CTkOptionMenu(tab_mod, values=["Buscando..."])
        self.combo_fechas_mod.pack(fill="x", padx=10, pady=5)
        
        ctk.CTkButton(tab_mod, text="🔍 CARGAR A LA LISTA", fg_color="#F59E0B", hover_color="#D97706",
                      font=("Segoe UI", 12, "bold"), command=self.buscar_modificar).pack(pady=10, padx=10, fill="x")
                      
        self.btn_actualizar = ctk.CTkButton(tab_mod, text="🔄 ACTUALIZAR EXCEL", fg_color="#3B82F6", hover_color="#2563EB",
                      font=("Segoe UI", 14, "bold"), height=40, command=self.actualizar_asistencia)
        self.btn_actualizar.pack(pady=20, padx=10, fill="x")

    def al_cambiar_grado(self, grado):
        self.cargar_estudiantes(grado)
        self.cargar_fechas()

    def cargar_estudiantes(self, grado=None):
        if grado is None: grado = self.combo_grado.get()
        for w in self.scroll_estudiantes.winfo_children(): w.destroy()
        self.entradas_asistencia.clear()
        self.col_a_modificar = None 
        
        ests = self.engine.obtener_estudiantes_completos(grado)
        for est in ests:
            row = ctk.CTkFrame(self.scroll_estudiantes, fg_color="transparent")
            row.pack(fill="x", pady=2)
            
            ctk.CTkLabel(row, text=f"{est['id']}.", width=30).pack(side="left")
            ctk.CTkLabel(row, text=est['nombre'], width=220, anchor="w").pack(side="left")
            
            seg_btn = ctk.CTkSegmentedButton(row, values=["P", "A", "T"], width=120, selected_color="#3B82F6")
            seg_btn.set("P")
            seg_btn.pack(side="left", padx=10)
            
            # Bloqueado por defecto porque inicia en "P" (Presente)
            entry_exc = ctk.CTkEntry(row, placeholder_text="Solo si falta o llega tarde", fg_color="#0F1923", state="disabled")
            entry_exc.pack(side="left", fill="x", expand=True, padx=5)
            
            seg_btn.configure(command=lambda valor, entry=entry_exc: self.activar_excusa(valor, entry))
            
            self.entradas_asistencia[est['id']] = {"nombre": est['nombre'], "btn": seg_btn, "exc": entry_exc}

    def activar_excusa(self, valor, entry_widget):
        """Habilita la casilla SOLO para ausencias y tardanzas."""
        if valor in ["A", "T"]:
            entry_widget.configure(state="normal", placeholder_text="Escriba la justificación...")
        else:
            entry_widget.delete(0, "end")
            entry_widget.configure(state="disabled", placeholder_text="Solo si falta o llega tarde")

    def cargar_fechas(self, *args):
        grado = self.combo_grado.get()
        trimestre = self.combo_trimestre_mod.get()
        fechas = self.engine.obtener_fechas_asistencia(grado, trimestre)
        
        if fechas:
            self.combo_fechas_mod.configure(values=fechas)
            self.combo_fechas_mod.set(fechas[-1]) 
        else:
            vacio = ["Sin registro en este trimestre"]
            self.combo_fechas_mod.configure(values=vacio)
            self.combo_fechas_mod.set(vacio[0])

    def recopilar_datos(self):
        dic_asistencia = {}
        lista_excusas = []
        for id_est, widgets in self.entradas_asistencia.items():
            estado_ui = widgets["btn"].get()
            excusa = widgets["exc"].get().strip()
            
            estado_excel = UI_A_EXCEL.get(estado_ui, ".")
            dic_asistencia[id_est] = {"estado": estado_excel}
            
            # Solo guardamos excusas si es Ausencia o Tardanza
            if estado_ui in ["A", "T"]:
                tipo_reg = "Ausencia" if estado_ui == "A" else "Tardanza"
                lista_excusas.append({
                    "nombre": widgets["nombre"],
                    "estado": tipo_reg,
                    "motivo": excusa if excusa else "Falta sin justificar"
                })
        return dic_asistencia, lista_excusas

    def guardar_asistencia(self):
        grado = self.combo_grado.get()
        trimestre = self.combo_trimestre.get()
        fecha = self.entry_fecha.get().strip()

        if not fecha:
            messagebox.showwarning("Atención", "Debe colocar una fecha.")
            return

        dic_asistencia, lista_excusas = self.recopilar_datos()

        self.btn_guardar_nueva.configure(text="⏳ GUARDANDO...", fg_color="#F59E0B", state="disabled")
        self.update()

        def tarea():
            exito, msj = self.engine.guardar_asistencia(grado, trimestre, fecha, dic_asistencia)
            self.after(0, lambda: self.finalizar_guardado(exito, msj, lista_excusas, grado, fecha))

        threading.Thread(target=tarea, daemon=True).start()

    def finalizar_guardado(self, exito, msj, lista_excusas, grado, fecha):
        self.btn_guardar_nueva.configure(text="💾 GUARDAR ASISTENCIA", fg_color="#10B981", state="normal")
        
        if not exito:
            messagebox.showerror("Error de Excel", msj)
            return
            
        self.cargar_fechas()
        self.procesar_expedientes_word(grado, fecha, lista_excusas, "Guardado exitoso.")
        self.cargar_estudiantes(grado)

    def buscar_modificar(self):
        grado = self.combo_grado.get()
        trimestre = self.combo_trimestre_mod.get()
        fecha = self.combo_fechas_mod.get()
        
        if "Sin registro" in fecha or "Buscando" in fecha: return
            
        resultado = self.engine.buscar_asistencia_existente(grado, trimestre, fecha)
        if not resultado: return
            
        self.col_a_modificar = resultado["columna"]
        datos_excel = resultado["asistencia"]
        
        for id_est, widgets in self.entradas_asistencia.items():
            if id_est in datos_excel:
                estado_excel = datos_excel[id_est]
                estado_ui = EXCEL_A_UI.get(estado_excel, "P")
                widgets["btn"].set(estado_ui)
                self.activar_excusa(estado_ui, widgets["exc"])
                widgets["exc"].delete(0, 'end') 
                
        messagebox.showinfo("Modo Edición", f"Asistencia del {fecha} cargada.\n\nEdítela y presione ACTUALIZAR EXCEL.")

    def actualizar_asistencia(self):
        if not self.col_a_modificar: return
            
        grado = self.combo_grado.get()
        trimestre = self.combo_trimestre_mod.get()
        fecha = self.combo_fechas_mod.get()
        dic_asistencia, lista_excusas = self.recopilar_datos()
        
        self.btn_actualizar.configure(text="⏳ ACTUALIZANDO...", fg_color="#F59E0B", state="disabled")
        self.update()
        
        def tarea():
            exito = self.engine.actualizar_asistencia(grado, trimestre, self.col_a_modificar, dic_asistencia)
            self.after(0, lambda: self.finalizar_actualizacion(exito, grado, fecha, lista_excusas))
            
        threading.Thread(target=tarea, daemon=True).start()

    def finalizar_actualizacion(self, exito, grado, fecha, lista_excusas):
        self.btn_actualizar.configure(text="🔄 ACTUALIZAR EXCEL", fg_color="#3B82F6", state="normal")
        if exito:
            self.col_a_modificar = None
            self.procesar_expedientes_word(grado, fecha, lista_excusas, "Asistencia actualizada correctamente.")
            self.cargar_estudiantes(grado)
        else:
            messagebox.showerror("Error", "No se pudo actualizar.")

    # =============================================================
    # MOTOR GENERADOR DE EXPEDIENTES EN WORD 
    # =============================================================
    def procesar_expedientes_word(self, grado, fecha, lista_excusas, mensaje_base):
        if not lista_excusas:
            messagebox.showinfo("Completado", f"{mensaje_base}\n\n(Asistencia perfecta. No se generaron justificaciones).")
            return
            
        if not DOCX_DISPONIBLE:
            messagebox.showinfo("Completado", f"{mensaje_base}\n\n(Falta la librería python-docx para generar los Word).")
            return

        carpeta_expedientes = os.path.join(BASE_DIR, "..", "Expedientes_Estudiantes")
        if not os.path.exists(carpeta_expedientes):
            os.makedirs(carpeta_expedientes)

        for exc in lista_excusas:
            self._actualizar_o_crear_word(carpeta_expedientes, exc['nombre'], grado, fecha, exc['estado'], exc['motivo'])
            
        messagebox.showinfo("Completado", f"{mensaje_base}\n\nSe actualizaron los expedientes en Word de {len(lista_excusas)} estudiantes por inasistencia/tardanza.")

    def _actualizar_o_crear_word(self, carpeta, nombre_est, grado, fecha, tipo, motivo):
        nombre_archivo = f"{nombre_est} - {grado.replace('°','')}.docx".replace("/", "-")
        ruta_archivo = os.path.join(carpeta, nombre_archivo)

        if os.path.exists(ruta_archivo):
            try:
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
            except Exception as e:
                print(f"No se pudo actualizar el Word de {nombre_est}: {e}")
        else:
            doc = Document()
        try:
            from utils.footer_utils import add_header_with_logo, get_school_logo_path
            logo_esc = get_school_logo_path()
            if logo_esc:
                add_header_with_logo(doc, logo_esc)
        except Exception: pass


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

            try:
                doc.save(ruta_archivo)
            except Exception as e:
                print(f"Error guardando Word: {e}")