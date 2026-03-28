import os
import openpyxl
from openpyxl.styles import Alignment, Font
from utils.cleaner import ExcelCleaner

class DataEngine:
    def __init__(self, ruta_excel, modalidad="premedia"):
        self.ruta = ruta_excel
        self.modalidad = modalidad.lower()
        self.cleaner = ExcelCleaner()
        self.fila_desc = 39 if self.modalidad == "primaria" else 42

    def reiniciar_libreta(self):
        return self.cleaner.limpiar_todo(self.ruta, self.modalidad)

    def _obtener_columna_nombres(self, grado):
        if self.modalidad == "primaria": return 2
        if "7" in grado: return 2
        if "8" in grado: return 4
        if "9" in grado: return 6
        return 2

    def obtener_materias_por_grado(self, grado):
        if not os.path.exists(self.ruta): return []
        wb = openpyxl.load_workbook(self.ruta, read_only=True)
        materias = []
        grado_num = grado.replace("°", "")
        
        for sheet in wb.sheetnames:
            if "PROM" in sheet.upper():
                if self.modalidad == "premedia" and grado_num in sheet:
                    mat = sheet.upper().replace("PROM", "").replace("(", "").replace(")", "").replace(grado, "").replace(grado_num, "").replace("°", "").strip()
                    materias.append(mat.title())
                elif self.modalidad == "primaria":
                    mat = sheet.upper().replace("PROM", "").replace("(", "").replace(")", "").strip()
                    materias.append(mat.title())
        wb.close()
        return sorted(list(set(materias))) if materias else ["Sin materias registradas"]

    def obtener_estudiantes_completos(self, grado):
        if not os.path.exists(self.ruta): return []
        wb = openpyxl.load_workbook(self.ruta, data_only=True)
        ws_m = wb["MAESTRO"]
        col_nom = self._obtener_columna_nombres(grado)
        
        ws_planilla = None
        if self.modalidad == "premedia":
            for sheet in wb.sheetnames:
                if "Planilla" in sheet and grado.replace("°","") in sheet:
                    ws_planilla = wb[sheet]
                    break

        estudiantes = []
        for r in range(5, 46):
            nom = ws_m.cell(row=r, column=col_nom).value
            if nom:
                cedula = ""
                if self.modalidad == "primaria":
                    cedula = ws_m.cell(row=r, column=5).value
                elif ws_planilla:
                    fila_plan = 15 + (r - 4) 
                    cedula = ws_planilla.cell(row=fila_plan, column=5).value
                
                estudiantes.append({
                    "id": r - 4,
                    "nombre": str(nom).strip(), 
                    "cedula": str(cedula).strip() if cedula else ""
                })
        wb.close()
        return estudiantes

    def agregar_estudiante(self, grado, nombre, cedula=""):
        if not os.path.exists(self.ruta): return False
        wb = openpyxl.load_workbook(self.ruta)
        ws_m = wb["MAESTRO"]
        col_nom = self._obtener_columna_nombres(grado)
        
        fila_vacia = None
        for r in range(5, 51):
            if not ws_m.cell(row=r, column=col_nom).value:
                fila_vacia = r
                break
                
        if not fila_vacia:
            wb.close()
            return False 
        
        ws_m.cell(row=fila_vacia, column=col_nom).value = nombre
        if self.modalidad == "primaria":
            ws_m.cell(row=fila_vacia, column=5).value = cedula
        else:
            num_grado = grado.replace("°","")
            id_est = fila_vacia - 4
            for sheet in wb.sheetnames:
                if "Planilla" in sheet and num_grado in sheet:
                    ws_p = wb[sheet]
                    ws_p.cell(row=15+id_est, column=5).value = cedula
        wb.save(self.ruta)
        wb.close()
        return True

    def guardar_cambios_estudiantes(self, grado, datos_modificados):
        if not os.path.exists(self.ruta): return False
        wb = openpyxl.load_workbook(self.ruta)
        ws_m = wb["MAESTRO"]
        col_nom = self._obtener_columna_nombres(grado)
        
        for id_est, datos in datos_modificados.items():
            fila = 4 + int(id_est)
            ws_m.cell(row=fila, column=col_nom).value = datos["nombre"]
            
            if self.modalidad == "primaria":
                ws_m.cell(row=fila, column=5).value = datos["cedula"]
            else:
                num_grado = grado.replace("°","")
                for sheet in wb.sheetnames:
                    if "Planilla" in sheet and num_grado in sheet:
                        ws_p = wb[sheet]
                        ws_p.cell(row=15+int(id_est), column=5).value = datos["cedula"]
        wb.save(self.ruta)
        wb.close()
        return True

    def get_dashboard_stats(self):
        if not os.path.exists(self.ruta): return {"total": 0, "riesgo": 0, "honor": "N/A", "asistencia": "0%"}
        if self.modalidad == "primaria":
            total = len(self.obtener_estudiantes_completos("7°"))
        else:
            total = sum(len(self.obtener_estudiantes_completos(g)) for g in ["7°", "8°", "9°"])
        return {"total": total, "riesgo": 0, "honor": "SANTOS FIDEL (4.9)", "asistencia": "98%"}

    def _encontrar_hoja_prom(self, wb, grado, materia):
        materia_clean = materia.lower().replace(" ", "").replace(".", "")
        for sheet in wb.sheetnames:
            if "PROM" in sheet.upper():
                sheet_clean = sheet.lower().replace(" ", "").replace(".", "")
                if materia_clean in sheet_clean:
                    if self.modalidad == "primaria" or grado.replace("°","") in sheet:
                        return sheet
        return None

    def _obtener_rango_columnas(self, ws, trimestre, tipo_nota):
        textos_busqueda = {"Diaria / Parcial": "PARCIAL", "Apreciación": "APRECIACIÓN", "Examen": "PRUEBA"}
        texto_b = textos_busqueda[tipo_nota]
        ocurrencias = []
        for c in range(1, 200):
            val = ws.cell(row=3, column=c).value
            if val and texto_b in str(val).upper():
                ocurrencias.append(c)
        if not ocurrencias: return None, None
        idx_trimestre = int(trimestre.replace("Trimestre ", "")) - 1
        if idx_trimestre >= len(ocurrencias): return None, None
        col_inicio = ocurrencias[idx_trimestre]
        col_fin = col_inicio
        for c in range(col_inicio + 1, col_inicio + 25):
            val = ws.cell(row=3, column=c).value
            if val and ("PROMEDIO" in str(val).upper() or "CALIFICACIÓN" in str(val).upper() or "NOTAS" in str(val).upper() or "PRUEBAS" in str(val).upper()):
                break
            col_fin = c
        return col_inicio, col_fin

    def guardar_columna_notas(self, grado, materia, trimestre, tipo_nota, fecha, desc, dic_notas):
        if not os.path.exists(self.ruta): return False, "El archivo Excel no existe."
        wb = openpyxl.load_workbook(self.ruta)
        nombre_hoja = self._encontrar_hoja_prom(wb, grado, materia)
        if not nombre_hoja:
            wb.close()
            return False, f"No se encontró la hoja PROM para {materia} {grado}"
        ws = wb[nombre_hoja]
        rango = self._obtener_rango_columnas(ws, trimestre, tipo_nota)
        if rango == (None, None):
            wb.close()
            return False, "No se encontró el bloque en el Excel."
        col_inicio, col_fin = rango
        col_vacia = None
        columnas_ocupadas = 0
        for c in range(col_inicio, col_fin + 1):
            if not ws.cell(row=self.fila_desc, column=c).value:
                col_vacia = c
                break
            else:
                columnas_ocupadas += 1
        if not col_vacia:
            wb.close()
            if tipo_nota == "Examen": return False, "Límite alcanzado: Solo hay espacio para 2 exámenes."
            return False, "El bloque de notas está lleno en el Excel."
            
        if tipo_nota == "Examen": desc = f"Examen {columnas_ocupadas + 1} ({fecha})"
        try:
            ws.cell(row=4, column=col_vacia).value = fecha
        except AttributeError:
            if tipo_nota != "Examen": desc = f"{fecha} - {desc}"
        try:
            cell_desc = ws.cell(row=self.fila_desc, column=col_vacia)
            cell_desc.value = desc
            cell_desc.alignment = Alignment(textRotation=90, wrapText=True, horizontal='center', vertical='center')
        except AttributeError: pass
        for id_estudiante, nota in dic_notas.items():
            fila_excel = 4 + int(id_estudiante)
            try: ws.cell(row=fila_excel, column=col_vacia).value = nota
            except AttributeError: pass
        wb.save(self.ruta)
        wb.close()
        return True, ""

    def obtener_descripciones_notas(self, grado, materia, trimestre, tipo_nota):
        if not os.path.exists(self.ruta): return []
        wb = openpyxl.load_workbook(self.ruta, data_only=True)
        nombre_hoja = self._encontrar_hoja_prom(wb, grado, materia)
        if not nombre_hoja:
            wb.close()
            return []
        ws = wb[nombre_hoja]
        rango = self._obtener_rango_columnas(ws, trimestre, tipo_nota)
        if rango == (None, None):
            wb.close()
            return []
        col_inicio, col_fin = rango
        descripciones = []
        for c in range(col_inicio, col_fin + 1):
            try:
                val = ws.cell(row=self.fila_desc, column=c).value
                if val:
                    desc_limpia = str(val).replace('\n', ' ').strip()
                    descripciones.append(desc_limpia)
            except AttributeError: continue
        wb.close()
        return descripciones

    def buscar_notas_por_descripcion_exacta(self, grado, materia, trimestre, tipo_nota, descripcion):
        if not os.path.exists(self.ruta): return None
        wb = openpyxl.load_workbook(self.ruta, data_only=True)
        nombre_hoja = self._encontrar_hoja_prom(wb, grado, materia)
        if not nombre_hoja: 
            wb.close()
            return None
        ws = wb[nombre_hoja]
        rango = self._obtener_rango_columnas(ws, trimestre, tipo_nota)
        if rango == (None, None):
            wb.close()
            return None
        col_inicio, col_fin = rango
        col_encontrada = None
        for c in range(col_inicio, col_fin + 1):
            try:
                val = ws.cell(row=self.fila_desc, column=c).value
                if val and str(val).replace('\n', ' ').strip().lower() == descripcion.lower():
                    col_encontrada = c
                    break
            except AttributeError: continue
        if not col_encontrada:
            wb.close()
            return None
        notas = {}
        for r in range(5, 46):
            try:
                nota = ws.cell(row=r, column=col_encontrada).value
                if nota is not None: notas[r - 4] = nota
            except AttributeError: continue
        wb.close()
        return {"columna": col_encontrada, "notas": notas}

    def actualizar_notas_existentes(self, grado, materia, columna, dic_notas):
        if not os.path.exists(self.ruta): return False
        wb = openpyxl.load_workbook(self.ruta)
        nombre_hoja = self._encontrar_hoja_prom(wb, grado, materia)
        if not nombre_hoja: 
            wb.close()
            return False
        ws = wb[nombre_hoja]
        for id_estudiante, nota in dic_notas.items():
            fila_excel = 4 + int(id_estudiante)
            try: ws.cell(row=fila_excel, column=columna).value = nota
            except AttributeError: continue
        wb.save(self.ruta)
        wb.close()
        return True

    # =========================================================
    # NUEVO MÓDULO: ASISTENCIA 
    # =========================================================
    
    def _encontrar_hoja_asistencia(self, wb, grado):
        for sheet in wb.sheetnames:
            if "Asistencia" in sheet and (self.modalidad == "primaria" or grado.replace('°','') in sheet):
                return sheet
        return None

    def obtener_fechas_asistencia(self, grado, trimestre):
        if not os.path.exists(self.ruta): return []
        wb = openpyxl.load_workbook(self.ruta, data_only=True)
        hoja = self._encontrar_hoja_asistencia(wb, grado)
        if not hoja:
            wb.close()
            return []

        ws = wb[hoja]
        mapa_trimestres = {"Trimestre 1": 2, "Trimestre 2": 45, "Trimestre 3": 88}
        fila_fechas = mapa_trimestres.get(trimestre, 2)
        
        fechas = []
        for c in range(3, 61):
            val = ws.cell(row=fila_fechas, column=c).value
            if val: fechas.append(str(val).strip())
        wb.close()
        return fechas

    def buscar_asistencia_existente(self, grado, trimestre, fecha):
        if not os.path.exists(self.ruta): return None
        wb = openpyxl.load_workbook(self.ruta, data_only=True)
        hoja = self._encontrar_hoja_asistencia(wb, grado)
        if not hoja:
            wb.close()
            return None

        ws = wb[hoja]
        mapa_trimestres = {"Trimestre 1": 2, "Trimestre 2": 45, "Trimestre 3": 88}
        fila_fechas = mapa_trimestres.get(trimestre, 2)
        
        col_encontrada = None
        for c in range(3, 61):
            val = ws.cell(row=fila_fechas, column=c).value
            if val and str(val).strip() == fecha.strip():
                col_encontrada = c
                break
                
        if not col_encontrada:
            wb.close()
            return None

        asistencia = {}
        for id_est in range(1, 46): # Extrae del estudiante 1 al 45
            val = ws.cell(row=fila_fechas + id_est, column=col_encontrada).value
            if val is not None: asistencia[id_est] = val
            
        wb.close()
        return {"columna": col_encontrada, "asistencia": asistencia}

    def guardar_asistencia(self, grado, trimestre, fecha, dic_asistencia):
        if not os.path.exists(self.ruta): return False, "El archivo Excel no existe."
        wb = openpyxl.load_workbook(self.ruta)
        
        hoja = self._encontrar_hoja_asistencia(wb, grado)
        if not hoja:
            wb.close()
            return False, f"No se encontró la hoja de Asistencia para {grado}."
            
        ws = wb[hoja]
        mapa_trimestres = {"Trimestre 1": 2, "Trimestre 2": 45, "Trimestre 3": 88}
        fila_fechas = mapa_trimestres.get(trimestre, 2)
        
        col_vacia = None
        for c in range(3, 61):
            if not ws.cell(row=fila_fechas, column=c).value:
                col_vacia = c
                break
                
        if not col_vacia:
            wb.close()
            return False, "No hay más columnas vacías para este trimestre."
            
        ws.cell(row=fila_fechas, column=col_vacia).value = fecha
        
        # EL BLINDAJE DE FUENTE
        fuente_meduca = Font(name='Calibri', size=9, bold=True)
        
        for id_estudiante, datos in dic_asistencia.items():
            fila_excel = fila_fechas + int(id_estudiante) 
            celda = ws.cell(row=fila_excel, column=col_vacia)
            celda.value = datos["estado"]
            celda.font = fuente_meduca
            
        wb.save(self.ruta)
        wb.close()
        return True, ""

    def actualizar_asistencia(self, grado, trimestre, columna, dic_asistencia):
        if not os.path.exists(self.ruta): return False
        wb = openpyxl.load_workbook(self.ruta)
        hoja = self._encontrar_hoja_asistencia(wb, grado)
        if not hoja:
            wb.close()
            return False

        ws = wb[hoja]
        mapa_trimestres = {"Trimestre 1": 2, "Trimestre 2": 45, "Trimestre 3": 88}
        fila_fechas = mapa_trimestres.get(trimestre, 2)
        fuente_meduca = Font(name='Calibri', size=9, bold=True)
        
        for id_estudiante, datos in dic_asistencia.items():
            fila_excel = fila_fechas + int(id_estudiante)
            celda = ws.cell(row=fila_excel, column=columna)
            celda.value = datos["estado"]
            celda.font = fuente_meduca
            
        wb.save(self.ruta)
        wb.close()
        return True