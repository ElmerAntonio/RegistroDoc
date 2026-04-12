import os
import re
from rdsecurity import validar_nota_meduca
import openpyxl
from openpyxl.styles import Alignment, Font
from utils.cleaner import ExcelCleaner

class DataEngine:
    def __init__(self, ruta_excel, modalidad="premedia"):
        self.ruta = ruta_excel
        self.modalidad = modalidad.lower()
        self.cleaner = ExcelCleaner()
        self.fila_desc = 39 if self.modalidad == "primaria" else 42
        self._wb_cache = None
        self._cargar_en_memoria()


    def _safe_set_value(self, ws, row, column, value):
        celda = ws.cell(row=row, column=column)
        # Even if the type is not MergedCell, it could be the top-left cell of a merged range.
        for merged_range in list(ws.merged_cells.ranges):
            if celda.coordinate in merged_range:
                ws.unmerge_cells(str(merged_range))
                break
        ws.cell(row=row, column=column).value = value

    def _safe_clear_value(self, ws, row, column):
        celda = ws.cell(row=row, column=column)
        if type(celda).__name__ == 'MergedCell':
            return # Don't clear if it's merged, or we could unmerge. Let's just ignore or unmerge? The code was doing: if type != MergedCell: cell.value = None. So if it's Merged, just skip.
        if celda.data_type != 'f':
            celda.value = None

    def _cargar_en_memoria(self):
        if self._wb_cache is not None:
            try:
                self._wb_cache.close()
            except Exception:
                pass
            self._wb_cache = None

        if os.path.exists(self.ruta):
            try:
                self._wb_cache = openpyxl.load_workbook(self.ruta, data_only=True)
            except Exception:
                self._wb_cache = None

    def reiniciar_libreta(self):
        result = self.cleaner.limpiar_todo(self.ruta, self.modalidad)
        if result:
            self._cargar_en_memoria()
        return result

    def _obtener_columna_nombres(self, grado, wb=None):
        if not os.path.exists(self.ruta) and wb is None and self._wb_cache is None: return 2

        should_close = False
        if wb is None:
            should_close = not bool(self._wb_cache)
            wb = self._wb_cache if self._wb_cache else openpyxl.load_workbook(self.ruta, data_only=True)

        if "MAESTRO" not in wb.sheetnames: 
            if should_close: wb.close()
            return 2
        ws = wb["MAESTRO"]
        grado_limpio = grado.replace("°", "").strip()
        col = 2
        for r in [3, 4]:
            for c in range(1, 40):
                val = str(ws.cell(row=r, column=c).value or "").strip()
                if grado_limpio in val or grado in val:
                    if "NOMBRE" in str(ws.cell(row=4, column=c).value or "").upper(): col = c
                    elif "NOMBRE" in str(ws.cell(row=4, column=c+1).value or "").upper(): col = c+1
                    else: col = c
                    break
            if col != 2: break
        if should_close: wb.close()
        
        if col != 2: return col
        if self.modalidad == "primaria": return 2
        if "7" in grado: return 2
        if "8" in grado: return 4
        if "9" in grado: return 6
        return 2

    # --- LECTOR EXACTO DE CELDAS ---
    def obtener_datos_generales(self):
        datos = {
            "docente_nombre": "", "docente_cedula": "", "seguro_social": "", "numero_posicion": "",
            "condicion_nombramiento": "", "escuela_nombre": "", "escuela_region": "", "distrito": "",
            "corregimiento": "", "zona_escolar": "", "director_nombre": "", "subdirector_nombre": "",
            "coordinador_nombre": "", "telefono": "", "correo": "", "ano_lectivo": "2026",
            "jornada": "", "fecha_t1": "", "fecha_t2": "", "fecha_t3": ""
        }
        if not os.path.exists(self.ruta) and self._wb_cache is None: return datos
        should_close = not bool(self._wb_cache)
        wb = self._wb_cache if self._wb_cache else openpyxl.load_workbook(self.ruta, data_only=True)
        
        # 1. EXTRACCIÓN EN "PORTADA"
        if "Portada" in wb.sheetnames:
            ws = wb["Portada"]
            datos["docente_nombre"] = str(ws["P14"].value or ws["Q14"].value or ws["R14"].value or "").strip()
            datos["docente_cedula"] = str(ws["H16"].value or "").strip()
            datos["seguro_social"] = str(ws["AA16"].value or "").strip()
            datos["numero_posicion"] = str(ws["AQ16"].value or "").strip()
            datos["condicion_nombramiento"] = str(ws["U18"].value or "").strip()
            datos["jornada"] = str(ws["H20"].value or "").strip()
            datos["director_nombre"] = str(ws["S22"].value or "").strip()
            datos["subdirector_nombre"] = str(ws["T24"].value or "").strip()
            datos["coordinador_nombre"] = str(ws["O26"].value or "").strip()
            datos["escuela_nombre"] = str(ws["Q28"].value or "").strip()
            datos["telefono"] = str(ws["F30"].value or "").strip()
            datos["correo"] = str(ws["AF30"].value or "").strip()
            datos["escuela_region"] = str(ws["M32"].value or "").strip()
            datos["distrito"] = str(ws["L34"].value or "").strip()
            datos["corregimiento"] = str(ws["Q36"].value or "").strip()
            datos["zona_escolar"] = str(ws["O38"].value or "").strip()
            datos["ano_lectivo"] = str(ws["P8"].value or "2026").strip()

        # 2. EXTRACCIÓN DE FECHAS PURAS (C1, C44, C87) Limpiando guiones
        for sheet in wb.sheetnames:
            if "ASISTENCIA" in sheet.upper():
                ws = wb[sheet]
                def extraer_fecha(celda):
                    return DataEngine._procesar_texto_asistencia(ws[celda].value)
                
                datos["fecha_t1"] = extraer_fecha("C1")
                datos["fecha_t2"] = extraer_fecha("C44")
                datos["fecha_t3"] = extraer_fecha("C87")
                break
                
        if should_close: wb.close()
        return datos

    @staticmethod
    def _procesar_texto_asistencia(valor):
        """Lógica pura para extraer la fecha de una celda de asistencia."""
        if not valor:
            return ""
        val = str(valor).upper()
        if "ASISTENCIA DEL" in val:
            # Corta la palabra, limpia guiones y espacios dobles
            fecha = val.replace("ASISTENCIA DEL", "").replace("_", " ").strip()
            return " ".join(fecha.split())
        return ""

    def obtener_horario(self):
        horario = [{"horas": "", "lunes": "", "martes": "", "miercoles": "", "jueves": "", "viernes": ""} for _ in range(8)]
        if not os.path.exists(self.ruta) and self._wb_cache is None: return horario
        
        should_close = not bool(self._wb_cache)
        wb = self._wb_cache if self._wb_cache else openpyxl.load_workbook(self.ruta, data_only=True)
        hoja_horario = next((s for s in wb.sheetnames if "HORARIO" in s.upper()), None)
        
        if hoja_horario:
            ws = wb[hoja_horario]
            idx = 0
            
            # Buscamos en las filas 10 a la 25 para asegurarnos de no fallar
            for r in range(10, 25):
                if idx >= 8: break
                
                # Leemos el inicio de la fila para saber si es Receso
                fila_texto = "".join(str(ws.cell(row=r, column=c).value or "").upper() for c in range(1, 10))
                if "RECESO" in fila_texto: continue
                
                # Comprobamos si esta fila tiene un periodo (I, II, III...)
                es_periodo = False
                for c in range(1, 8):
                    if str(ws.cell(row=r, column=c).value or "").strip().upper() in ["I", "II", "III", "IV", "V", "VI", "VII", "VIII"]:
                        es_periodo = True
                        break
                
                if es_periodo:
                    # Función que atrapa el texto aunque la celda esté combinada mal
                    def atrapar_texto(col_base):
                        for offset in range(3): # Busca en la columna y 2 más a la derecha
                            v = str(ws.cell(row=r, column=col_base + offset).value or "").strip()
                            if v and v != "None": return v
                        return ""

                    # Extraemos con las coordenadas base que me diste
                    horario[idx]["horas"] = atrapar_texto(10)
                    horario[idx]["lunes"] = atrapar_texto(15)
                    horario[idx]["martes"] = atrapar_texto(16)
                    horario[idx]["miercoles"] = atrapar_texto(17)
                    horario[idx]["jueves"] = atrapar_texto(18)
                    horario[idx]["viernes"] = atrapar_texto(19)
                    idx += 1
                    
        if should_close: wb.close()
        return horario

    def guardar_horario(self, datos_horario):
        if not os.path.exists(self.ruta): return False
        wb = openpyxl.load_workbook(self.ruta)
        hoja_horario = next((s for s in wb.sheetnames if "HORARIO" in s.upper()), None)
        
        if hoja_horario:
            ws = wb[hoja_horario]
            idx = 0
            
            for r in range(10, 25):
                if idx >= len(datos_horario): break
                
                fila_texto = "".join(str(ws.cell(row=r, column=c).value or "").upper() for c in range(1, 10))
                if "RECESO" in fila_texto: continue
                
                es_periodo = False
                for c in range(1, 8):
                    if str(ws.cell(row=r, column=c).value or "").strip().upper() in ["I", "II", "III", "IV", "V", "VI", "VII", "VIII"]:
                        es_periodo = True
                        break
                
                if es_periodo:
                    h = datos_horario[idx]
                    # Escribimos exactamente en la primera celda de la combinación
                    try: ws.cell(row=r, column=10).value = h["horas"]
                    except: pass
                    try: ws.cell(row=r, column=15).value = h["lunes"]
                    except: pass
                    try: ws.cell(row=r, column=16).value = h["martes"]
                    except: pass
                    try: ws.cell(row=r, column=17).value = h["miercoles"]
                    except: pass
                    try: ws.cell(row=r, column=18).value = h["jueves"]
                    except: pass
                    try: ws.cell(row=r, column=19).value = h["viernes"]
                    except: pass
                    idx += 1
                    
            wb.save(self.ruta)
            wb.close()
            self._cargar_en_memoria()
            return True
        return False

    def sincronizar_plantilla_maestra(self, config):
        if not os.path.exists(self.ruta): return False
        wb = openpyxl.load_workbook(self.ruta)
        
        docente = str(config.get("docente_nombre", "DOCENTE")).upper()
        docente_titulo = str(config.get("docente_nombre", "Docente")).title()
        cedula = str(config.get("docente_cedula", ""))
        ss = str(config.get("seguro_social", ""))
        pos = str(config.get("numero_posicion", ""))
        condicion = str(config.get("condicion_nombramiento", "")).title()
        ano = str(config.get("ano_lectivo", "2026"))
        escuela = str(config.get("escuela_nombre", "ESCUELA")).upper()
        region = str(config.get("escuela_region", "REGIÓN")).title()
        distrito = str(config.get("distrito", "")).upper()
        corregimiento = str(config.get("corregimiento", "")).upper()
        zona = str(config.get("zona_escolar", ""))
        director = str(config.get("director_nombre", "")).title()
        subdirector = str(config.get("subdirector_nombre", "")).title()
        coordinador = str(config.get("coordinador_nombre", "")).title()
        jornada = str(config.get("jornada", "MATUTINA")).upper()
        telefono = str(config.get("telefono", ""))
        correo = str(config.get("correo", ""))
        ft1 = str(config.get("fecha_t1", "")).upper()
        ft2 = str(config.get("fecha_t2", "")).upper()
        ft3 = str(config.get("fecha_t3", "")).upper()

        grados_activos = self.obtener_grados_activos()
        str_grados = " - ".join(grados_activos)
        todas_materias = []
        for g in grados_activos: todas_materias.extend(self.obtener_materias_por_grado(g))
        str_materias = " - ".join(sorted(list(set(todas_materias))))

        # 1. INYECCIÓN EXACTA EN "PORTADA"
        if "Portada" in wb.sheetnames:
            ws = wb["Portada"]
            ws["R10"] = str_grados       # <--- ¡AGREGA ESTA LÍNEA AQUÍ!
            ws["P14"] = docente
            ws["H16"] = cedula
            ws["AA16"] = ss
            ws["AQ16"] = pos
            ws["U18"] = condicion
            ws["H20"] = jornada.capitalize()
            ws["AD20"] = str_materias
            ws["S22"] = director
            ws["T24"] = subdirector
            ws["O26"] = coordinador
            ws["L34"] = distrito
            ws["O38"] = zona
            ws["AF30"] = correo
            ws["Q28"] = escuela
            ws["F30"] = telefono
            ws["M32"] = region
            ws["Q36"] = corregimiento
            ws["J38"] = zona
            ws["P8"] = ano

        # 2. INYECCIÓN EXACTA EN "HORARIO"
        for sheet in wb.sheetnames:
            if "HORARIO" in sheet.upper():
                ws = wb[sheet]
                ws["O2"] = docente
                ws["H4"] = jornada
                ws["I6"] = str_materias
                ws["O8"] = str_grados

        # 3. CARÁTULA Y PORTADA VISTOSA 
        for sheet in wb.sheetnames:
            if "CARATULA" in sheet.upper() or "CARÁTULA" in sheet.upper() or "VISTOSA" in sheet.upper():
                ws = wb[sheet]
                for r in range(1, 40):
                    for c in range(1, 10):
                        try:
                            texto = str(ws.cell(row=r, column=c).value or "")
                            if not texto.strip(): continue
                            modificado = False
                            
                            if "Prof." in texto or "Profesor" in texto:
                                texto = re.sub(r'Prof\.\s*.*', f"Prof. {docente_titulo}", texto)
                                modificado = True
                            if "Instructor" in texto or "Pre-Media" in texto:
                                texto = f" {condicion}"
                                modificado = True
                            if "Grupo:" in texto:
                                texto = f"Grupo: {str_grados}"
                                modificado = True
                            if "Ano lectivo" in texto or "Año lectivo" in texto:
                                ws.cell(row=r, column=c+1).value = ano
                                
                            if "CERRO CACICON" in texto.upper(): texto = re.sub(r'CERRO CACICON|CERRO CACICÓN', escuela, texto, flags=re.IGNORECASE); modificado = True
                            if "ELMER TUGRI" in texto.upper(): texto = re.sub(r'ELMER TUGRI', docente_titulo, texto, flags=re.IGNORECASE); modificado = True
                            
                            if modificado: ws.cell(row=r, column=c).value = texto
                        except: pass

        # 4. INYECCIÓN DE FECHAS EN ASISTENCIA (Solo escribe en la celda)
        for sheet in wb.sheetnames:
            if "ASISTENCIA" in sheet.upper():
                ws = wb[sheet]
                if ft1: ws["C1"] = f"ASISTENCIA DEL {ft1}"
                if ft2: ws["C44"] = f"ASISTENCIA DEL {ft2}"
                if ft3: ws["C87"] = f"ASISTENCIA DEL {ft3}"
                
                for r in range(1, 6):
                    for c in range(1, 10):
                        try:
                            val = str(ws.cell(row=r, column=c).value or "").upper()
                            if "JORNADA:" in val and "ASIGNATURA:" in val:
                                ws.cell(row=r, column=c).value = f"JORNADA:           {jornada}                  ASIGNATURA:      Todas (Multigrado)"
                            if "AÑO:" in val and "AULA:" in val:
                                g_match = re.search(r'\((.*?)\)', sheet)
                                aula = g_match.group(1) if g_match else "Multigrado"
                                ws.cell(row=r, column=c).value = f"AÑO: {ano}     AULA:      {aula}                   NOMBRE DEL PROF./A CONSEJERO/A:      {docente}"
                        except: pass

        wb.save(self.ruta)
        wb.close()
        self._cargar_en_memoria()
        return True

    def obtener_grados_activos(self, wb=None):
        if not os.path.exists(self.ruta) and wb is None and self._wb_cache is None: return []

        should_close = False
        if wb is None:
            should_close = not bool(self._wb_cache)
            wb = self._wb_cache if self._wb_cache else openpyxl.load_workbook(self.ruta, data_only=True)

        grados = []
        for sheet in wb.sheetnames:
            if "Asistencia" in sheet and "(" in sheet and ")" in sheet:
                g = sheet.split("(")[1].split(")")[0].strip()
                grados.append(g)

        if should_close: wb.close()
        return sorted(list(set(grados))) if grados else []

    def obtener_materias_por_grado(self, grado):
        if not os.path.exists(self.ruta) and self._wb_cache is None: return []
        should_close = not bool(self._wb_cache)
        wb = self._wb_cache if self._wb_cache else openpyxl.load_workbook(self.ruta, data_only=True)
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
        if should_close: wb.close()
        return sorted(list(set(materias))) if materias else ["Sin materias registradas"]

    def obtener_estudiantes_completos(self, grado, wb=None):
        if not os.path.exists(self.ruta) and wb is None and self._wb_cache is None: return []

        should_close = False
        if wb is None:
            should_close = not bool(self._wb_cache)
            wb = self._wb_cache if self._wb_cache else openpyxl.load_workbook(self.ruta, data_only=True)

        ws_m = wb["MAESTRO"]
        col_nom = self._obtener_columna_nombres(grado, wb=wb)
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
                estudiantes.append({"id": r - 4, "nombre": str(nom).strip(), "cedula": str(cedula).strip() if cedula else ""})

        if should_close: wb.close()
        return estudiantes

    def agregar_estudiante(self, grado, nombre, cedula=""):
        if not os.path.exists(self.ruta): return False

        # Validar limites
        max_estudiantes = 34 if self.modalidad == "primaria" else 36
        actuales = self.obtener_estudiantes_completos(grado)
        if len(actuales) >= max_estudiantes:
            return False

        wb = openpyxl.load_workbook(self.ruta)
        ws_m = wb["MAESTRO"]
        col_nom = self._obtener_columna_nombres(grado)
        fila_vacia = None
        for r in range(5, 5 + max_estudiantes):
            if not ws_m.cell(row=r, column=col_nom).value:
                fila_vacia = r
                break
        if not fila_vacia:
            wb.close()
            return False 
        ws_m.cell(row=fila_vacia, column=col_nom).value = nombre
        if self.modalidad == "primaria": ws_m.cell(row=fila_vacia, column=5).value = cedula
        else:
            num_grado = grado.replace("°","")
            id_est = fila_vacia - 4
            for sheet in wb.sheetnames:
                if "Planilla" in sheet and num_grado in sheet:
                    ws_p = wb[sheet]
                    ws_p.cell(row=15+id_est, column=5).value = cedula
        wb.save(self.ruta)
        wb.close()
        self._cargar_en_memoria()
        return True

    def guardar_cambios_estudiantes(self, grado, datos_modificados):
        if not os.path.exists(self.ruta): return False
        wb = openpyxl.load_workbook(self.ruta)
        ws_m = wb["MAESTRO"]
        col_nom = self._obtener_columna_nombres(grado)
        for id_est, datos in datos_modificados.items():
            fila = 4 + int(id_est)
            ws_m.cell(row=fila, column=col_nom).value = datos["nombre"]
            if self.modalidad == "primaria": ws_m.cell(row=fila, column=5).value = datos["cedula"]
            else:
                num_grado = grado.replace("°","")
                for sheet in wb.sheetnames:
                    if "Planilla" in sheet and num_grado in sheet:
                        ws_p = wb[sheet]
                        ws_p.cell(row=15+int(id_est), column=5).value = datos["cedula"]
        wb.save(self.ruta)
        wb.close()
        self._cargar_en_memoria()
        return True

    def obtener_promedios_reales(self, grado, materia, trimestre):
        if not os.path.exists(self.ruta) and self._wb_cache is None: return {}
        should_close = not bool(self._wb_cache)
        wb = self._wb_cache if self._wb_cache else openpyxl.load_workbook(self.ruta, data_only=True)

        datos = {}
        grado_num = grado.replace("°", "")

        if trimestre == "Trimestre 1": col_b = "T.1"
        elif trimestre == "Trimestre 2": col_b = "T.2"
        elif trimestre == "Trimestre 3": col_b = "T.3"
        else: col_b = "ANUAL"

        hoja_res = None
        for s in wb.sheetnames:
            if "RESUMEN" in s.upper() and (self.modalidad == "primaria" or grado_num in s):
                hoja_res = s
                break

        if hoja_res:
            ws_res = wb[hoja_res]
            # Buscamos la fila de materias para encontrar la columna correcta de la materia y trimestre
            # En premedia, cada materia tiene columnas de T.1, T.2, T.3, PROMEDIO
            # Si materia="General", sacamos el promedio ANUAL final

            col_nom = None
            col_nota = None

            for c in range(1, 40):
                val = str(ws_res.cell(row=9, column=c).value or "").upper()
                if "NOMBRE" in val: col_nom = c

            if materia and materia not in ["Sin materias", "No hay materias", "General"]:
                # Buscar materia
                fila_materias = None
                col_inicio_materia = None
                for r in range(4, 15):
                    for c in range(2, 40):
                        if materia.upper() in str(ws_res.cell(row=r, column=c).value or "").upper():
                            fila_materias = r
                            col_inicio_materia = c
                            break
                    if col_inicio_materia: break

                if col_inicio_materia:
                    # Encontrar la columna del trimestre debajo de la materia
                    for c in range(col_inicio_materia, col_inicio_materia + 5):
                        val = str(ws_res.cell(row=fila_materias + 1, column=c).value or "").upper()
                        val2 = str(ws_res.cell(row=fila_materias + 2, column=c).value or "").upper()
                        if col_b in val or col_b in val2:
                            col_nota = c
                            break
            else:
                # General o todas las materias -> Anual
                for c in range(1, 40):
                    val = str(ws_res.cell(row=9, column=c).value or "").upper()
                    val2 = str(ws_res.cell(row=8, column=c).value or "").upper()
                    if "ANUAL" in val or "FINAL" in val or "ANUAL" in val2 or "FINAL" in val2:
                        col_nota = c
                        break

            if col_nom and col_nota:
                for r in range(10, 50):
                    nom = str(ws_res.cell(row=r, column=col_nom).value or "").strip()
                    if nom:
                        try:
                            valido, nota, _ = validar_nota_meduca(ws_res.cell(row=r, column=col_nota).value)
                            if valido:
                                datos[nom] = nota
                            else:
                                datos[nom] = 1.0
                        except (ValueError, TypeError):
                            datos[nom] = 1.0 # Default if empty or invalid

        if should_close: wb.close()
        return datos

    def obtener_historial_real(self, grado, materia, nombre_estudiante):
        if not os.path.exists(self.ruta) and self._wb_cache is None: return []
        should_close = not bool(self._wb_cache)
        wb = self._wb_cache if self._wb_cache else openpyxl.load_workbook(self.ruta, data_only=True)

        historial = []
        grado_num = grado.replace("°", "")

        hoja_res = None
        for s in wb.sheetnames:
            if "RESUMEN" in s.upper() and (self.modalidad == "primaria" or grado_num in s):
                hoja_res = s
                break

        if hoja_res:
            ws_res = wb[hoja_res]
            col_nom = None
            for c in range(1, 40):
                val = str(ws_res.cell(row=9, column=c).value or "").upper()
                if "NOMBRE" in val: col_nom = c

            fila_estudiante = None
            if col_nom:
                for r in range(10, 50):
                    if nombre_estudiante.upper() == str(ws_res.cell(row=r, column=col_nom).value or "").strip().upper():
                        fila_estudiante = r
                        break

            if fila_estudiante:
                if materia and materia not in ["Sin materias", "No hay materias", "General"]:
                    col_inicio_materia = None
                    for r in range(4, 15):
                        for c in range(2, 40):
                            if materia.upper() in str(ws_res.cell(row=r, column=c).value or "").upper():
                                col_inicio_materia = c
                                break
                        if col_inicio_materia: break

                    if col_inicio_materia:
                        # Buscar columnas de trimestre debajo de la materia
                        fila_materias = None
                        for rmat in range(4, 15):
                            if materia.upper() in str(ws_res.cell(row=rmat, column=col_inicio_materia).value or "").upper():
                                fila_materias = rmat
                                break
                        cols_trimestres = []
                        for c in range(col_inicio_materia, col_inicio_materia + 5):
                            val = str(ws_res.cell(row=fila_materias + 1, column=c).value or "").upper()
                            val2 = str(ws_res.cell(row=fila_materias + 2, column=c).value or "").upper()
                            if "T.1" in val or "T.1" in val2 or "T.2" in val or "T.2" in val2 or "T.3" in val or "T.3" in val2:
                                cols_trimestres.append(c)
                        for c in cols_trimestres:
                            try:
                                valido, nota, _ = validar_nota_meduca(ws_res.cell(row=fila_estudiante, column=c).value)
                                if valido:
                                    historial.append(nota)
                            except (ValueError, TypeError): pass
                else:
                    cols_promedios = []
                    for c in range(5, 40):
                        val = str(ws_res.cell(row=9, column=c).value or "").upper()
                        val2 = str(ws_res.cell(row=8, column=c).value or "").upper()
                        if "PROMEDIO" in val or "T.1" in val or "T.2" in val or "T.3" in val:
                            cols_promedios.append(c)

                    for c in cols_promedios:
                        try:
                            valido, nota, _ = validar_nota_meduca(ws_res.cell(row=fila_estudiante, column=c).value)
                            if valido:
                                historial.append(nota)
                        except (ValueError, TypeError): pass

        if should_close: wb.close()
        return historial if len(historial) >= 2 else [3.0, 3.0] # Fallback to avoid math errors in scipy

    def obtener_datos_reportes(self, grado):
        if not os.path.exists(self.ruta) and self._wb_cache is None: return {"docente": [], "aprobados": [], "direccion": []}
        should_close = not bool(self._wb_cache)
        wb = self._wb_cache if self._wb_cache else openpyxl.load_workbook(self.ruta, data_only=True)

        datos = {"docente": [], "aprobados": [], "direccion": []}
        grado_num = grado.replace("°", "")

        def extraer_de_hoja(palabra_clave):
            hoja_obj = None
            for s in wb.sheetnames:
                if palabra_clave in s.upper() and (self.modalidad == "primaria" or grado_num in s):
                    hoja_obj = wb[s]
                    break

            filas = []
            if hoja_obj:
                col_nom = None; col_ced = None; col_estado = None; col_anual = None
                for c in range(1, 40):
                    for r in [3, 4, 8, 9]:
                        val = str(hoja_obj.cell(row=r, column=c).value or "").upper()
                        if "NOMBRE" in val or "APELLIDO" in val: col_nom = c
                        if "CÉDULA" in val or "CEDULA" in val: col_ced = c
                        if "ESTADO" in val: col_estado = c
                        if "ANUAL" in val or "FINAL" in val: col_anual = c

                if col_nom:
                    for r in range(5, 50):
                        nom = hoja_obj.cell(row=r, column=col_nom).value
                        if nom:
                            ced = hoja_obj.cell(row=r, column=col_ced).value if col_ced else ""
                            estado = hoja_obj.cell(row=r, column=col_estado).value if col_estado else ""
                            anual = hoja_obj.cell(row=r, column=col_anual).value if col_anual else ""
                            filas.append([nom, ced, anual, estado])
            return filas

        # Si las hojas dedicadas existen, leemos de ahí, de lo contrario fallback a RESUMEN

        docente_data = extraer_de_hoja("REPORTE_DOCENTE")
        aprobados_data = extraer_de_hoja("REPORTE_APROBADOS")
        direccion_data = extraer_de_hoja("REPORTE_DIRECCIÓN")

        if not docente_data: docente_data = extraer_de_hoja("RESUMEN")
        if not aprobados_data: aprobados_data = extraer_de_hoja("RESUMEN")
        if not direccion_data: direccion_data = extraer_de_hoja("RESUMEN")

        for fila in docente_data: datos["docente"].append([fila[0], fila[1], fila[2]])
        for fila in aprobados_data: datos["aprobados"].append([fila[0], fila[3]])
        for fila in direccion_data: datos["direccion"].append(fila)

        if should_close: wb.close()
        return datos

    def get_dashboard_stats(self):
        if not os.path.exists(self.ruta) and self._wb_cache is None: return {"total": 0, "riesgo": 0, "honor": "N/A", "asistencia": "0%"}

        # Optimization: Use a single workbook load for all dashboard stats
        should_close = not bool(self._wb_cache)
        wb = self._wb_cache if self._wb_cache else openpyxl.load_workbook(self.ruta, data_only=True)
        try:
            grados = self.obtener_grados_activos(wb=wb)
            total = sum(len(self.obtener_estudiantes_completos(g, wb=wb)) for g in grados)
        finally:
            if should_close: wb.close()

        return {"total": total, "riesgo": 0, "honor": "SANTOS FIDEL (4.9)", "asistencia": "98%"}

    def _encontrar_hoja_prom(self, wb, grado, materia):
        materia_clean = materia.lower().replace(" ", "").replace(".", "")
        for sheet in wb.sheetnames:
            if "PROM" in sheet.upper():
                sheet_clean = sheet.lower().replace(" ", "").replace(".", "")
                if materia_clean in sheet_clean:
                    if self.modalidad == "primaria" or grado.replace("°","") in sheet: return sheet
        return None

    def _obtener_rango_columnas(self, ws, trimestre, tipo_nota):
        textos_busqueda = {"Diaria / Parcial": "PARCIAL", "Apreciación": "APRECIACIÓN", "Examen": "PRUEBA"}
        texto_b = textos_busqueda[tipo_nota]
        ocurrencias = []
        for c in range(1, 200):
            val = ws.cell(row=3, column=c).value
            if val and texto_b in str(val).upper(): ocurrencias.append(c)
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
            else: columnas_ocupadas += 1
        if not col_vacia:
            wb.close()
            if tipo_nota == "Examen": return False, "Límite alcanzado: Solo hay espacio para 2 exámenes."
            return False, "El bloque de notas está lleno en el Excel."
            
        if tipo_nota == "Examen": desc = f"Examen {columnas_ocupadas + 1} ({fecha})"
        try: ws.cell(row=4, column=col_vacia).value = fecha
        except AttributeError:
            if tipo_nota != "Examen": desc = f"{fecha} - {desc}"
        try:
            cell_desc = ws.cell(row=self.fila_desc, column=col_vacia)
            cell_desc.value = desc
            cell_desc.alignment = Alignment(wrapText=True, horizontal='center', vertical='center')
            cell_desc.font = Font(name='Calibri', size=11, bold=False)
        except AttributeError: pass
        for id_estudiante, nota in dic_notas.items():
            fila_excel = 4 + int(id_estudiante)
            try: ws.cell(row=fila_excel, column=col_vacia).value = nota
            except AttributeError: pass
        wb.save(self.ruta)
        wb.close()
        self._cargar_en_memoria()
        self.actualizar_resumen(grado)
        return True, ""

    def obtener_descripciones_notas(self, grado, materia, trimestre, tipo_nota):
        if not os.path.exists(self.ruta) and self._wb_cache is None: return []
        should_close = not bool(self._wb_cache)
        wb = self._wb_cache if self._wb_cache else openpyxl.load_workbook(self.ruta, data_only=True)
        nombre_hoja = self._encontrar_hoja_prom(wb, grado, materia)
        if not nombre_hoja:
            if should_close: wb.close()
            return []
        ws = wb[nombre_hoja]
        rango = self._obtener_rango_columnas(ws, trimestre, tipo_nota)
        if rango == (None, None):
            if should_close: wb.close()
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
        if should_close: wb.close()
        return descripciones

    def buscar_notas_por_descripcion_exacta(self, grado, materia, trimestre, tipo_nota, descripcion):
        if not os.path.exists(self.ruta) and self._wb_cache is None: return None
        should_close = not bool(self._wb_cache)
        wb = self._wb_cache if self._wb_cache else openpyxl.load_workbook(self.ruta, data_only=True)
        nombre_hoja = self._encontrar_hoja_prom(wb, grado, materia)
        if not nombre_hoja: 
            if should_close: wb.close()
            return None
        ws = wb[nombre_hoja]
        rango = self._obtener_rango_columnas(ws, trimestre, tipo_nota)
        if rango == (None, None):
            if should_close: wb.close()
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
            if should_close: wb.close()
            return None
        notas = {}
        for r in range(5, 46):
            try:
                nota = ws.cell(row=r, column=col_encontrada).value
                if nota is not None: notas[r - 4] = nota
            except AttributeError: continue
        if should_close: wb.close()
        return {"columna": col_encontrada, "notas": notas}

    def actualizar_notas_existentes(self, grado, materia, columna, dic_notas):
        if not os.path.exists(self.ruta): return False

        # Optimization: Use cache to find sheet name first, then open writeable workbook
        nombre_hoja = self._encontrar_hoja_prom(self._wb_cache, grado, materia) if self._wb_cache else None

        wb = openpyxl.load_workbook(self.ruta)
        if not nombre_hoja:
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
        self._cargar_en_memoria()
        self.actualizar_resumen(grado)
        return True

    def _encontrar_hoja_asistencia(self, wb, grado):
        for sheet in wb.sheetnames:
            if "Asistencia" in sheet and (self.modalidad == "primaria" or grado.replace('°','') in sheet):
                return sheet
        return None

    def obtener_fechas_asistencia(self, grado, trimestre):
        if not os.path.exists(self.ruta) and self._wb_cache is None: return []
        should_close = not bool(self._wb_cache)
        wb = self._wb_cache if self._wb_cache else openpyxl.load_workbook(self.ruta, data_only=True)
        hoja = self._encontrar_hoja_asistencia(wb, grado)
        if not hoja:
            if should_close: wb.close()
            return []
        ws = wb[hoja]
        mapa_trimestres = {"Trimestre 1": 2, "Trimestre 2": 45, "Trimestre 3": 88}
        fila_fechas = mapa_trimestres.get(trimestre, 2)
        fechas = []
        for c in range(3, 61):
            val = ws.cell(row=fila_fechas, column=c).value
            if val: fechas.append(str(val).strip())
        if should_close: wb.close()
        return fechas

    def buscar_asistencia_existente(self, grado, trimestre, fecha):
        if not os.path.exists(self.ruta) and self._wb_cache is None: return None
        should_close = not bool(self._wb_cache)
        wb = self._wb_cache if self._wb_cache else openpyxl.load_workbook(self.ruta, data_only=True)
        hoja = self._encontrar_hoja_asistencia(wb, grado)
        if not hoja:
            if should_close: wb.close()
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
            if should_close: wb.close()
            return None
        asistencia = {}
        for id_est in range(1, 46): 
            val = ws.cell(row=fila_fechas + id_est, column=col_encontrada).value
            if val is not None: asistencia[id_est] = val
        if should_close: wb.close()
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
        fuente_meduca = Font(name='Calibri', size=9, bold=True)
        for id_estudiante, datos in dic_asistencia.items():
            fila_excel = fila_fechas + int(id_estudiante) 
            celda = ws.cell(row=fila_excel, column=col_vacia)
            celda.value = datos["estado"]
            celda.font = fuente_meduca
        wb.save(self.ruta)
        wb.close()
        self._cargar_en_memoria()
        self.actualizar_resumen(grado)
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
        self._cargar_en_memoria()
        return True

    def actualizar_datos_generales(self, nombre_docente, ano_lectivo):
        if not os.path.exists(self.ruta): return False
        wb = openpyxl.load_workbook(self.ruta)
        if "MAESTRO" in wb.sheetnames:
            ws_m = wb["MAESTRO"]
            titulo_actual = str(ws_m.cell(row=1, column=1).value or "")
            nuevo_titulo = re.sub(r'20\d{2}', str(ano_lectivo), titulo_actual)
            self._safe_set_value(ws_m, 1, 1, nuevo_titulo)
        wb.save(self.ruta)
        wb.close()
        self._cargar_en_memoria()
        return True

    def obtener_consejero_actual(self, grado):
        if not os.path.exists(self.ruta) and self._wb_cache is None: return "No asignado"
        should_close = not bool(self._wb_cache)
        wb = self._wb_cache if self._wb_cache else openpyxl.load_workbook(self.ruta, data_only=True)
        grado_num = grado.replace("°", "")
        consejero = "No asignado"
        for sheet in wb.sheetnames:
            if "PLANILLA" in sheet.upper() and (self.modalidad == "primaria" or grado_num in sheet):
                ws = wb[sheet]
                for r in range(1, 15):
                    for c in range(1, 15):
                        val = str(ws.cell(row=r, column=c).value or "").upper()
                        if "CONSEJERO" in val or "CONSEJERA" in val:
                            if len(val) > 25: 
                                consejero = val.split(":")[-1].strip()
                            else:
                                consejero = str(ws.cell(row=r, column=c+2).value or "").strip()
                            if should_close: wb.close()
                            return consejero
        if should_close: wb.close()
        return consejero

    def actualizar_consejero(self, grado, nuevo_consejero):
        if not os.path.exists(self.ruta): return False
        wb = openpyxl.load_workbook(self.ruta)
        grado_num = grado.replace("°", "")
        for sheet in wb.sheetnames:
            sheet_upper = sheet.upper()
            if "PROM" in sheet_upper and (self.modalidad == "primaria" or grado_num in sheet):
                ws = wb[sheet]
                for r in range(1, 15):
                    for c in range(1, 15):
                        try:
                            celda = ws.cell(row=r, column=c)
                            val = str(celda.value or "").upper()
                            if "CONSEJERO" in val or "CONSEJERA" in val:
                                celda.value = f"PROF. CONSEJERO: {nuevo_consejero.upper()}"
                        except AttributeError: pass
            
            elif "PLANILLA" in sheet_upper and (self.modalidad == "primaria" or grado_num in sheet):
                ws = wb[sheet]
                for r in range(1, 15):
                    for c in range(1, 15):
                        try:
                            celda = ws.cell(row=r, column=c)
                            val = str(celda.value or "").upper()
                            if "CONSEJERO" in val or "CONSEJERA" in val:
                                if len(val) > 20: celda.value = f"PROF. CONSEJERO (A): {nuevo_consejero.upper()}"
                                else:
                                    ws.cell(row=r, column=c+1).value = nuevo_consejero.upper()
                                    ws.cell(row=r, column=c+2).value = nuevo_consejero.upper()
                        except AttributeError: pass
        wb.save(self.ruta)
        wb.close()
        self._cargar_en_memoria()
        return True

    def agregar_grado(self, nuevo_grado, consejero, jornada):
        if not os.path.exists(self.ruta): return False, "Archivo no encontrado"
        wb = openpyxl.load_workbook(self.ruta)
        hoja_base_asist = None
        for sheet in wb.sheetnames:
            if "Asistencia" in sheet:
                hoja_base_asist = sheet
                break
        if not hoja_base_asist:
            wb.close()
            return False, "No se encontró hoja de Asistencia base para clonar."

        nueva_hoja_asist = f"Asistencia ({nuevo_grado})"
        if nueva_hoja_asist in wb.sheetnames:
            wb.close()
            return False, f"El grado {nuevo_grado} ya existe."

        ws_nueva = wb.copy_worksheet(wb[hoja_base_asist])
        ws_nueva.title = nueva_hoja_asist
        for r in range(2, 60):
            for c in range(3, 100):
                try: ws_nueva.cell(row=r, column=c).value = None
                except AttributeError: pass

        if "MAESTRO" in wb.sheetnames:
            ws_m = wb["MAESTRO"]
            col_vacia = None
            for c in range(2, 40, 2):
                if not ws_m.cell(row=4, column=c).value:
                    col_vacia = c
                    break
            if col_vacia:
                self._safe_set_value(ws_m, 3, col_vacia, f"GRADO {nuevo_grado}")
                self._safe_set_value(ws_m, 4, col_vacia, "NOMBRES Y APELLIDOS")
                if self.modalidad == "premedia":
                    self._safe_set_value(ws_m, 4, col_vacia+1, "N° DE CÉDULA")
                    
        hoja_base_resumen = None
        for sheet in wb.sheetnames:
            if "RESUMEN" in sheet.upper():
                hoja_base_resumen = sheet
                break
                
        if hoja_base_resumen:
            ws_resumen = wb.copy_worksheet(wb[hoja_base_resumen])
            ws_resumen.title = f"RESUMEN ({nuevo_grado})"
            for r in range(4, 50):
                for c in range(1, 30):
                    try:
                        celda = ws_resumen.cell(row=r, column=c)
                        self._safe_clear_value(ws_resumen, r, c)
                    except AttributeError: pass
        
        wb.save(self.ruta)
        wb.close()
        self._cargar_en_memoria()
        return True, "Grado creado exitosamente."

    def eliminar_grado(self, grado):
        if not os.path.exists(self.ruta): return False
        wb = openpyxl.load_workbook(self.ruta)
        hojas_borrar = []
        grado_limpio = grado.replace("°", "")
        for sheet in wb.sheetnames:
            if f"({grado})" in sheet or f" {grado_limpio})" in sheet or f" {grado})" in sheet:
                hojas_borrar.append(sheet)

        if not hojas_borrar:
            wb.close()
            return False

        for h in hojas_borrar: del wb[h]

        if "MAESTRO" in wb.sheetnames:
            ws_m = wb["MAESTRO"]
            for c in range(1, 40):
                val = str(ws_m.cell(row=3, column=c).value or "")
                if grado in val:
                    for r in range(3, 51):
                        try:
                            self._safe_clear_value(ws_m, r, c)
                            self._safe_clear_value(ws_m, r, c+1)
                        except AttributeError: pass
                    break
        wb.save(self.ruta)
        wb.close()
        self._cargar_en_memoria()
        return True

    def clonar_materia(self, grado, materia_origen, nueva_materia, jornada):
        if not os.path.exists(self.ruta): return False, "Archivo no encontrado"
        wb = openpyxl.load_workbook(self.ruta)
        hoja_prom_origen = self._encontrar_hoja_prom(wb, grado, materia_origen)
        hoja_plan_origen = None
        grado_num = grado.replace("°", "")

        consejero = "No asignado"
        for sheet in wb.sheetnames:
            if "PLANILLA" in sheet.upper() and (self.modalidad == "primaria" or grado_num in sheet):
                ws = wb[sheet]
                for r in range(1, 15):
                    for c in range(1, 15):
                        val = str(ws.cell(row=r, column=c).value or "").upper()
                        if "CONSEJERO" in val or "CONSEJERA" in val:
                            if len(val) > 25: consejero = val.split(":")[-1].strip()
                            else: consejero = str(ws.cell(row=r, column=c+2).value or "").strip()
                            break

        if self.modalidad == "premedia":
            for sheet in wb.sheetnames:
                if "PLANILLA" in sheet.upper() and materia_origen.lower().replace(" ", "") in sheet.lower().replace(" ", "") and grado_num in sheet:
                    hoja_plan_origen = sheet
                    break

        if not hoja_prom_origen:
            wb.close()
            return False, "No se encontró materia base para clonar."

        nueva_prom = wb.copy_worksheet(wb[hoja_prom_origen])
        if self.modalidad == "primaria": nueva_prom.title = f"PROM ({nueva_materia.title()})"
        else: nueva_prom.title = f"PROM ({nueva_materia.title()} {grado})"
        
        for r in range(1, 15):
            for c in range(1, 15):
                try:
                    val = str(nueva_prom.cell(row=r, column=c).value or "").upper()
                    if "ASIGNATURA" in val and len(val) < 20: self._safe_set_value(nueva_prom, r, c, f"ASIGNATURA: {nueva_materia.upper()}")
                    if "CONSEJERO" in val and len(val) < 30: self._safe_set_value(nueva_prom, r, c, f"PROF. CONSEJERO: {consejero.upper()}")
                    if "AULA" in val and len(val) < 15: self._safe_set_value(nueva_prom, r, c, f"AULA: {grado}")
                    if "JORNADA" in val and len(val) < 20: self._safe_set_value(nueva_prom, r, c, f"JORNADA: {jornada.upper()}")
                except AttributeError: pass

        for r in range(4, 50):
            for c in range(3, 100):
                try:
                    celda = nueva_prom.cell(row=r, column=c)
                    self._safe_clear_value(nueva_prom, r, c)
                except AttributeError: pass

        if hoja_plan_origen:
            nueva_plan = wb.copy_worksheet(wb[hoja_plan_origen])
            nueva_plan.title = f"Planilla ({nueva_materia.title()} {grado})"
            for r in range(1, 15):
                for c in range(1, 10):
                    try:
                        val = str(nueva_plan.cell(row=r, column=c).value or "").upper()
                        if "ASIGNATURA:" in val: self._safe_set_value(nueva_plan, r, c+2, nueva_materia.upper())
                        if "GRUPO:" in val: self._safe_set_value(nueva_plan, r, c+2, grado)
                        if "CONSEJERO" in val:
                            if len(val) > 20: self._safe_set_value(nueva_plan, r, c, f"PROF. CONSEJERO (A): {consejero.upper()}")
                            else: self._safe_set_value(nueva_plan, r, c+2, consejero.upper())
                    except AttributeError: pass
            
            for r in range(15, 60):
                for c in range(5, 20):
                    try:
                        celda = nueva_plan.cell(row=r, column=c)
                        self._safe_clear_value(nueva_prom, r, c)
                    except AttributeError: pass

        hoja_resumen = None
        for sheet in wb.sheetnames:
            if "RESUMEN" in sheet.upper() and grado_num in sheet:
                hoja_resumen = sheet
                break
                
        if hoja_resumen:
            ws_res = wb[hoja_resumen]
            fila_materias = None
            for r in range(10, 20):
                for c in range(2, 15):
                    if materia_origen.upper() in str(ws_res.cell(row=r, column=c).value or "").upper():
                        fila_materias = r
                        break
                if fila_materias: break
                
            if fila_materias:
                for c in range(2, 20):
                    if not ws_res.cell(row=fila_materias, column=c).value:
                        self._safe_set_value(ws_res, fila_materias, c, nueva_materia.upper())
                        break

        wb.save(self.ruta)
        wb.close()
        self._cargar_en_memoria()
        return True, "Materia clonada y agregada al Resumen."


    def actualizar_resumen(self, grado):
        if not os.path.exists(self.ruta) and self._wb_cache is None: return False
        should_close = not bool(self._wb_cache)
        wb = self._wb_cache if self._wb_cache else openpyxl.load_workbook(self.ruta, data_only=True)
        hoja_res = None
        grado_num = grado.replace("°", "")
        for sheet in wb.sheetnames:
            if "RESUMEN" in sheet.upper() and (self.modalidad == "primaria" or grado_num in sheet):
                hoja_res = sheet
                break

        if not hoja_res:
            if should_close: wb.close()
            return False

        ws_res = wb[hoja_res]

        # Primero vamos a identificar las columnas de Promedio (T1, T2, T3) y Nota Final/Anual
        cols_promedios = []
        col_anual = None
        col_estado = None

        for c in range(5, 40):
            val = str(ws_res.cell(row=9, column=c).value or "").upper()
            val2 = str(ws_res.cell(row=8, column=c).value or "").upper()
            if "PROMEDIO" in val or "T.1" in val or "T.2" in val or "T.3" in val:
                cols_promedios.append(c)
            if "ANUAL" in val or "FINAL" in val or "ANUAL" in val2 or "FINAL" in val2:
                col_anual = c
            if "ESTADO" in val or "ESTADO" in val2:
                col_estado = c

        if col_anual and col_estado:
            wb_write = openpyxl.load_workbook(self.ruta)
            ws_write = wb_write[hoja_res]

            for r in range(10, 50):
                # Calcular anual
                notas = []
                for c in cols_promedios:
                    try:
                        valido, nota, _ = validar_nota_meduca(ws_res.cell(row=r, column=c).value)
                        if valido:
                            notas.append(nota)
                    except (ValueError, TypeError): pass

                if notas:
                    anual = sum(notas) / len(notas)
                    anual = round(anual, 1)
                    try:
                        ws_write.cell(row=r, column=col_anual).value = anual
                        ws_write.cell(row=r, column=col_estado).value = "Aprobado" if anual >= 3.0 else "Reprobado"
                    except AttributeError: pass

            wb_write.save(self.ruta)
            wb_write.close()

        if should_close: wb.close()
        self._cargar_en_memoria()
        return True

    def eliminar_materia(self, grado, materia):
        if not os.path.exists(self.ruta): return False
        wb = openpyxl.load_workbook(self.ruta)
        hojas_a_borrar = []
        materia_clean = materia.lower().replace(" ", "").replace(".", "")
        grado_num = grado.replace("°", "")

        for sheet in wb.sheetnames:
            sheet_clean = sheet.lower().replace(" ", "").replace(".", "")
            if materia_clean in sheet_clean and (self.modalidad == "primaria" or grado_num in sheet):
                if "PROM" in sheet.upper() or "PLANILLA" in sheet.upper():
                    hojas_a_borrar.append(sheet)

        if not hojas_a_borrar:
            wb.close()
            return False

        for hoja in hojas_a_borrar: del wb[hoja]

        hoja_resumen = None
        for sheet in wb.sheetnames:
            if "RESUMEN" in sheet.upper() and grado_num in sheet:
                hoja_resumen = sheet
                break
                
        if hoja_resumen:
            ws_res = wb[hoja_resumen]
            for r in range(10, 20):
                for c in range(2, 20):
                    if materia.upper() in str(ws_res.cell(row=r, column=c).value or "").upper():
                        self._safe_clear_value(ws_res, r, c)

        wb.save(self.ruta)
        wb.close()
        self._cargar_en_memoria()
        return True