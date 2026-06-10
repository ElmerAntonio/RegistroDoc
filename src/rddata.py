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
        self._wb_formulas_cache = None
        self._last_mtime = 0.0
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
        if self._wb_formulas_cache is not None:
            try:
                self._wb_formulas_cache.close()
            except Exception:
                pass
            self._wb_formulas_cache = None

        if os.path.exists(self.ruta):
            try:
                self._wb_cache = openpyxl.load_workbook(self.ruta, data_only=True)
                self._wb_formulas_cache = openpyxl.load_workbook(self.ruta, data_only=False)
                self._last_mtime = os.path.getmtime(self.ruta)
            except Exception:
                self._wb_cache = None
                self._wb_formulas_cache = None
                self._last_mtime = 0.0
    def _verificar_y_recargar_cache(self):
        if os.path.exists(self.ruta):
            try:
                current_mtime = os.path.getmtime(self.ruta)
                if current_mtime != self._last_mtime:
                    self._cargar_en_memoria()
            except Exception:
                pass

    def _parse_cell_ref(self, cell_ref):
        match = re.match(r"^([A-Z]+)([0-9]+)$", cell_ref)
        if not match:
            return 1, 1
        col_str = match.group(1)
        row = int(match.group(2))
        col = 0
        for char in col_str:
            col = col * 26 + (ord(char) - ord('A') + 1)
        return row, col

    def _resolver_celda(self, sheet_name, row, col, wb_data=None, wb_formulas=None):
        if wb_data is None:
            wb_data = self._wb_cache
        if wb_formulas is None:
            wb_formulas = self._wb_formulas_cache

        if not wb_data or not wb_formulas:
            return None

        if sheet_name not in wb_formulas or sheet_name not in wb_data:
            return None

        ws_f = wb_formulas[sheet_name]
        ws_d = wb_data[sheet_name]
        val = ws_f.cell(row=row, column=col).value
        if val is None:
            return None
        
        if not str(val).startswith("="):
            return ws_d.cell(row=row, column=col).value

        formula = str(val)
        
        # Strip IF wrapper if present e.g. =IF(MAESTRO!D5="","",...)
        if formula.startswith("=IF(") and ',"",' in formula:
            parts = formula.split(',"",', 1)
            actual_part = parts[1]
            if actual_part.endswith(")"):
                actual_part = actual_part[:-1]
            formula = "=" + actual_part

        # 1. Resolve sheet references inside IFERROR or direct sheet references
        ref_match = re.search(r"IFERROR\('?([^']+)'?!([A-Z]+[0-9]+)", formula)
        if ref_match:
            target_sheet = ref_match.group(1)
            target_cell = ref_match.group(2)
            t_row, t_col = self._parse_cell_ref(target_cell)
            return self._resolver_celda(target_sheet, t_row, t_col, wb_data, wb_formulas)

        direct_ref = re.search(r"([A-Za-z0-9_]+)!([A-Z]+[0-9]+)", formula)
        if direct_ref:
            target_sheet = direct_ref.group(1)
            target_cell = direct_ref.group(2)
            t_row, t_col = self._parse_cell_ref(target_cell)
            return self._resolver_celda(target_sheet, t_row, t_col, wb_data, wb_formulas)

        # 2. Resolve average formulas in PROM sheets
        if "AVERAGE" in formula or "PROMEDIO" in formula:
            args_match = re.search(r"AVERAGE\(([^)]+)\)", formula)
            if args_match:
                args_str = args_match.group(1)
                parts = [p.strip() for p in args_str.split(",")]
                vals = []
                for part in parts:
                    if ":" in part:
                        start_cell, end_cell = part.split(":")
                        r_start, c_start = self._parse_cell_ref(start_cell)
                        r_end, c_end = self._parse_cell_ref(end_cell)
                        for r_idx in range(r_start, r_end + 1):
                            for c_idx in range(c_start, c_end + 1):
                                val_d = self._resolver_celda(sheet_name, r_idx, c_idx, wb_data, wb_formulas)
                                if val_d is not None and isinstance(val_d, (int, float)):
                                    vals.append(val_d)
                    else:
                        r_ref, c_ref = self._parse_cell_ref(part)
                        val_d = self._resolver_celda(sheet_name, r_ref, c_ref, wb_data, wb_formulas)
                        if val_d is not None and isinstance(val_d, (int, float)):
                            vals.append(val_d)
                
                if vals:
                    avg = sum(vals) / len(vals)
                    if "TRUNC" in formula:
                        return int(avg * 10) / 10.0
                    return round(avg, 2)
                return None

        # Fallback
        return ws_d.cell(row=row, column=col).value

    def _save_wb(self, wb):
        import shutil
        import tempfile
        import datetime

        # 1. Backup de seguridad antes de guardar
        bak_ruta = self.ruta.replace(".xlsx", "_bak.xlsx")
        if os.path.exists(self.ruta):
            try:
                shutil.copy2(self.ruta, bak_ruta)
                
                # Respaldo histórico local automatizado (offline)
                dir_base = os.path.dirname(os.path.abspath(self.ruta))
                dir_respaldos = os.path.join(dir_base, "Respaldos_Locales")
                if not os.path.exists(dir_respaldos):
                    os.makedirs(dir_respaldos)
                
                # Crear nombre con timestamp
                ahora = datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
                nombre_base = os.path.basename(self.ruta).replace(".xlsx", "")
                nombre_respaldo = f"{nombre_base}_respaldo_{ahora}.xlsx"
                shutil.copy2(self.ruta, os.path.join(dir_respaldos, nombre_respaldo))
                
                # Auto-limpieza: mantener solo los últimos 15 respaldos
                respaldos = sorted(
                    [os.path.join(dir_respaldos, f) for f in os.listdir(dir_respaldos) if f.endswith(".xlsx")],
                    key=os.path.getmtime
                )
                while len(respaldos) > 15:
                    viejo = respaldos.pop(0)
                    try: os.remove(viejo)
                    except: pass
            except Exception:
                pass

        # 2. Guardar a través de archivo temporal de forma atómica
        dir_base = os.path.dirname(os.path.abspath(self.ruta))
        fd, temp_path = tempfile.mkstemp(suffix=".xlsx", dir=dir_base)
        os.close(fd)

        try:
            wb.save(temp_path)
            # Reemplazar de forma atómica
            if os.path.exists(self.ruta):
                os.replace(temp_path, self.ruta)
            else:
                shutil.move(temp_path, self.ruta)
        except Exception as e:
            if os.path.exists(temp_path):
                try: os.remove(temp_path)
                except: pass
            raise e

        # Actualizar mtime para evitar recarga de caché redundante por nuestra propia escritura
        try:
            self._last_mtime = os.path.getmtime(self.ruta)
        except Exception:
            self._last_mtime = 0.0


    def reiniciar_libreta(self):
        result = self.cleaner.limpiar_todo(self.ruta, self.modalidad)
        if result:
            self._cargar_en_memoria()
        return result

    def _encontrar_hoja_maestro(self, wb):
        if "MAESTRO" in wb.sheetnames:
            return "MAESTRO"
        for s in wb.sheetnames:
            if "MAESTRO" in s.upper():
                return s
        return wb.sheetnames[0] if wb.sheetnames else "MAESTRO"

    def _obtener_columna_nombres(self, grado, wb=None):
        if wb is None:
            self._verificar_y_recargar_cache()
        if not os.path.exists(self.ruta) and wb is None and self._wb_cache is None: return 2
        should_close = not bool(self._wb_cache) if wb is None else False
        if wb is None:
            wb = self._wb_cache if self._wb_cache else openpyxl.load_workbook(self.ruta, data_only=True)

        nombre_maestro = self._encontrar_hoja_maestro(wb)
        ws = wb[nombre_maestro]
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
    def obtener_datos_generales(self, wb=None):
        if wb is None:
            self._verificar_y_recargar_cache()
        datos = {
            "docente_nombre": "", "docente_cedula": "", "seguro_social": "", "numero_posicion": "",
            "condicion_nombramiento": "", "escuela_nombre": "", "escuela_region": "", "distrito": "",
            "corregimiento": "", "zona_escolar": "", "director_nombre": "", "subdirector_nombre": "",
            "coordinador_nombre": "", "telefono": "", "correo": "", "ano_lectivo": "2026",
            "jornada": "", "fecha_t1": "", "fecha_t2": "", "fecha_t3": ""
        }
        if not os.path.exists(self.ruta) and wb is None and self._wb_cache is None: return datos
        should_close = not bool(self._wb_cache) if wb is None else False
        if wb is None:
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

    def obtener_horario(self, wb=None):
        if wb is None:
            self._verificar_y_recargar_cache()
        horario = [{"horas": "", "lunes": "", "martes": "", "miercoles": "", "jueves": "", "viernes": ""} for _ in range(8)]
        if not os.path.exists(self.ruta) and wb is None and self._wb_cache is None: return horario
        should_close = not bool(self._wb_cache) if wb is None else False
        if wb is None:
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
                    mapeo_columnas = {10: "horas", 15: "lunes", 16: "martes", 17: "miercoles", 18: "jueves", 19: "viernes"}
                    for col, llave in mapeo_columnas.items():
                        try:
                            ws.cell(row=r, column=col).value = h[llave]
                        except Exception as e:
                            print(f"[!] Error al escribir horario ({llave}) en fila {r}: {e}")
                    idx += 1
                    
            self._save_wb(wb)
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
                                
                            if "CERRO CACICON" in texto.upper() or "ESCUELA DE EJEMPLO" in texto.upper():
                                texto = re.sub(r'CERRO CACICON|CERRO CACICÓN|ESCUELA DE EJEMPLO', escuela, texto, flags=re.IGNORECASE)
                                modificado = True
                            if "ELMER TUGRI" in texto.upper() or "JUAN PÉREZ" in texto.upper() or "JUAN PEREZ" in texto.upper():
                                texto = re.sub(r'ELMER TUGRI|JUAN PÉREZ|JUAN PEREZ', docente_titulo, texto, flags=re.IGNORECASE)
                                modificado = True
                            
                            if modificado: ws.cell(row=r, column=c).value = texto
                        except Exception as e:
                            print(f"[!] Error in sincronizar_plantilla_maestra (carátula) at row {r}, col {c}: {e}")

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
                        except Exception as e:
                            print(f"[!] Error in sincronizar_plantilla_maestra (asistencia) at row {r}, col {c}: {e}")

        self._save_wb(wb)
        wb.close()
        self._cargar_en_memoria()
        return True

    def obtener_grados_activos(self, wb=None):
        if wb is None:
            self._verificar_y_recargar_cache()
        if not os.path.exists(self.ruta) and wb is None and self._wb_cache is None: return []
        should_close = not bool(self._wb_cache) if wb is None else False
        if wb is None:
            wb = self._wb_cache if self._wb_cache else openpyxl.load_workbook(self.ruta, data_only=True)

        grados = []
        for sheet in wb.sheetnames:
            if "Asistencia" in sheet and "(" in sheet and ")" in sheet:
                g = sheet.split("(")[1].split(")")[0].strip()
                grados.append(g)

        if should_close: wb.close()
        return sorted(list(set(grados))) if grados else []

    def obtener_materias_por_grado(self, grado, wb=None):
        if wb is None:
            self._verificar_y_recargar_cache()
        if not os.path.exists(self.ruta) and wb is None and self._wb_cache is None: return []
        should_close = not bool(self._wb_cache) if wb is None else False
        if wb is None:
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
        if wb is None:
            self._verificar_y_recargar_cache()
        if not os.path.exists(self.ruta) and wb is None and self._wb_cache is None: return []
        should_close = not bool(self._wb_cache) if wb is None else False
        if wb is None:
            wb = self._wb_cache if self._wb_cache else openpyxl.load_workbook(self.ruta, data_only=True)

        nombre_maestro = self._encontrar_hoja_maestro(wb)
        ws_m = wb[nombre_maestro]
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
        ws_m = wb[self._encontrar_hoja_maestro(wb)]
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
        self._save_wb(wb)
        wb.close()
        self._cargar_en_memoria()
        return True

    def guardar_cambios_estudiantes(self, grado, datos_modificados):
        if not os.path.exists(self.ruta): return False
        wb = openpyxl.load_workbook(self.ruta)
        ws_m = wb[self._encontrar_hoja_maestro(wb)]
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
        self._save_wb(wb)
        wb.close()
        self._cargar_en_memoria()
        return True

    def limpiar_acentos(self, texto):
        if not texto:
            return ""
        texto = str(texto).upper()
        for k, v in {"Á":"A", "É":"E", "Í":"I", "Ó":"O", "Ú":"U", "Ü":"U"}.items():
            texto = texto.replace(k, v)
        return texto

    def obtener_promedios_reales(self, grado, materia, trimestre, wb=None):
        if wb is None:
            self._verificar_y_recargar_cache()
        if not os.path.exists(self.ruta) and wb is None and self._wb_cache is None: return {}
        should_close = not bool(self._wb_cache) if wb is None else False
        if wb is None:
            wb = self._wb_cache if self._wb_cache else openpyxl.load_workbook(self.ruta, data_only=True)

        datos = {}
        grado_num = grado.replace("°", "")

        # Normalizar trimestre a variantes
        if trimestre == "Trimestre 1":
            col_b_variants = ["T.1", "T1", "T-1"]
        elif trimestre == "Trimestre 2":
            col_b_variants = ["T.2", "T2", "T-2"]
        elif trimestre == "Trimestre 3":
            col_b_variants = ["T.3", "T3", "T-3"]
        else:
            col_b_variants = ["ANUAL", "FINAL", "PROMEDIO"]

        hoja_res = None
        for s in wb.sheetnames:
            if "RESUMEN" in s.upper() and (self.modalidad == "primaria" or grado_num in s):
                hoja_res = s
                break

        if hoja_res:
            ws_res = wb[hoja_res]
            col_nom = None
            for c in range(1, 40):
                val = str(ws_res.cell(row=3, column=c).value or "").upper()
                if "NOMBRE" in val:
                    col_nom = c
                    break
            if not col_nom:
                col_nom = 2

            col_nota = None
            estudiantes = self.obtener_estudiantes_completos(grado, wb=wb)

            if materia and materia not in ["Sin materias", "No hay materias", "General", "Todas las Materias"]:
                # Buscar materia
                fila_materias = None
                col_inicio_materia = None
                for r in range(3, 15):
                    for c in range(2, 40):
                        val_cell = str(ws_res.cell(row=r, column=c).value or "").upper()
                        if self.limpiar_acentos(materia) in self.limpiar_acentos(val_cell):
                            fila_materias = r
                            col_inicio_materia = c
                            break
                    if col_inicio_materia: break

                if col_inicio_materia:
                    # Encontrar la columna del trimestre debajo de la materia
                    for c in range(col_inicio_materia, col_inicio_materia + 5):
                        val = str(ws_res.cell(row=fila_materias + 1, column=c).value or "").upper()
                        val2 = str(ws_res.cell(row=fila_materias + 2, column=c).value or "").upper()
                        if any(variant in val or variant in val2 for variant in col_b_variants):
                            col_nota = c
                            break
                    
                    if col_nota:
                        for est in estudiantes:
                            r = est["id"] + 4
                            val = self._resolver_celda(hoja_res, r, col_nota, wb_data=wb, wb_formulas=self._wb_formulas_cache)
                            valido, nota, _ = validar_nota_meduca(val)
                            if valido:
                                datos[est["nombre"]] = nota
            else:
                # Caso materia general: promedio de todas las materias del trimestre
                materias_grado = self.obtener_materias_por_grado(grado, wb=wb)
                materias_limpias = [m for m in materias_grado if m and m not in ["Sin materias", "No hay materias", "General", "Todas las Materias"]]
                
                if materias_limpias:
                    bulk_data = self.obtener_promedios_reales_bulk(grado, materias_limpias, trimestre, wb=wb)
                    
                    # Identificar materias de tecnología
                    tech_mats = []
                    norm_mats = []
                    for m in materias_limpias:
                        m_upper = self.limpiar_acentos(m)
                        if "HOGAR" in m_upper or "DESARROLLO" in m_upper or "AGRO" in m_upper or "COMERCIO" in m_upper:
                            tech_mats.append(m)
                        else:
                            norm_mats.append(m)
                            
                    estudiantes = self.obtener_estudiantes_completos(grado, wb=wb)
                    for est in estudiantes:
                        nom = est["nombre"]
                        norm_notas = []
                        for m in norm_mats:
                            if m in bulk_data and nom in bulk_data[m] and bulk_data[m][nom] is not None:
                                norm_notas.append(bulk_data[m][nom])
                                
                        tech_notas = []
                        for m in tech_mats:
                            if m in bulk_data and nom in bulk_data[m] and bulk_data[m][nom] is not None:
                                tech_notas.append(bulk_data[m][nom])
                                
                        tec_prom = sum(tech_notas) / len(tech_notas) if tech_notas else None
                        
                        comb_notas = list(norm_notas)
                        if tec_prom is not None:
                            comb_notas.append(tec_prom)
                            
                        if comb_notas:
                            datos[nom] = round(sum(comb_notas) / len(comb_notas), 2)
                
                # Fallback a la columna ANUAL/FINAL si no hay materias cargadas o falló
                if not datos:
                    col_nota = None
                    for c in range(1, 40):
                        val = str(ws_res.cell(row=3, column=c).value or "").upper()
                        val2 = str(ws_res.cell(row=4, column=c).value or "").upper()
                        if any(variant in val or variant in val2 for variant in ["ANUAL", "FINAL", "PROMEDIO"]):
                            col_nota = c
                            break
                    if col_nota:
                        for est in estudiantes:
                            r = est["id"] + 4
                            val = self._resolver_celda(hoja_res, r, col_nota, wb_data=wb, wb_formulas=self._wb_formulas_cache)
                            valido, nota, _ = validar_nota_meduca(val)
                            if valido:
                                datos[est["nombre"]] = nota

        if should_close: wb.close()
        return datos

    def obtener_notas_estudiante(self, nombre_estudiante, grado, trimestre, wb=None):
        """
        Obtiene las notas finales (promedios) de cada materia para un estudiante en un trimestre.
        trimestre: int (1, 2, 3) o str ("Trimestre 1", etc.)
        Retorna: dict {materia_nombre: nota_valor}
        """
        if wb is None:
            self._verificar_y_recargar_cache()
            
        nombre_est_clean = nombre_estudiante.strip().lower()
        trim_str = f"Trimestre {trimestre}" if isinstance(trimestre, int) else trimestre
        
        materias = self.obtener_materias_por_grado(grado, wb=wb)
        notas_estudiante = {}
        
        for mat in materias:
            if mat in ["Sin materias registradas", "Sin materias", "No hay materias", "General"]:
                continue
            proms = self.obtener_promedios_reales(grado, mat, trim_str, wb=wb)
            for est_nom, nota in proms.items():
                if est_nom.strip().lower() == nombre_est_clean:
                    notas_estudiante[mat] = nota
                    break
        return notas_estudiante

    def obtener_promedios_reales_bulk(self, grado, materias, trimestre, wb=None):
        if wb is None:
            self._verificar_y_recargar_cache()
        if not os.path.exists(self.ruta) and wb is None and self._wb_cache is None: return {}
        should_close = not bool(self._wb_cache) if wb is None else False
        if wb is None:
            wb = self._wb_cache if self._wb_cache else openpyxl.load_workbook(self.ruta, data_only=True)

        resultados = {m: {} for m in materias}
        grado_num = grado.replace("°", "")

        if trimestre == "Trimestre 1":
            col_b_variants = ["T.1", "T1", "T-1"]
        elif trimestre == "Trimestre 2":
            col_b_variants = ["T.2", "T2", "T-2"]
        elif trimestre == "Trimestre 3":
            col_b_variants = ["T.3", "T3", "T-3"]
        else:
            col_b_variants = ["ANUAL", "FINAL", "PROMEDIO"]

        hoja_res = None
        for s in wb.sheetnames:
            if "RESUMEN" in s.upper() and (self.modalidad == "primaria" or grado_num in s):
                hoja_res = s
                break

        if hoja_res:
            ws_res = wb[hoja_res]
            col_nom = None
            for c in range(1, 40):
                val = str(ws_res.cell(row=3, column=c).value or "").upper()
                if "NOMBRE" in val:
                    col_nom = c
                    break
            if not col_nom:
                col_nom = 2

            materia_to_col = {}
            for materia in materias:
                if materia and materia not in ["Sin materias", "No hay materias", "General"]:
                    fila_materias = None
                    col_inicio_materia = None
                    for r in range(3, 15):
                        for c in range(2, 40):
                            val_cell = str(ws_res.cell(row=r, column=c).value or "").upper()
                            if self.limpiar_acentos(materia) in self.limpiar_acentos(val_cell):
                                fila_materias = r
                                col_inicio_materia = c
                                break
                        if col_inicio_materia: break

                    if col_inicio_materia:
                        for c in range(col_inicio_materia, col_inicio_materia + 5):
                            val = str(ws_res.cell(row=fila_materias + 1, column=c).value or "").upper()
                            val2 = str(ws_res.cell(row=fila_materias + 2, column=c).value or "").upper()
                            if any(variant in val or variant in val2 for variant in col_b_variants):
                                materia_to_col[materia] = c
                                break

            estudiantes = self.obtener_estudiantes_completos(grado, wb=wb)
            for est in estudiantes:
                r = est["id"] + 4
                for materia, col_nota in materia_to_col.items():
                    if col_nota:
                        val = self._resolver_celda(hoja_res, r, col_nota, wb_data=wb, wb_formulas=self._wb_formulas_cache)
                        valido, nota, _ = validar_nota_meduca(val)
                        if valido:
                            resultados[materia][est["nombre"]] = nota

        if should_close: wb.close()
        return resultados

    def obtener_historial_real(self, grado, materia, nombre_estudiante, wb=None):
        if wb is None:
            self._verificar_y_recargar_cache()
        if not os.path.exists(self.ruta) and wb is None and self._wb_cache is None: return []
        should_close = not bool(self._wb_cache) if wb is None else False
        if wb is None:
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
                val = str(ws_res.cell(row=3, column=c).value or "").upper()
                if "NOMBRE" in val:
                    col_nom = c
                    break
            if not col_nom:
                col_nom = 2

            # Encontrar el ID/fila del estudiante a través de obtener_estudiantes_completos
            estudiantes = self.obtener_estudiantes_completos(grado, wb=wb)
            fila_estudiante = None
            for est in estudiantes:
                if est["nombre"].upper() == nombre_estudiante.strip().upper():
                    fila_estudiante = est["id"] + 4
                    break

            if fila_estudiante:
                if materia and materia not in ["Sin materias", "No hay materias", "General"]:
                    col_inicio_materia = None
                    for r in range(3, 15):
                        for c in range(2, 40):
                            val_cell = str(ws_res.cell(row=r, column=c).value or "").upper()
                            if self.limpiar_acentos(materia) in self.limpiar_acentos(val_cell):
                                col_inicio_materia = c
                                break
                        if col_inicio_materia: break

                    if col_inicio_materia:
                        # Buscar columnas de trimestre debajo de la materia
                        fila_materias = None
                        for rmat in range(3, 15):
                            val_cell = str(ws_res.cell(row=rmat, column=col_inicio_materia).value or "").upper()
                            if self.limpiar_acentos(materia) in self.limpiar_acentos(val_cell):
                                fila_materias = rmat
                                break
                        cols_trimestres = []
                        for c in range(col_inicio_materia, col_inicio_materia + 5):
                            val = str(ws_res.cell(row=fila_materias + 1, column=c).value or "").upper()
                            val2 = str(ws_res.cell(row=fila_materias + 2, column=c).value or "").upper()
                            if any(v in val or v in val2 for v in ["T.1", "T1", "T-1", "T.2", "T2", "T-2", "T.3", "T3", "T-3"]):
                                cols_trimestres.append(c)
                        for c in cols_trimestres:
                            try:
                                val = self._resolver_celda(hoja_res, fila_estudiante, c, wb_data=wb, wb_formulas=self._wb_formulas_cache)
                                valido, nota, _ = validar_nota_meduca(val)
                                if valido:
                                    historial.append(nota)
                            except (ValueError, TypeError): pass
                else:
                    cols_promedios = []
                    for c in range(2, 60):
                        val = str(ws_res.cell(row=3, column=c).value or "").upper()
                        val2 = str(ws_res.cell(row=4, column=c).value or "").upper()
                        if any(v in val or v in val2 for v in ["T.1", "T1", "T-1", "T.2", "T2", "T-2", "T.3", "T3", "T-3", "PROMEDIO", "ANUAL", "FINAL"]):
                            if c not in cols_promedios:
                                cols_promedios.append(c)

                    for c in cols_promedios:
                        try:
                            val = self._resolver_celda(hoja_res, fila_estudiante, c, wb_data=wb, wb_formulas=self._wb_formulas_cache)
                            valido, nota, _ = validar_nota_meduca(val)
                            if valido:
                                historial.append(nota)
                        except (ValueError, TypeError): pass

        if should_close: wb.close()
        return historial if len(historial) >= 2 else [3.0, 3.0] # Fallback to avoid math errors in scipy

    def obtener_datos_reportes(self, grado, wb=None):
        if wb is None:
            self._verificar_y_recargar_cache()
        if not os.path.exists(self.ruta) and wb is None and self._wb_cache is None:
            return {"docente": [], "aprobados": [], "direccion": []}
        should_close = not bool(self._wb_cache) if wb is None else False
        if wb is None:
            wb = self._wb_cache if self._wb_cache else openpyxl.load_workbook(self.ruta, data_only=True)

        datos = {"docente": [], "aprobados": [], "direccion": []}
        try:
            estudiantes = self.obtener_estudiantes_completos(grado, wb=wb)
            proms = self.obtener_promedios_reales(grado, None, "Anual", wb=wb)
            if not proms:
                proms = self.obtener_promedios_reales(grado, None, "Trimestre 1", wb=wb)

            for est in estudiantes:
                nom = est["nombre"]
                ced = est.get("cedula", "")
                prom_val = proms.get(nom, None)
                
                # Format average
                if prom_val is not None:
                    prom_str = f"{prom_val:.2f}"
                    estado = "APROBADO" if prom_val >= 3.0 else "REPROBADO"
                else:
                    prom_str = "—"
                    estado = "SIN NOTAS"

                datos["docente"].append([nom, ced, prom_str])
                datos["aprobados"].append([nom, estado])
                datos["direccion"].append([nom, ced, prom_str, estado])
        except Exception as e:
            print(f"Error generando datos de reportes: {e}")
        finally:
            if should_close: wb.close()

        return datos

    def obtener_cuadro_honor_general(self, wb=None):
        if wb is None:
            self._verificar_y_recargar_cache()
        if not os.path.exists(self.ruta) and wb is None and self._wb_cache is None:
            return []

        should_close = not bool(self._wb_cache) if wb is None else False
        if wb is None:
            wb = self._wb_cache if self._wb_cache else openpyxl.load_workbook(self.ruta, data_only=True)

        resultados = []
        try:
            grados = self.obtener_grados_activos(wb=wb)
            for g in grados:
                estudiantes = self.obtener_estudiantes_completos(g, wb=wb)
                if not estudiantes:
                    continue
                materias_grado = self.obtener_materias_por_grado(g, wb=wb)
                
                # Materias reales limpias
                materias_limpias = [m for m in materias_grado if m and m not in ["Sin materias", "No hay materias", "General", "Todas las Materias"]]
                if not materias_limpias:
                    continue
                
                # Identificar materias de tecnología y materias normales
                tech_mats = []
                norm_mats = []
                for m in materias_limpias:
                    m_upper = m.upper()
                    if "HOGAR" in m_upper or "DESARROLLO" in m_upper or "AGRO" in m_upper or "COMERCIO" in m_upper:
                        tech_mats.append(m)
                    else:
                        norm_mats.append(m)
                
                # Obtener notas para todas las materias en todos los trimestres
                notas_por_materia_trimestre = {}
                for trim in ["Trimestre 1", "Trimestre 2", "Trimestre 3"]:
                    for m in materias_limpias:
                        proms = self.obtener_promedios_reales(g, m, trim, wb=wb)
                        if proms:
                            notas_por_materia_trimestre[(m, trim)] = proms
                            
                # Para cada estudiante, calcular promedio general anual
                for est in estudiantes:
                    nom = est["nombre"]
                    
                    # Para cada trimestre, calcular promedio general
                    promedios_trimestrales = []
                    for trim in ["Trimestre 1", "Trimestre 2", "Trimestre 3"]:
                        # Notas de materias normales
                        norm_notas = []
                        for m in norm_mats:
                            proms_dict = notas_por_materia_trimestre.get((m, trim), {})
                            if nom in proms_dict and proms_dict[nom] is not None:
                                norm_notas.append(proms_dict[nom])
                                
                        # Notas de tecnología
                        tech_notas = []
                        for m in tech_mats:
                            proms_dict = notas_por_materia_trimestre.get((m, trim), {})
                            if nom in proms_dict and proms_dict[nom] is not None:
                                tech_notas.append(proms_dict[nom])
                        
                        # Promedio de Tecnología
                        tec_prom = sum(tech_notas) / len(tech_notas) if tech_notas else None
                        
                        # Combinar
                        trim_notas = list(norm_notas)
                        if tec_prom is not None:
                            trim_notas.append(tec_prom)
                            
                        if trim_notas:
                            trim_avg = sum(trim_notas) / len(trim_notas)
                            promedios_trimestrales.append(trim_avg)
                    
                    # Promedio general acumulado del año (promedio de los promedios trimestrales con notas)
                    if promedios_trimestrales:
                        prom_general = sum(promedios_trimestrales) / len(promedios_trimestrales)
                        # Redondear a 2 decimales
                        prom_general = round(prom_general, 2)
                        if 4.5 <= prom_general <= 5.0:
                            resultados.append({
                                "nombre": nom,
                                "grado": g,
                                "promedio": prom_general
                            })
        except Exception as e:
            print(f"[!] Error calculating cuadro de honor: {e}")
        finally:
            if should_close: wb.close()

        # Ordenar de mayor a menor promedio
        resultados = sorted(resultados, key=lambda x: x["promedio"], reverse=True)
        return resultados

    def get_dashboard_stats(self, wb=None):
        if wb is None:
            self._verificar_y_recargar_cache()
        if not os.path.exists(self.ruta) and wb is None and self._wb_cache is None:
            return {"total": 0, "riesgo": 0, "honor": "N/A", "honor_cant": 0, "asistencia": "—", "tareas_sin_nota": 0, "excusas": 0, "habitos": {"S": 0, "R": 0, "X": 0}}

        should_close = not bool(self._wb_cache) if wb is None else False
        if wb is None:
            wb = self._wb_cache if self._wb_cache else openpyxl.load_workbook(self.ruta, data_only=True)
        
        from utils.date_helpers import obtener_trimestre_actual
        trimestre_actual = obtener_trimestre_actual()
        
        total = 0
        riesgo = 0
        excusas = 0
        tareas_sin_nota = 0
        total_asist_dias = 0
        total_asist_ausencias = 0

        # Obtener Cuadro de Honor (General del año entero)
        cuadro_honor = self.obtener_cuadro_honor_general(wb=wb)
        honor_cant = len(cuadro_honor)
        best_student = cuadro_honor[0]["nombre"] if honor_cant > 0 else "N/A"
        best_prom = cuadro_honor[0]["promedio"] if honor_cant > 0 else 0.0

        try:
            grados = self.obtener_grados_activos(wb=wb)
            for g in grados:
                estudiantes = self.obtener_estudiantes_completos(g, wb=wb)
                total += len(estudiantes)
                
                # Promedios de riesgo para el trimestre actual solamente
                proms = self.obtener_promedios_reales(g, None, trimestre_actual, wb=wb)
                for nom, prom in proms.items():
                    if prom is not None:
                        if prom < 3.0:
                            riesgo += 1

                # Conteo de excusas E y promedio de asistencia para el trimestre actual solamente
                hoja_asist = None
                g_clean = g.replace("°", "")
                for s in wb.sheetnames:
                    if "ASISTENCIA" in s.upper() and (self.modalidad == "primaria" or g_clean in s):
                        hoja_asist = s
                        break
                if hoja_asist:
                    ws_as = wb[hoja_asist]
                    mapa_trimestres = {"Trimestre 1": 2, "Trimestre 2": 45, "Trimestre 3": 88}
                    fila_fechas = mapa_trimestres.get(trimestre_actual, 2)
                    for r in range(fila_fechas + 1, fila_fechas + 1 + len(estudiantes)):
                        for c in range(3, 61):
                            val_fecha = ws_as.cell(row=fila_fechas, column=c).value
                            if val_fecha:
                                val = ws_as.cell(row=r, column=c).value
                                if val is not None and str(val).strip():
                                    total_asist_dias += 1
                                    if val == "E":
                                        excusas += 1
                                    elif val == "-":
                                        total_asist_ausencias += 1

                # Conteo de tareas sin nota (vacías) para el trimestre actual solamente
                materias = self.obtener_materias_por_grado(g, wb=wb)
                for mat in materias:
                    if mat in ["Sin materias", "No hay materias", "General"]:
                        continue
                    hoja_prom = self._encontrar_hoja_prom(wb, g, mat)
                    if not hoja_prom:
                        continue
                    ws_pm = wb[hoja_prom]
                    for trimestre in [trimestre_actual]:
                        for tipo_nota in ["Diaria / Parcial", "Apreciación", "Examen"]:
                            col_inicio, col_fin = self._obtener_rango_columnas(ws_pm, trimestre, tipo_nota)
                            if col_inicio is None or col_fin is None:
                                continue
                            for c in range(col_inicio, col_fin + 1):
                                desc = ws_pm.cell(row=self.fila_desc, column=c).value
                                if desc and str(desc).strip():
                                    for r in range(5, 5 + len(estudiantes)):
                                        nota = ws_pm.cell(row=r, column=c).value
                                        if nota is None or str(nota).strip() == "":
                                            tareas_sin_nota += 1
        except Exception as e:
            print(f"[!] Error loading dashboard stats: {e}")
        finally:
            if should_close: wb.close()

        asistencia_pct = "—"
        if total_asist_dias > 0:
            presentes = total_asist_dias - total_asist_ausencias
            pct = (presentes / total_asist_dias) * 100
            asistencia_pct = f"{pct:.1f}%"

        honor = f"{best_student} ({best_prom:.2f})" if best_prom > 0.0 else "N/A"

        # Conteo de habitos y actitudes
        s_count = 0
        r_count = 0
        x_count = 0
        try:
            import json
            ruta_json = os.path.abspath(os.path.join(os.path.dirname(self.ruta), "Expedientes_Estudiantes", "habitos_evaluaciones.json"))
            if os.path.exists(ruta_json):
                with open(ruta_json, "r", encoding="utf-8") as f:
                    habitos_data = json.load(f)
                grados_activos = self.obtener_grados_activos(wb=wb)
                for key, val_entry in habitos_data.items():
                    grade_part = key.split("::")[0]
                    if any(g.replace("°","") in grade_part.replace("°","") for g in grados_activos):
                        est_evals = val_entry.get("estudiantes", {})
                        for est_id, crit_vals in est_evals.items():
                            for score in crit_vals.values():
                                if score == "S":
                                    s_count += 1
                                elif score == "R":
                                    r_count += 1
                                elif score == "X":
                                    x_count += 1
        except Exception as e:
            print(f"Error calculating habits stats in dashboard: {e}")

        return {
            "total": total,
            "riesgo": riesgo,
            "honor": honor,
            "honor_cant": honor_cant,
            "asistencia": asistencia_pct,
            "tareas_sin_nota": tareas_sin_nota,
            "excusas": excusas,
            "habitos": {
                "S": s_count,
                "R": r_count,
                "X": x_count
            }
        }

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
        self._save_wb(wb)
        wb.close()
        self._cargar_en_memoria()
        self.actualizar_resumen(grado)
        return True, ""

    def obtener_descripciones_notas(self, grado, materia, trimestre, tipo_nota, wb=None):
        if wb is None:
            self._verificar_y_recargar_cache()
        if not os.path.exists(self.ruta) and wb is None and self._wb_cache is None: return []
        should_close = not bool(self._wb_cache) if wb is None else False
        if wb is None:
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

    def buscar_notas_por_descripcion_exacta(self, grado, materia, trimestre, tipo_nota, descripcion, wb=None):
        if wb is None:
            self._verificar_y_recargar_cache()
        if not os.path.exists(self.ruta) and wb is None and self._wb_cache is None: return None
        should_close = not bool(self._wb_cache) if wb is None else False
        if wb is None:
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

        self._save_wb(wb)
        wb.close()
        self._cargar_en_memoria()
        self.actualizar_resumen(grado)
        return True

    def _encontrar_hoja_asistencia(self, wb, grado):
        for sheet in wb.sheetnames:
            if "Asistencia" in sheet and (self.modalidad == "primaria" or grado.replace('°','') in sheet):
                return sheet
        return None

    def obtener_fechas_asistencia(self, grado, trimestre, wb=None):
        if wb is None:
            self._verificar_y_recargar_cache()
        if not os.path.exists(self.ruta) and wb is None and self._wb_cache is None: return []
        should_close = not bool(self._wb_cache) if wb is None else False
        if wb is None:
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

    def buscar_asistencia_existente(self, grado, trimestre, fecha, wb=None):
        if wb is None:
            self._verificar_y_recargar_cache()
        if not os.path.exists(self.ruta) and wb is None and self._wb_cache is None: return None
        should_close = not bool(self._wb_cache) if wb is None else False
        if wb is None:
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
        self._save_wb(wb)
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
        self._save_wb(wb)
        wb.close()
        self._cargar_en_memoria()
        return True

    def actualizar_datos_generales(self, nombre_docente, ano_lectivo):
        if not os.path.exists(self.ruta): return False
        wb = openpyxl.load_workbook(self.ruta)
        nombre_maestro = self._encontrar_hoja_maestro(wb)
        ws_m = wb[nombre_maestro]
        titulo_actual = str(ws_m.cell(row=1, column=1).value or "")
        nuevo_titulo = re.sub(r'20\d{2}', str(ano_lectivo), titulo_actual)
        self._safe_set_value(ws_m, 1, 1, nuevo_titulo)
        self._save_wb(wb)
        wb.close()
        self._cargar_en_memoria()
        return True

    def obtener_consejero_actual(self, grado, wb=None):
        if wb is None:
            self._verificar_y_recargar_cache()
        if not os.path.exists(self.ruta) and wb is None and self._wb_cache is None: return "No asignado"
        should_close = not bool(self._wb_cache) if wb is None else False
        if wb is None:
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
        self._save_wb(wb)
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

        if self.modalidad == "primaria":
            hoja_base_maestro = None
            for sheet in wb.sheetnames:
                if "MAESTRO" in sheet.upper():
                    hoja_base_maestro = sheet
                    break
            if hoja_base_maestro:
                nueva_hoja_maestro = f"MAESTRO ({nuevo_grado})"
                if nueva_hoja_maestro not in wb.sheetnames:
                    ws_maestro_nueva = wb.copy_worksheet(wb[hoja_base_maestro])
                    ws_maestro_nueva.title = nueva_hoja_maestro
                    # Clean the names and keep the layout
                    for r in range(5, 50):
                        try:
                            self._safe_clear_value(ws_maestro_nueva, r, 2)
                        except AttributeError: pass
                    self._safe_set_value(ws_maestro_nueva, 3, 2, f"GRADO {nuevo_grado}")
        else:
            nombre_maestro = self._encontrar_hoja_maestro(wb)
            ws_m = wb[nombre_maestro]
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
        
        self._save_wb(wb)
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

        nombre_maestro = self._encontrar_hoja_maestro(wb)
        ws_m = wb[nombre_maestro]
        for c in range(1, 40):
            val = str(ws_m.cell(row=3, column=c).value or "")
            if grado in val:
                for r in range(3, 51):
                    try:
                        self._safe_clear_value(ws_m, r, c)
                        self._safe_clear_value(ws_m, r, c+1)
                    except AttributeError: pass
                break
        self._save_wb(wb)
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

        self._save_wb(wb)
        wb.close()
        self._cargar_en_memoria()
        return True, "Materia clonada y agregada al Resumen."


    def actualizar_resumen(self, grado, wb=None):
        if wb is None:
            self._verificar_y_recargar_cache()
        if not os.path.exists(self.ruta) and wb is None and self._wb_cache is None: return False
        should_close = not bool(self._wb_cache) if wb is None else False
        if wb is None:
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

            self._save_wb(wb_write)
            wb_write.close()

        if should_close:
            wb.close()
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

        self._save_wb(wb)
        wb.close()
        self._cargar_en_memoria()
        return True

    def obtener_estadisticas_asistencia(self, grado, trimestre, id_estudiante, wb=None):
        if wb is None:
            self._verificar_y_recargar_cache()
        if not os.path.exists(self.ruta) and wb is None and self._wb_cache is None:
            return {"total_dias": 0, "ausencias": 0, "tardanzas": 0, "excusas": 0}
        should_close = not bool(self._wb_cache) if wb is None else False
        if wb is None:
            wb = self._wb_cache if self._wb_cache else openpyxl.load_workbook(self.ruta, data_only=True)
        hoja = self._encontrar_hoja_asistencia(wb, grado)
        if not hoja:
            if should_close: wb.close()
            return {"total_dias": 0, "ausencias": 0, "tardanzas": 0, "excusas": 0}
        ws = wb[hoja]
        mapa_trimestres = {"Trimestre 1": 2, "Trimestre 2": 45, "Trimestre 3": 88}
        fila_fechas = mapa_trimestres.get(trimestre, 2)
        
        total_dias = 0
        ausencias = 0
        tardanzas = 0
        excusas = 0
        
        for c in range(3, 61):
            fecha_val = ws.cell(row=fila_fechas, column=c).value
            if not fecha_val:
                continue
            val = ws.cell(row=fila_fechas + int(id_estudiante), column=c).value
            if val is not None and str(val).strip():
                total_dias += 1
                if val == "-":
                    ausencias += 1
                elif val == "T":
                    tardanzas += 1
                elif val == "E":
                    excusas += 1
                
        if should_close: wb.close()
        return {
            "total_dias": total_dias,
            "ausencias": ausencias,
            "tardanzas": tardanzas,
            "excusas": excusas
        }

    def obtener_tareas_sin_nota(self, grado, trimestre, id_estudiante, wb=None):
        if wb is None:
            self._verificar_y_recargar_cache()
        if not os.path.exists(self.ruta) and wb is None and self._wb_cache is None:
            return {"tareas_vacias_por_materia": {}, "total_vacias": 0}
        should_close = not bool(self._wb_cache) if wb is None else False
        if wb is None:
            wb = self._wb_cache if self._wb_cache else openpyxl.load_workbook(self.ruta, data_only=True)
            
        materias = self.obtener_materias_por_grado(grado, wb=wb)
        tareas_vacias = {}
        total_vacias = 0
        
        row_est = 4 + int(id_estudiante)
        
        for mat in materias:
            if mat == "Sin materias registradas":
                continue
            hoja = self._encontrar_hoja_prom(wb, grado, mat)
            if not hoja:
                continue
            ws = wb[hoja]
            
            cnt_mat = 0
            for tipo_nota in ["Diaria / Parcial", "Apreciación", "Examen"]:
                col_inicio, col_fin = self._obtener_rango_columnas(ws, trimestre, tipo_nota)
                if col_inicio is None or col_fin is None:
                    continue
                for c in range(col_inicio, col_fin + 1):
                    desc = ws.cell(row=self.fila_desc, column=c).value
                    if desc and str(desc).strip():
                        nota = ws.cell(row=row_est, column=c).value
                        if nota is None or str(nota).strip() == "":
                            cnt_mat += 1
            if cnt_mat > 0:
                tareas_vacias[mat] = cnt_mat
                total_vacias += cnt_mat
                
        if should_close: wb.close()
        return {
            "tareas_vacias_por_materia": tareas_vacias,
            "total_vacias": total_vacias
        }