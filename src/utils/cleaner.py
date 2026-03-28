import openpyxl

class ExcelCleaner:
    @staticmethod
    def limpiar_todo(ruta, modalidad):
        try:
            wb = openpyxl.load_workbook(ruta)
            mod = modalidad.lower()
            
            # 1. Limpiar MAESTRO (Nombres y Cédulas)
            ws_m = wb["MAESTRO"]
            for r in range(5, 51):
                if mod == "primaria":
                    ws_m.cell(row=r, column=2).value = None # Nombres
                    ws_m.cell(row=r, column=5).value = None # Cédulas
                else: # Premedia
                    for col in [2, 4, 6]: # Columnas B, D, F para 7°, 8°, 9°
                        ws_m.cell(row=r, column=col).value = None

            # 2. Limpiar Notas y Descripciones (PROM)
            fila_desc = 39 if mod == "primaria" else 42
            for sheet in wb.sheetnames:
                if "PROM" in sheet:
                    ws = wb[sheet]
                    for c in range(3, 80):
                        ws.cell(row=4, column=c).value = None # Fechas
                        ws.cell(row=fila_desc, column=c).value = None # Descripciones
                    for r in range(5, 45):
                        for c in range(3, 80):
                            cell = ws.cell(row=r, column=c)
                            if cell.data_type != 'f': # Evitar borrar fórmulas
                                cell.value = None

            # 3. Limpiar Asistencia (Vertical para ambos)
            hoja_asist = "Asistencia " if mod == "primaria" else "ASISTENCIA"
            if hoja_asist in wb.sheetnames:
                ws_a = wb[hoja_asist]
                for inicio, fin in [(3, 41), (46, 86), (89, 130)]:
                    for r in range(inicio, fin + 1):
                        for c in range(3, 61): 
                            ws_a.cell(row=r, column=c).value = None

            wb.save(ruta)
            return True
        except Exception as e:
            print(f"Error crítico al limpiar el archivo {ruta}: {e}")
            return False