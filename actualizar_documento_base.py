import os
import time
import xlwings as xw

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
ORIGEN = os.path.join(BASE_DIR, "RegistroDoc_Multigrado_v1_FINAL.xlsx")
SALIDA = os.path.join(BASE_DIR, "RegistroDoc_Multigrado_v1_FINAL_actualizado.xlsx")

XL_SHEET_VISIBLE = -1
XL_SHEET_HIDDEN = 0


def rgb_xl(r, g, b):
    return r + g * 256 + b * 65536


def bg(ws, rng, r, g, b):
    ws.range(rng).api.Interior.Color = rgb_xl(r, g, b)


def ft(ws, rng, r, g, b, bold=False, sz=10):
    ws.range(rng).api.Font.Color = rgb_xl(r, g, b)
    ws.range(rng).api.Font.Bold = bold
    ws.range(rng).api.Font.Size = sz


def safe_clear(ws, rng):
    try:
        ws.range(rng).clear()
    except Exception:
        ws.range(rng).api.UnMerge()
        ws.range(rng).clear()


def get_or_create_sheet(wb, name, before=None):
    try:
        return wb.sheets[name]
    except Exception:
        if before is None:
            return wb.sheets.add(name)
        return wb.sheets.add(name, before=before)


def crear_portada_alternativa(wb):
    """Portada vistosa con informacion academica relevante."""
    ws = get_or_create_sheet(wb, "PORTADA_VISTOSA", before=wb.sheets[0])
    safe_clear(ws, "A1:K40")
    ws.api.Cells.UnMerge()

    bg(ws, "A1:K40", 242, 248, 255)
    bg(ws, "A1:K4", 31, 56, 100)
    bg(ws, "B7:J15", 255, 255, 255)
    bg(ws, "B18:J30", 255, 255, 255)

    ws.range("A1:K1").api.Merge()
    ws.range("A2:K2").api.Merge()
    ws.range("A3:K3").api.Merge()
    ws.range("A1").value = "REGISTRODOC MULTIGRADO"
    ws.range("A2").value = "PORTADA ACADEMICA"
    ws.range("A3").value = "Control de notas y asistencia"
    ft(ws, "A1", 255, 255, 255, True, 30)
    ft(ws, "A2", 220, 235, 255, True, 16)
    ft(ws, "A3", 200, 220, 245, False, 12)
    for c in ["A1", "A2", "A3"]:
        ws.range(c).api.HorizontalAlignment = -4108

    ws.range("B7:J7").api.Merge()
    ws.range("B7").value = "INFORMACION INSTITUCIONAL"
    ft(ws, "B7", 31, 56, 100, True, 14)

    info = [
        ["Escuela", "Escuela Cerro Cacicon"],
        ["Docente", "Elmer Tugri"],
        ["Ano lectivo", "2026"],
        ["Grados", "7°, 8° y 9°"],
        ["Materias", "Ingles, Agropecuaria, Comercio, Hogar y Desarrollo, Artistica"],
    ]
    ws.range("B9").value = info
    ft(ws, "B9:B13", 31, 56, 100, True, 11)
    ws.api.Columns("B").ColumnWidth = 20
    ws.api.Columns("C").ColumnWidth = 55

    ws.range("B18:J18").api.Merge()
    ws.range("B18").value = "HOJAS CLAVE DEL LIBRO"
    ft(ws, "B18", 31, 56, 100, True, 14)

    hojas = [
        ["MAESTRO", "Lista base de estudiantes por grado"],
        ["DASHBOARD", "Resumen visual con graficas"],
        ["Asistencia 7/8/9", "Registro directo por hoja de grado"],
        ["PROM y PLANILLAS", "Calificaciones por materia y grado"],
    ]
    ws.range("B20").value = hojas
    ft(ws, "B20:B23", 0, 102, 51, True, 11)

    for rng in ["B7:J15", "B18:J30"]:
        ws.range(rng).api.Borders(7).LineStyle = 1
        ws.range(rng).api.Borders(8).LineStyle = 1
        ws.range(rng).api.Borders(9).LineStyle = 1
        ws.range(rng).api.Borders(10).LineStyle = 1

    ws.api.Columns("A:K").ColumnWidth = 12
    ws.api.Rows(1).RowHeight = 38


def crear_dashboard_graficos(wb):
    ws = get_or_create_sheet(wb, "DASHBOARD")
    safe_clear(ws, "A1:N50")
    ws.api.Cells.UnMerge()

    bg(ws, "A1:N50", 248, 251, 255)
    bg(ws, "A1:N3", 31, 56, 100)
    ws.range("A1:N1").api.Merge()
    ws.range("A1").value = "DASHBOARD ACADEMICO"
    ft(ws, "A1", 255, 255, 255, True, 24)

    ws.range("A5").value = [["Indicador", "Valor"]]
    bg(ws, "A5:B5", 31, 56, 100)
    ft(ws, "A5:B5", 255, 255, 255, True, 11)

    ws.range("A6").value = [
        ["Total estudiantes 7°", "=IFERROR(COUNTA(MAESTRO!B5:B44),0)"],
        ["Total estudiantes 8°", "=IFERROR(COUNTA(MAESTRO!D5:D44),0)"],
        ["Total estudiantes 9°", "=IFERROR(COUNTA(MAESTRO!F5:F44),0)"],
        ["Total general", "=SUM(B6:B8)"],
    ]

    ws.range("D5").value = [["Grado", "Riesgo Academico"]]
    bg(ws, "D5:E5", 192, 0, 0)
    ft(ws, "D5:E5", 255, 255, 255, True, 11)
    ws.range("D6").value = [["7°", "=IFERROR(COUNTIF('PROM (Ingles 7°)'!N5:N44,\"<3\"),0)"],
                             ["8°", "=IFERROR(COUNTIF('PROM (Ingles 8°)'!N5:N44,\"<3\"),0)"],
                             ["9°", "=IFERROR(COUNTIF('PROM (Ingles 9°)'!N5:N44,\"<3\"),0)"]]


    ws.range("G5").value = [["Grado", "Riesgo Inasistencia"]]
    bg(ws, "G5:H5", 112, 48, 160)
    ft(ws, "G5:H5", 255, 255, 255, True, 11)
    ws.range("G6").value = [
        ["7°", "=IFERROR(SUMPRODUCT(--('Asistencia (7°)'!C3:AZ42=\"A\")),0)"],
        ["8°", "=IFERROR(SUMPRODUCT(--('Asistencia (8°)'!C3:AZ42=\"A\")),0)"],
        ["9°", "=IFERROR(SUMPRODUCT(--('Asistencia (9°)'!C3:AZ42=\"A\")),0)"],
    ]


    ws.range("A12").value = [["Prioridad", "Descripcion"]]
    bg(ws, "A12:B12", 0, 102, 51)
    ft(ws, "A12:B12", 255, 255, 255, True, 11)
    ws.range("A13").value = [
        ["Atencion academica", "=IF(MAX(E6:E8)>0,\"Reforzar grado \"&INDEX(D6:D8,MATCH(MAX(E6:E8),E6:E8,0)),\"Sin alertas\")"],
        ["Atencion asistencia", "=IF(MAX(H6:H8)>0,\"Revisar ausencias en grado \"&INDEX(G6:G8,MATCH(MAX(H6:H8),H6:H8,0)),\"Sin alertas\")"],
    ]


    ws.range("A16:H18").api.Merge()
    ws.range("A16").value = "Las graficas de riesgo y seguimiento estan en la hoja GRAFICOS."
    ft(ws, "A16", 60, 60, 60, True, 12)

    ws.api.Columns("A").ColumnWidth = 28
    ws.api.Columns("B").ColumnWidth = 45
    ws.api.Columns("D").ColumnWidth = 18
    ws.api.Columns("E").ColumnWidth = 18
    ws.api.Columns("G").ColumnWidth = 20
    ws.api.Columns("H").ColumnWidth = 20


def crear_panel_asistencia(wb):
    """Panel de asistencia con filtro por celda (grado) y fecha editable."""
    ws = get_or_create_sheet(wb, "PANEL_ASISTENCIA")
    safe_clear(ws, "A1:H75")
    ws.api.Cells.UnMerge()

    bg(ws, "A1:H75", 246, 251, 246)
    bg(ws, "A1:H4", 0, 102, 51)

    ws.range("A1:H1").api.Merge()
    ws.range("A1").value = "PANEL DE ASISTENCIA"
    ft(ws, "A1", 255, 255, 255, True, 22)

    ws.range("A3").value = "Grado (filtro):"
    ft(ws, "A3", 255, 255, 255, True, 11)
    ws.range("B3").value = 7
    bg(ws, "B3", 255, 242, 204)
    ft(ws, "B3", 120, 60, 0, True, 12)
    dv = ws.range("B3").api.Validation
    try:
        dv.Delete()
    except Exception:
        pass
    dv.Add(Type=3, AlertStyle=1, Formula1="7,8,9")

    ws.range("D3").value = "Fecha:"
    ft(ws, "D3", 255, 255, 255, True, 11)
    ws.range("E3").formula = "=TODAY()"
    ws.range("E3").number_format = "dd/mm/yyyy"
    bg(ws, "E3", 255, 242, 204)
    ft(ws, "E3", 120, 60, 0, True, 11)

    ws.range("A5:D5").value = [["No.", "APELLIDO, NOMBRE", "ASISTENCIA", "OBSERVACION"]]
    bg(ws, "A5:D5", 31, 56, 100)
    ft(ws, "A5:D5", 255, 255, 255, True, 10)

    for i in range(40):
        fila = 6 + i
        ws.range(f"A{fila}").value = i + 1
        ws.range(f"B{fila}").formula = f"=IF($B$3=7,MAESTRO!B{5+i},IF($B$3=8,MAESTRO!D{5+i},MAESTRO!F{5+i}))"
        ws.range(f"C{fila}").formula = f"=IFERROR(INDIRECT(\"'Asistencia (\"&$B$3&\"°)'!C\"&{3+i}),\"\")"
        ws.range(f"D{fila}").formula = f"=IFERROR(INDIRECT(\"'Asistencia (\"&$B$3&\"°)'!D\"&{3+i}),\"\")"

    ws.range("F3:H3").api.Merge()
    ws.range("F3").value = "Cambia B3 para ver solo alumnos del grado seleccionado"
    bg(ws, "F3:H3", 198, 239, 206)
    ft(ws, "F3", 0, 97, 0, True, 10)

    ws.range("F5:H5").value = [["Indicador", "Valor", "Semaforo"]]
    bg(ws, "F5:H5", 112, 48, 160)
    ft(ws, "F5:H5", 255, 255, 255, True, 10)
    ws.range("F6").value = "Total alumnos visibles"
    ws.range("G6").formula = "=COUNTA(B6:B45)"
    ws.range("F7").value = "Ausencias registradas"
    ws.range("G7").formula = "=COUNTIF(C6:C45,\"A\")"
    ws.range("F8").value = "Alerta"
    ws.range("G8").formula = "=IF(G7>=5,\"ALTA\",\"NORMAL\")"
    ws.range("H8").formula = "=IF(G7>=5,\"ALERTA\",\"OK\")"

    ws.api.Columns("A").ColumnWidth = 7
    ws.api.Columns("B").ColumnWidth = 38
    ws.api.Columns("C").ColumnWidth = 14
    ws.api.Columns("D").ColumnWidth = 30
    ws.api.Columns("F").ColumnWidth = 24
    ws.api.Columns("G").ColumnWidth = 14
    ws.api.Columns("H").ColumnWidth = 12


def crear_hoja_graficos(wb):
    """Hoja visual para decisiones: inasistencia, bajo rendimiento y mejor estudiante."""
    ws = get_or_create_sheet(wb, "GRAFICOS")
    safe_clear(ws, "A1:Q200")
    ws.api.Cells.UnMerge()

    bg(ws, "A1:Q200", 250, 250, 255)
    bg(ws, "A1:Q3", 31, 56, 100)
    ws.range("A1:Q1").api.Merge()
    ws.range("A1").value = "GRAFICOS Y ALERTAS DE RENDIMIENTO"
    ft(ws, "A1", 255, 255, 255, True, 22)

    # Base de analisis por estudiante (3 grados x 40 filas)
    ws.range("A5:F5").value = [["Grado", "No", "Estudiante", "Ausencias", "NotaFinal", "TareasVacias"]]
    bg(ws, "A5:F5", 31, 56, 100)
    ft(ws, "A5:F5", 255, 255, 255, True, 10)

    fila = 6
    for grado in [7, 8, 9]:
        for i in range(40):
            n = i + 1
            src_as = 3 + i
            src_ma = 5 + i
            ws.range(f"A{fila}").value = grado
            ws.range(f"B{fila}").value = n
            if grado == 7:
                ws.range(f"C{fila}").formula = f"=IFERROR(MAESTRO!B{src_ma},"")"
                ws.range(f"D{fila}").formula = f"=IFERROR(COUNTIF('Asistencia (7°)'!C{src_as}:AZ{src_as},\"A\"),0)"
                ws.range(f"E{fila}").formula = f"=IFERROR('PROM (Ingles 7°)'!N{src_ma},0)"
                ws.range(f"F{fila}").formula = f"=IFERROR(COUNTBLANK('PROM (Ingles 7°)'!C{src_ma}:M{src_ma}),0)"
            elif grado == 8:
                ws.range(f"C{fila}").formula = f"=IFERROR(MAESTRO!D{src_ma},"")"
                ws.range(f"D{fila}").formula = f"=IFERROR(COUNTIF('Asistencia (8°)'!C{src_as}:AZ{src_as},\"A\"),0)"
                ws.range(f"E{fila}").formula = f"=IFERROR('PROM (Ingles 8°)'!N{src_ma},0)"
                ws.range(f"F{fila}").formula = f"=IFERROR(COUNTBLANK('PROM (Ingles 8°)'!C{src_ma}:M{src_ma}),0)"
            else:
                ws.range(f"C{fila}").formula = f"=IFERROR(MAESTRO!F{src_ma},"")"
                ws.range(f"D{fila}").formula = f"=IFERROR(COUNTIF('Asistencia (9°)'!C{src_as}:AZ{src_as},\"A\"),0)"
                ws.range(f"E{fila}").formula = f"=IFERROR('PROM (Ingles 9°)'!N{src_ma},0)"
                ws.range(f"F{fila}").formula = f"=IFERROR(COUNTBLANK('PROM (Ingles 9°)'!C{src_ma}:M{src_ma}),0)"
            fila += 1

    # Tarjetas de decision
    ws.range("H5:J5").value = [["Indicador", "Estudiante", "Valor"]]
    bg(ws, "H5:J5", 192, 0, 0)
    ft(ws, "H5:J5", 255, 255, 255, True, 10)

    ws.range("H6").value = "Mas inasistencia"
    ws.range("I6").formula = "=IFERROR(INDEX(C6:C125,MATCH(MAX(D6:D125),D6:D125,0)),"")"
    ws.range("J6").formula = "=MAX(D6:D125)"

    ws.range("H7").value = "Nota mas baja"
    ws.range("I7").formula = "=IFERROR(INDEX(C6:C125,MATCH(MINIFS(E6:E125,E6:E125,\">0\"),E6:E125,0)),\"\")"
    ws.range("J7").formula = "=IFERROR(MINIFS(E6:E125,E6:E125,\">0\"),0)"

    ws.range("H8").value = "No entrega tareas"
    ws.range("I8").formula = "=IFERROR(INDEX(C6:C125,MATCH(MAX(F6:F125),F6:F125,0)),"")"
    ws.range("J8").formula = "=MAX(F6:F125)"

    ws.range("H9").value = "Mejor estudiante"
    ws.range("I9").formula = "=IFERROR(INDEX(C6:C125,MATCH(MAX(E6:E125),E6:E125,0)),"")"
    ws.range("J9").formula = "=MAX(E6:E125)"

    # Resumen por grado para graficas
    ws.range("L5:N5").value = [["Grado", "Ausencias", "RiesgoAcademico"]]
    bg(ws, "L5:N5", 0, 102, 51)
    ft(ws, "L5:N5", 255, 255, 255, True, 10)
    ws.range("L6").value = [[7], [8], [9]]
    ws.range("M6").formula = "=SUMIFS(D6:D125,A6:A125,L6)"
    ws.range("M7").formula = "=SUMIFS(D6:D125,A6:A125,L7)"
    ws.range("M8").formula = "=SUMIFS(D6:D125,A6:A125,L8)"
    ws.range("N6").formula = "=COUNTIFS(A6:A125,L6,E6:E125,\"<3\",E6:E125,\">0\")"
    ws.range("N7").formula = "=COUNTIFS(A6:A125,L7,E6:E125,\"<3\",E6:E125,\">0\")"
    ws.range("N8").formula = "=COUNTIFS(A6:A125,L8,E6:E125,\"<3\",E6:E125,\">0\")"

    # Graficas
    c1 = ws.api.ChartObjects().Add(ws.range("H12").left, ws.range("H12").top, 380, 220)
    ch1 = c1.Chart
    ch1.SetSourceData(ws.range("L5:M8").api)
    ch1.ChartType = 51
    ch1.HasTitle = True
    ch1.ChartTitle.Text = "Ausencias por grado"

    c2 = ws.api.ChartObjects().Add(ws.range("L12").left, ws.range("L12").top, 380, 220)
    ch2 = c2.Chart
    ch2.SetSourceData(ws.range("L5:N8").api)
    ch2.ChartType = 57
    ch2.HasTitle = True
    ch2.ChartTitle.Text = "Estudiantes en riesgo academico"

    ws.range("H35:Q38").api.Merge()
    ws.range("H35").value = "Usa estas alertas para priorizar: 1) mas inasistencia, 2) nota mas baja, 3) no entrega tareas."
    ft(ws, "H35", 60, 60, 60, True, 11)

    ws.api.Columns("A").ColumnWidth = 8
    ws.api.Columns("B").ColumnWidth = 6
    ws.api.Columns("C").ColumnWidth = 34
    ws.api.Columns("D:F").ColumnWidth = 14
    ws.api.Columns("H").ColumnWidth = 22
    ws.api.Columns("I").ColumnWidth = 32
    ws.api.Columns("J").ColumnWidth = 12
    ws.api.Columns("L:N").ColumnWidth = 15


def limpiar_hojas_de_sistema(wb):
    """Quita hojas que simulan aplicacion/sistema y no son necesarias para excel base."""
    hojas_quitar = [
        "MENU", "MENÚ", "CONFIGURACION", "CONFIGURACIÓN", "BD_Usuarios",
        "REP_NOTAS", "REP_ASISTENCIA", "BD_Auditoria", "BD_Config", "BD_Materias", "PANEL_ASISTENCIA"
    ]
    for nombre in hojas_quitar:
        try:
            wb.sheets[nombre].delete()
        except Exception:
            pass


def ajustar_visibilidad(wb):
    """Libro base normal: deja visibles hojas clave y oculta Asistencia 7/8/9."""
    for ws in wb.sheets:
        try:
            ws.api.Visible = XL_SHEET_VISIBLE
        except Exception:
            pass

    # Priorizar portada original si existe; si no, la alternativa.
    for nom in ["PORTADA", "PORTADA_VISTOSA", "Asistencia (7°)"]:
        try:
            wb.sheets[nom].activate()
            break
        except Exception:
            pass


def main():
    if not os.path.exists(ORIGEN):
        raise FileNotFoundError(f"No existe el archivo origen: {ORIGEN}")

    app = xw.App(visible=False)
    app.display_alerts = False
    app.screen_updating = False

    try:
        wb = app.books.open(ORIGEN)

        print("[1/5] Manteniendo PORTADA original y creando PORTADA_VISTOSA...")
        crear_portada_alternativa(wb)

        print("[2/5] Creando DASHBOARD con alertas de riesgo...")
        crear_dashboard_graficos(wb)

        print("[3/5] Creando hoja GRAFICOS de apoyo a decisiones...")
        crear_hoja_graficos(wb)

        print("[4/5] Eliminando hojas de sistema no requeridas...")
        limpiar_hojas_de_sistema(wb)

        print("[5/5] Ajustando visibilidad y guardando Excel normal...")
        ajustar_visibilidad(wb)
        time.sleep(1)
        wb.api.SaveAs(SALIDA, FileFormat=51)
        wb.close()
        print(f"OK. Archivo actualizado: {SALIDA}")
    finally:
        try:
            app.quit()
        except Exception:
            pass


if __name__ == "__main__":
    main()
