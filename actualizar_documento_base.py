import os
import time
import xlwings as xw


BASE_DIR = os.path.dirname(os.path.abspath(__file__))
ORIGEN = os.path.join(BASE_DIR, "RegistroDoc_Multigrado_v1_FINAL.xlsx")
SALIDA = os.path.join(BASE_DIR, "RegistroDoc_Multigrado_v1_FINAL_actualizado.xlsm")


XL_SHEET_VISIBLE = -1
XL_SHEET_VERY_HIDDEN = 2


def rgb_xl(r, g, b):
    return r + g * 256 + b * 65536


def bg(ws, rng, r, g, b):
    ws.range(rng).api.Interior.Color = rgb_xl(r, g, b)


def ft(ws, rng, r, g, b, bold=False, sz=10):
    ws.range(rng).api.Font.Color = rgb_xl(r, g, b)
    ws.range(rng).api.Font.Bold = bold
    ws.range(rng).api.Font.Size = sz


def safe_clear(ws, rng):
    """Limpia rangos aunque existan celdas combinadas."""
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


def crear_portada_educativa(wb):
    ws = get_or_create_sheet(wb, "PORTADA", before=wb.sheets[0])
    safe_clear(ws, "A1:J40")
    ws.api.Cells.UnMerge()

    bg(ws, "A1:J40", 242, 248, 255)
    bg(ws, "A1:J4", 31, 56, 100)
    bg(ws, "B6:I14", 255, 255, 255)

    ws.range("A1:J1").api.Merge()
    ws.range("A2:J2").api.Merge()
    ws.range("A3:J3").api.Merge()
    ws.range("A1").value = "REGISTRODOC MULTIGRADO"
    ws.range("A2").value = "Sistema de Gestion Academica"
    ws.range("A3").value = "Documento base actualizado"
    ft(ws, "A1", 255, 255, 255, True, 28)
    ft(ws, "A2", 226, 240, 255, True, 16)
    ft(ws, "A3", 200, 220, 245, False, 12)
    for c in ["A1", "A2", "A3"]:
        ws.range(c).api.HorizontalAlignment = -4108

    ws.range("B6:I6").api.Merge()
    ws.range("B6").value = "IMPORTANTE: HABILITA MACROS"
    ft(ws, "B6", 31, 56, 100, True, 18)
    ws.range("B6").api.HorizontalAlignment = -4108

    ws.range("B8:I8").api.Merge()
    ws.range("B8").value = "1) Activa contenido/macros al abrir"
    ws.range("B9:I9").api.Merge()
    ws.range("B9").value = "2) Si no aparece login, cierra y abre nuevamente"
    ws.range("B10:I10").api.Merge()
    ws.range("B10").value = "3) Accede y usa menu Reportes para generar salidas"
    for c in ["B8", "B9", "B10"]:
        ft(ws, c, 0, 102, 51, True, 12)

    ws.api.Columns("A:J").ColumnWidth = 12
    ws.api.Rows(1).RowHeight = 34
    ws.api.Rows(2).RowHeight = 26
    ws.api.Rows(6).RowHeight = 28

    ws.activate()


def crear_dashboard_graficos(wb):
    ws = get_or_create_sheet(wb, "DASHBOARD")
    safe_clear(ws, "A1:N45")

    ws.range("A1:F1").api.Merge()
    ws.range("A1").value = "Dashboard Academico"
    ft(ws, "A1", 31, 56, 100, True, 22)
    bg(ws, "A1:N3", 226, 239, 218)

    ws.range("A4").value = ["Indicador", "Valor"]
    ft(ws, "A4:B4", 255, 255, 255, True, 11)
    bg(ws, "A4:B4", 31, 56, 100)

    indicadores = [
        ["Estudiantes Grado 7", "=COUNTA(MAESTRO!B5:B44)"],
        ["Estudiantes Grado 8", "=COUNTA(MAESTRO!D5:D44)"],
        ["Estudiantes Grado 9", "=COUNTA(MAESTRO!F5:F44)"],
    ]
    ws.range("A5").value = indicadores

    ws.range("D4").value = ["Grado", "Total"]
    ws.range("D5").value = [["7", "=B5"], ["8", "=B6"], ["9", "=B7"]]

    # En algunas versiones de xlwings, ws.charts.add puede devolver una tupla;
    # usamos COM directo para evitar incompatibilidades.
    chart_obj = ws.api.ChartObjects().Add(
        ws.range("D9").left,
        ws.range("D9").top,
        460,
        260,
    )
    chart = chart_obj.Chart
    chart.SetSourceData(ws.range("D4:E7").api)
    chart.ChartType = 51  # xlColumnClustered
    chart.HasTitle = True
    chart.ChartTitle.Text = "Cantidad de estudiantes por grado"

    ws.api.Columns("A:B").ColumnWidth = 28
    ws.api.Columns("D:E").ColumnWidth = 14


def crear_hojas_reportes(wb):
    ws_notas = get_or_create_sheet(wb, "REP_NOTAS")
    safe_clear(ws_notas, "A1:H200")
    ws_notas.range("A1:H1").value = [[
        "Fecha", "Grado", "Materia", "Estudiante", "Trimestre", "Nota", "Docente", "Observacion"
    ]]
    bg(ws_notas, "A1:H1", 31, 56, 100)
    ft(ws_notas, "A1:H1", 255, 255, 255, True, 10)
    ws_notas.api.Columns("A:H").ColumnWidth = 18

    ws_asis = get_or_create_sheet(wb, "REP_ASISTENCIA")
    safe_clear(ws_asis, "A1:G200")
    ws_asis.range("A1:G1").value = [[
        "Fecha", "Grado", "Estudiante", "Presente", "Ausente", "Tardanza", "Observacion"
    ]]
    bg(ws_asis, "A1:G1", 112, 48, 160)
    ft(ws_asis, "A1:G1", 255, 255, 255, True, 10)
    ws_asis.api.Columns("A:G").ColumnWidth = 18


def inyectar_modulo_reportes(wb):
    """Crea/actualiza modulo VBA para generar reportes desde formularios/menu."""
    vbproj = wb.api.VBProject
    modulo_nombre = "RD_REPORTES"

    modulo = None
    for comp in vbproj.VBComponents:
        if comp.Name == modulo_nombre:
            modulo = comp
            break

    if modulo is None:
        modulo = vbproj.VBComponents.Add(1)  # vbext_ct_StdModule
        modulo.Name = modulo_nombre

    cm = modulo.CodeModule
    if cm.CountOfLines > 0:
        cm.DeleteLines(1, cm.CountOfLines)

    codigo = r'''
Option Explicit

Public Sub GenerarReporteNotas()
    On Error GoTo ErrHandler
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("REP_NOTAS")

    ws.Visible = xlSheetVisible
    ws.Activate
    MsgBox "Reporte de notas listo. Complete o exporte esta hoja.", vbInformation, "RegistroDoc"
    Exit Sub

ErrHandler:
    MsgBox "No se pudo abrir REP_NOTAS: " & Err.Description, vbExclamation, "RegistroDoc"
End Sub

Public Sub GenerarReporteAsistencia()
    On Error GoTo ErrHandler
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("REP_ASISTENCIA")

    ws.Visible = xlSheetVisible
    ws.Activate
    MsgBox "Reporte de asistencia listo. Complete o exporte esta hoja.", vbInformation, "RegistroDoc"
    Exit Sub

ErrHandler:
    MsgBox "No se pudo abrir REP_ASISTENCIA: " & Err.Description, vbExclamation, "RegistroDoc"
End Sub

Public Sub ActualizarDashboard()
    On Error Resume Next
    ThisWorkbook.Sheets("DASHBOARD").Visible = xlSheetVisible
    ThisWorkbook.Sheets("DASHBOARD").Activate
    On Error GoTo 0
End Sub
'''
    cm.AddFromString(codigo)


def asegurar_visibilidad_controlada(wb):
    """Mantiene solo PORTADA visible para evitar mostrar el libro completo al abrir."""
    for ws in wb.sheets:
        ws.api.Visible = XL_SHEET_VERY_HIDDEN
    wb.sheets["PORTADA"].api.Visible = XL_SHEET_VISIBLE
    wb.sheets["PORTADA"].activate()


def main():
    if not os.path.exists(ORIGEN):
        raise FileNotFoundError(f"No existe el archivo origen: {ORIGEN}")

    app = xw.App(visible=False)
    app.display_alerts = False
    app.screen_updating = False

    try:
        wb = app.books.open(ORIGEN)

        print("[1/5] Ajustando portada y celdas combinadas...")
        crear_portada_educativa(wb)

        print("[2/5] Creando hoja DASHBOARD con graficos...")
        crear_dashboard_graficos(wb)

        print("[3/5] Creando hojas REP_NOTAS y REP_ASISTENCIA...")
        crear_hojas_reportes(wb)

        print("[4/5] Inyectando modulo VBA de reportes...")
        inyectar_modulo_reportes(wb)

        print("[5/5] Aplicando visibilidad controlada y guardando...")
        asegurar_visibilidad_controlada(wb)
        time.sleep(1)
        wb.api.SaveAs(SALIDA, FileFormat=52)  # xlsm actualizado
        wb.close(save_changes=False)
        print(f"OK. Archivo actualizado: {SALIDA}")
    finally:
        try:
            app.quit()
        except Exception:
            pass


if __name__ == "__main__":
    main()
