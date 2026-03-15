Attribute VB_Name = "RD_Macros"

Sub IrMenu()
Sheets("MENU PRINCIPAL").Activate
End Sub

Sub IrMaestro()
Sheets("MAESTRO").Activate
End Sub

Sub IrNotas()
Sheets("PANEL_NOTAS").Activate
End Sub

Sub IrAsistencia()
Sheets("PANEL_ASISTENCIA").Activate
End Sub

Sub IrDashboard()
Sheets("DASHBOARD").Activate
End Sub

Sub IrResumen7()
Sheets("RESUMEN_7").Activate
End Sub

Sub IrResumen8()
Sheets("RESUMEN_8").Activate
End Sub

Sub IrResumen9()
Sheets("RESUMEN_9").Activate
End Sub

Sub Grado7Notas()
Dim ws As Worksheet
Set ws = Sheets("PANEL_NOTAS")
Dim i As Integer
For i = 1 To 40
ws.Cells(9 + i, 3).Formula = "=IF(MAESTRO!B" & (i + 4) & "="""","""",MAESTRO!B" & (i + 4) & ")"
Next i
MsgBox "Grado 7 cargado. Ingles + Agropecuaria + Hogar", vbInformation, "OK"
End Sub

Sub Grado8Notas()
Dim ws As Worksheet
Set ws = Sheets("PANEL_NOTAS")
Dim i As Integer
For i = 1 To 40
ws.Cells(9 + i, 3).Formula = "=IF(MAESTRO!D" & (i + 4) & "="""","""",MAESTRO!D" & (i + 4) & ")"
Next i
MsgBox "Grado 8 cargado. Ingles + Agropecuaria + Comercio", vbInformation, "OK"
End Sub

Sub Grado9Notas()
Dim ws As Worksheet
Set ws = Sheets("PANEL_NOTAS")
Dim i As Integer
For i = 1 To 40
ws.Cells(9 + i, 3).Formula = "=IF(MAESTRO!F" & (i + 4) & "="""","""",MAESTRO!F" & (i + 4) & ")"
Next i
MsgBox "Grado 9 cargado. Ingles + Agropecuaria + Comercio + Artistica", vbInformation, "OK"
End Sub

Sub Grado7Asistencia()
Dim ws As Worksheet
Set ws = Sheets("PANEL_ASISTENCIA")
Dim i As Integer
For i = 1 To 40
ws.Cells(5 + i, 2).Formula = "=IF(MAESTRO!B" & (i + 4) & "="""","""",MAESTRO!B" & (i + 4) & ")"
Next i
MsgBox "Asistencia Grado 7 cargada", vbInformation, "OK"
End Sub

Sub Grado8Asistencia()
Dim ws As Worksheet
Set ws = Sheets("PANEL_ASISTENCIA")
Dim i As Integer
For i = 1 To 40
ws.Cells(5 + i, 2).Formula = "=IF(MAESTRO!D" & (i + 4) & "="""","""",MAESTRO!D" & (i + 4) & ")"
Next i
MsgBox "Asistencia Grado 8 cargada", vbInformation, "OK"
End Sub

Sub Grado9Asistencia()
Dim ws As Worksheet
Set ws = Sheets("PANEL_ASISTENCIA")
Dim i As Integer
For i = 1 To 40
ws.Cells(5 + i, 2).Formula = "=IF(MAESTRO!F" & (i + 4) & "="""","""",MAESTRO!F" & (i + 4) & ")"
Next i
MsgBox "Asistencia Grado 9 cargada", vbInformation, "OK"
End Sub

Sub Respaldo()
Dim ruta As String
ruta = ThisWorkbook.Path & "\" & "Respaldo_" & Format(Now, "YYYYMMDD_HHMM") & ".xlsm"
ThisWorkbook.SaveCopyAs ruta
MsgBox "Respaldo guardado:" & Chr(10) & ruta, vbInformation, "OK"
End Sub

Sub LimpiarNotas()
Dim r As Integer
r = MsgBox("Borrar TODAS las notas del Panel?" & Chr(10) & "Los nombres NO se borran.", vbYesNo + vbCritical, "Confirmar")
If r = vbNo Then Exit Sub
Dim ws As Worksheet
Set ws = Sheets("PANEL_NOTAS")
Dim i As Integer
Dim c As Integer
Dim cols(8) As Integer
cols(0) = 4
cols(1) = 5
cols(2) = 6
cols(3) = 8
cols(4) = 9
cols(5) = 10
cols(6) = 12
cols(7) = 13
cols(8) = 14
For i = 1 To 40
For c = 0 To 8
ws.Cells(9 + i, cols(c)).ClearContents
Next c
Next i
MsgBox "Notas borradas. Nombres conservados.", vbInformation, "Listo"
End Sub

Sub BuscarAlumno()
Dim nombre As String
nombre = InputBox("Escribe apellido o nombre:", "Buscar")
If nombre = "" Then Exit Sub
nombre = LCase(nombre)
Dim ws As Worksheet
Set ws = Sheets("MAESTRO")
Dim res As String
res = ""
Dim i As Integer
Dim val As String
For i = 5 To 44
val = LCase(CStr(ws.Range("B" & i).Value))
If val <> "" And InStr(val, nombre) > 0 Then
res = res & "7 - Fila " & (i - 4) & ": " & ws.Range("B" & i).Value & Chr(10)
End If
val = LCase(CStr(ws.Range("D" & i).Value))
If val <> "" And InStr(val, nombre) > 0 Then
res = res & "8 - Fila " & (i - 4) & ": " & ws.Range("D" & i).Value & Chr(10)
End If
val = LCase(CStr(ws.Range("F" & i).Value))
If val <> "" And InStr(val, nombre) > 0 Then
res = res & "9 - Fila " & (i - 4) & ": " & ws.Range("F" & i).Value & Chr(10)
End If
Next i
If res = "" Then
MsgBox "No se encontro: " & nombre, vbInformation, "Sin resultados"
Else
MsgBox "Encontrados:" & Chr(10) & Chr(10) & res, vbInformation, "Resultado"
End If
End Sub

Sub Estadisticas()
Dim ws As Worksheet
Set ws = Sheets("MAESTRO")
Dim t7 As Integer
Dim t8 As Integer
Dim t9 As Integer
t7 = Application.WorksheetFunction.CountA(ws.Range("B5:B44"))
t8 = Application.WorksheetFunction.CountA(ws.Range("D5:D44"))
t9 = Application.WorksheetFunction.CountA(ws.Range("F5:F44"))
MsgBox "TOTAL DE ESTUDIANTES" & Chr(10) & _
String(30, "-") & Chr(10) & _
"Grado 7:  " & t7 & Chr(10) & _
"Grado 8:  " & t8 & Chr(10) & _
"Grado 9:  " & t9 & Chr(10) & _
String(30, "-") & Chr(10) & _
"TOTAL:    " & (t7 + t8 + t9), vbInformation, "Estadisticas"
End Sub

Sub NuevoAnio()
Dim r As Integer
r = MsgBox("NUEVO ANO LECTIVO" & Chr(10) & Chr(10) & _
"Esto borrara nombres, notas y asistencia." & Chr(10) & _
"Ya tienes un respaldo guardado?", vbYesNo + vbCritical, "Confirmar")
If r = vbNo Then
MsgBox "Ejecuta Respaldo() primero.", vbInformation, "Cancelado"
Exit Sub
End If
Sheets("MAESTRO").Range("B5:B44").ClearContents
Sheets("MAESTRO").Range("D5:D44").ClearContents
Sheets("MAESTRO").Range("F5:F44").ClearContents
Dim ws As Worksheet
Set ws = Sheets("PANEL_NOTAS")
Dim i As Integer
Dim c As Integer
Dim cols(8) As Integer
cols(0) = 4
cols(1) = 5
cols(2) = 6
cols(3) = 8
cols(4) = 9
cols(5) = 10
cols(6) = 12
cols(7) = 13
cols(8) = 14
For i = 1 To 40
For c = 0 To 8
ws.Cells(9 + i, cols(c)).ClearContents
Next c
Next i
Sheets("PANEL_ASISTENCIA").Range("C6:AF45").ClearContents
MsgBox "Sistema listo para el nuevo ano." & Chr(10) & "Ve a MAESTRO e ingresa los alumnos.", vbInformation, "Listo"
Sheets("MAESTRO").Activate
End Sub
