Attribute VB_Name = "RD_CORE"
' ====================================
' REGISTRODOC PRO - MÓDULO PRINCIPAL
' ====================================
' Sistema de Registro Académico
' Panamá - 2026

Option Explicit

' ----- CONSTANTES -----
Public Const APP_VERSION = "2.0"
Public Const SHEET_MAESTRO = "MAESTRO"
Public Const SHEET_AUDITORIA = "_Auditoria"

' ----- VARIABLES GLOBALES -----
Public usuario_actual As String
Public fecha_apertura As Date

' ====================================
' INICIALIZACIÓN
' ====================================

Sub Auto_Open()
    ' Se ejecuta al abrir el libro
    On Error Resume Next
    
    usuario_actual = Environ("USERNAME")
    fecha_apertura = Now()
    
    ' Crear hoja de auditoría si no existe
    Call CrearAuditoria
    
    ' Registrar apertura
    Call RegistrarAuditoria("APERTURA", "Libro abierto", "")
    
    MsgBox "Bienvenido a RegistroDoc PRO v" & APP_VERSION, vbInformation, "Inicio"
End Sub

Sub Auto_Close()
    ' Se ejecuta al cerrar el libro
    Call RegistrarAuditoria("CIERRE", "Libro cerrado", "")
End Sub

' ====================================
' AUDITORÍA
' ====================================

Sub CrearAuditoria()
    On Error Resume Next
    
    Dim ws As Worksheet
    Dim exist As Boolean
    
    ' Verificar si existe
    For Each ws In ThisWorkbook.Sheets
        If ws.Name = SHEET_AUDITORIA Then
            exist = True
            Exit For
        End If
    Next ws
    
    ' Si no existe, crearla
    If Not exist Then
        Set ws = ThisWorkbook.Sheets.Add
        ws.Name = SHEET_AUDITORIA
        ws.Visible = xlSheetVeryHidden
    End If
End Sub

Sub RegistrarAuditoria(accion As String, detalles As String, tabla As String)
    On Error Resume Next
    
    Dim ws As Worksheet
    Dim row As Long
    
    Set ws = ThisWorkbook.Sheets(SHEET_AUDITORIA)
    
    if ws Is Nothing Then
        Call CrearAuditoria
        Set ws = ThisWorkbook.Sheets(SHEET_AUDITORIA)
    End If
    
    row = ws.Cells(ws.Rows.Count, 1).End(xlUp).row + 1
    
    With ws
        .Cells(row, 1).value = row - 1 ' ID
        .Cells(row, 2).value = usuario_actual ' Usuario
        .Cells(row, 3).value = accion ' Acción
        .Cells(row, 4).value = tabla ' Tabla
        .Cells(row, 5).value = detalles ' Detalles
        .Cells(row, 6).value = Format(Now(), "dd/mm/yyyy hh:mm:ss") ' Timestamp
    End With
End Sub

' ====================================
' VALIDACIÓN - ESTUDIANTES
' ====================================

Function ValidarCedula(cedula As String) As Boolean
    ' Valida formato de cédula panameña (8-10 dígitos)
    ValidarCedula = (cedula <> "" And Len(cedula) >= 8 And Len(cedula) <= 10)
End Function

Function ValidarEmail(email As String) As Boolean
    ' Validación básica de email
    ValidarEmail = (InStr(email, "@") > 0 And InStr(email, ".") > 0)
End Function

Sub ValidarMaestro()
    ' Valida datos en hoja MAESTRO
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim row As Long
    Dim last_row As Long
    
    Set ws = ThisWorkbook.Sheets(SHEET_MAESTRO)
    last_row = ws.Cells(ws.Rows.Count, 1).End(xlUp).row
    
    For row = 2 To last_row
        ' Validar cédula si está presente
        If ws.Cells(row, 2).value <> "" Then
            If Not ValidarCedula(ws.Cells(row, 2).value) Then
                ws.Cells(row, 2).Interior.Color = RGB(255, 0, 0) ' Rojo
            Else
                ws.Cells(row, 2).Interior.Color = xlNone
            End If
        End If
        
        ' Validar email si está presente
        If ws.Cells(row, 5).value <> "" Then
            If Not ValidarEmail(ws.Cells(row, 5).value) Then
                ws.Cells(row, 5).Interior.Color = RGB(255, 255, 0) ' Amarillo
            Else
                ws.Cells(row, 5).Interior.Color = xlNone
            End If
        End If
    Next row
    
    MsgBox "Validación completada", vbInformation
    Exit Sub
    
ErrorHandler:
    MsgBox "Error en validación: " & Err.Description, vbCritical
End Sub

' ====================================
' IMPRESIÓN
' ====================================

Sub ImprimirMaestro()
    ' Imprime la hoja MAESTRO
    On Error GoTo ErrorHandler
    
    ThisWorkbook.Sheets(SHEET_MAESTRO).PrintOut
    Exit Sub
    
ErrorHandler:
    MsgBox "Error al imprimir: " & Err.Description, vbCritical
End Sub

Sub ImprimirCalificaciones()
    ' Imprime todas las hojas de calificaciones
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    
    For Each ws In ThisWorkbook.Sheets
        If Left(ws.Name, 1) <> "_" And ws.Name <> "Portada" And ws.Name <> SHEET_MAESTRO Then
            ws.PrintOut ' Imprime cada hoja de materias
        End If
    Next ws
    
    MsgBox "Impresión completada", vbInformation
    Exit Sub
    
ErrorHandler:
    MsgBox "Error al imprimir: " & Err.Description, vbCritical
End Sub

' ====================================
' UTILIDADES
' ====================================

Sub LimpiarFormatosTemporales()
    ' Limpia colores y formatos temporales
    On Error Resume Next
    ThisWorkbook.Sheets(SHEET_MAESTRO).Cells.Interior.Color = xlNone
End Sub

Function ObtenerFechaActual() As String
    ObtenerFechaActual = Format(Now(), "dd/mm/yyyy")
End Function

Function ObtenerHoraActual() As String
    ObtenerHoraActual = Format(Now(), "hh:mm:ss")
End Function
