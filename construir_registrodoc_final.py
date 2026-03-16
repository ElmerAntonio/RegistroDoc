import xlwings as xw
import shutil, os, time

# ============================================================
# REGISTRODOC MULTIGRADO - SCRIPT UNICO
# Toma el Excel original, corrige errores, agrega MAESTRO
# Tu agregas el .bas manualmente con Alt+F11
# Prof. Elmer Tugri - Panama 2026
# ============================================================

ORIGEN  = r"C:\RegistroDoc\RegistroDoc_Multigrado_v1_FINAL.xlsx"
SALIDA  = r"C:\RegistroDoc\RegistroDoc_v2.xlsm"

if not os.path.exists(ORIGEN):
    print(f"ERROR: No se encontro el archivo base:")
    print(f"  {ORIGEN}")
    input("\nPresiona Enter para salir...")
    exit()

print("=" * 60)
print("  RegistroDoc Multigrado v2.0")
print("  Construyendo sistema completo...")
print("=" * 60)

# ============================================================
# NOMBRES EXACTOS DE LAS HOJAS (del archivo original)
# ============================================================
PROM = {
    "Ingles 7":       "PROM (Ingles 7\u00b0)",
    "Ingles 8":       "PROM (Ingles 8\u00b0)",
    "Ingles 9":       "PROM (Ingles 9\u00b0)",
    "Hogar 7":        "PROM (Hogar y Desarrollo 7\u00b0)",
    "Comercio 8":     "PROM (Comercio 8\u00b0)",
    "Comercio 9":     "PROM (Comercio 9\u00b0)",
    "Artistica 9":    "PROM (Artistica 9\u00b0)",
    "Agro 7":         "PROM (Agropecuaria 7\u00b0)",
    "Agro 8":         "PROM (Agropecuaria 8\u00b0)",
    "Agro 9":         "PROM (Agropecuaria 9\u00b0)",
}
PLANILLA = {
    "Ingles 7":       "Planilla (Ingles 7\u00b0) ",
    "Ingles 8":       "Planilla (Ingles 8\u00b0)",
    "Ingles 9":       "Planilla (Ingles 9\u00b0) ",
    "Hogar 7":        "Planilla (Hogar y Desarrollo 7)",
    "Comercio 8":     "Planilla (Comercio 8\u00b0)",
    "Comercio 9":     "Planilla (Comercio 9\u00b0)",
    "Artistica 9":    "Planilla (Artistica 9\u00b0)",
    "Agro 7":         "Planilla (Agropecuaria 7\u00b0)",
    "Agro 8":         "Planilla (Agropecuaria 8\u00b0) ",
    "Agro 9":         "Planilla (Agropecuaria 9\u00b0)",
}
ASISTENCIA = {
    "7": "Asistencia (7\u00b0)",
    "8": "Asistencia (8\u00b0)",
    "9": "Asistencia (9\u00b0)",
}
GRADO = {
    "Ingles 7":"7","Hogar 7":"7","Agro 7":"7",
    "Ingles 8":"8","Comercio 8":"8","Agro 8":"8",
    "Ingles 9":"9","Comercio 9":"9","Artistica 9":"9","Agro 9":"9",
}

print("\n[1/6] Abriendo Excel...")
app = xw.App(visible=True)
app.display_alerts = False
wb = app.books.open(ORIGEN)
print("      OK")
time.sleep(1)

# ============================================================
# PASO 1 - CREAR HOJA MAESTRO
# ============================================================
print("\n[2/6] Creando hoja MAESTRO...")
names = [s.name for s in wb.sheets]
if "MAESTRO" not in names:
    wb.sheets.add("MAESTRO", before=wb.sheets[0])

ws_m = wb.sheets["MAESTRO"]
ws_m.range("A1:J60").clear()

# Colores helper
def bg(ws, rng, r, g, b):
    ws.range(rng).api.Interior.Color = r + g*256 + b*65536
def ft(ws, rng, r, g, b, bold=False, size=10):
    ws.range(rng).api.Font.Color = r + g*256 + b*65536
    ws.range(rng).api.Font.Bold = bold
    ws.range(rng).api.Font.Size = size
def al(ws, rng, h="Center"):
    m = {"Center":-4108,"Left":-4131,"Right":-4152}
    ws.range(rng).api.HorizontalAlignment = m[h]

# Titulo
ws_m.range("A1").value = "MAESTRO DE ESTUDIANTES - RegistroDoc Multigrado 2026"
ft(ws_m,"A1",31,56,100,True,14)
bg(ws_m,"A1:H1",255,242,204)

# Instruccion
ws_m.range("A2").value = "Escribe los nombres en formato: APELLIDO, Nombre  |  Col B = 7 Grado  |  Col D = 8 Grado  |  Col F = 9 Grado"
ft(ws_m,"A2",80,80,80,False,9)
ws_m.range("A2").api.Font.Italic = True

# Headers grados
for col, grado, r, g, b in [("B3","7 GRADO",31,56,100),("D3","8 GRADO",112,48,160),("F3","9 GRADO",55,86,35)]:
    ws_m.range(col).value = grado
    bg(ws_m,col,r,g,b)
    ft(ws_m,col,255,255,255,True,11)
    al(ws_m,col)

ws_m.range("A3").value = "No."
ft(ws_m,"A3",80,80,80,True,10)

# Numeros + celdas amarillas
for i in range(1,41):
    fila = 4 + i
    ws_m.range(f"A{fila}").value = i
    for col in ["B","D","F"]:
        bg(ws_m, f"{col}{fila}", 255,242,204)
        ws_m.range(f"{col}{fila}").api.Borders(9).LineStyle = 1

# Anchos
ws_m.api.Columns("A").ColumnWidth = 5
for c in ["B","D","F"]: ws_m.api.Columns(c).ColumnWidth = 28
for c in ["C","E","G"]: ws_m.api.Columns(c).ColumnWidth = 3

print("      MAESTRO creado OK")

# ============================================================
# PASO 2 - CONECTAR MAESTRO A PROM Y ASISTENCIA
# ============================================================
print("\n[3/6] Conectando MAESTRO con hojas PROM y Asistencia...")

for clave, prom_nombre in PROM.items():
    grado = GRADO[clave]
    mcol = ["B","D","F"][["7","8","9"].index(grado)]
    ws = wb.sheets[prom_nombre]
    for i in range(40):
        fm = 5 + i  # fila maestro
        fp = 5 + i  # fila prom
        ws.range(f"B{fp}").value = f"=MAESTRO!{mcol}{fm}"

for grado, asist_nombre in ASISTENCIA.items():
    mcol = ["B","D","F"][["7","8","9"].index(grado)]
    ws = wb.sheets[asist_nombre]
    for i in range(40):
        fm = 5 + i
        fa = 3 + i
        ws.range(f"B{fa}").value = f"=MAESTRO!{mcol}{fm}"

print("      Conexiones OK")

# ============================================================
# PASO 3 - CORREGIR PLANILLAS (nombres + ausencias + Hogar T3)
# ============================================================
print("\n[4/6] Corrigiendo Planillas...")

for clave, plan_nombre in PLANILLA.items():
    grado = GRADO[clave]
    prom_nombre  = PROM[clave]
    asist_nombre = ASISTENCIA[grado]
    mcol = ["B","D","F"][["7","8","9"].index(grado)]
    ws = wb.sheets[plan_nombre]

    for i in range(40):
        fp  = 16 + i   # fila planilla
        fpr = 5 + i    # fila prom
        fm  = 5 + i    # fila maestro
        fa  = 3 + i    # fila asistencia

        # Col C = nombre desde MAESTRO
        ws.range(f"C{fp}").value = f"=MAESTRO!{mcol}{fm}"

        # Col J = ausencias desde asistencia correcta
        ws.range(f"J{fp}").value = f"='{asist_nombre}'!BI{fa}"

# Error especial: Hogar 7 col H (T3) apuntaba a Ingles 7
ws_hogar = wb.sheets[PLANILLA["Hogar 7"]]
for i in range(40):
    fp  = 16 + i
    fpr = 5 + i
    ws_hogar.range(f"H{fp}").value = f"='{PROM['Hogar 7']}'!CU{fpr}"

print("      Planillas corregidas OK")

# ============================================================
# PASO 4 - CORREGIR FORMULAS PROM INCOMPLETAS
# ============================================================
print("\n[5/6] Corrigiendo formulas PROM...")

correcciones = {
    # Ingles 8: T1, T2, T3 le faltaba la prueba
    "PROM (Ingles 8\u00b0)": {
        "AE": "=TRUNC(AVERAGE(P{r},AA{r},AD{r}),1)",
        "BM": "=TRUNC(AVERAGE(AX{r},BI{r},BL{r}),1)",
        "CU": "=TRUNC(AVERAGE(CF{r},CQ{r},CT{r}),1)",
    },
    # Hogar 7, Comercio 8, Artistica 9: T2 le faltaba la prueba
    "PROM (Hogar y Desarrollo 7\u00b0)": {
        "BM": "=TRUNC(AVERAGE(AX{r},BI{r},BL{r}),1)",
    },
    "PROM (Comercio 8\u00b0)": {
        "BM": "=TRUNC(AVERAGE(AX{r},BI{r},BL{r}),1)",
    },
    "PROM (Artistica 9\u00b0)": {
        "BM": "=TRUNC(AVERAGE(AX{r},BI{r},BL{r}),1)",
    },
}

for prom_nombre, cols in correcciones.items():
    ws = wb.sheets[prom_nombre]
    for col, formula in cols.items():
        for fila in range(5, 45):
            ws.range(f"{col}{fila}").value = formula.replace("{r}", str(fila))
    print(f"      {prom_nombre.replace('PROM (','')} OK")

# ============================================================
# PASO 5 - AGREGAR VBA Y FORMS
# ============================================================
print("\n[6/6] Agregando VBA y Forms...")
vba = wb.api.VBProject

# ---- MODULO PRINCIPAL RD_CORE ----
m = vba.VBComponents.Add(1)
m.Name = "RD_CORE"
m.CodeModule.AddFromString("""
' ============================================================
' RD_CORE - Funciones de acceso a las hojas PROM originales
' Escribe directamente en las celdas correctas de MEDUCA
' Prof. Elmer Tugri - Panama 2026
' ============================================================

Function NombreHojaPROM(materia As String, grado As String) As String
    Dim clave As String: clave = materia & "|" & grado
    Select Case clave
        Case "Ingles|7":             NombreHojaPROM = "PROM (Ingles 7" & Chr(176) & ")"
        Case "Ingles|8":             NombreHojaPROM = "PROM (Ingles 8" & Chr(176) & ")"
        Case "Ingles|9":             NombreHojaPROM = "PROM (Ingles 9" & Chr(176) & ")"
        Case "Hogar y Desarrollo|7": NombreHojaPROM = "PROM (Hogar y Desarrollo 7" & Chr(176) & ")"
        Case "Comercio|8":           NombreHojaPROM = "PROM (Comercio 8" & Chr(176) & ")"
        Case "Comercio|9":           NombreHojaPROM = "PROM (Comercio 9" & Chr(176) & ")"
        Case "Ed. Artistica|9":      NombreHojaPROM = "PROM (Artistica 9" & Chr(176) & ")"
        Case "Agropecuaria|7":       NombreHojaPROM = "PROM (Agropecuaria 7" & Chr(176) & ")"
        Case "Agropecuaria|8":       NombreHojaPROM = "PROM (Agropecuaria 8" & Chr(176) & ")"
        Case "Agropecuaria|9":       NombreHojaPROM = "PROM (Agropecuaria 9" & Chr(176) & ")"
        Case Else:                   NombreHojaPROM = ""
    End Select
End Function

Function ColComponente(trimestre As Integer, componente As String) As Integer
    Select Case trimestre
        Case 1
            If componente = "Parciales"   Then ColComponente = 3
            If componente = "Apreciacion" Then ColComponente = 18
            If componente = "Prueba"      Then ColComponente = 28
        Case 2
            If componente = "Parciales"   Then ColComponente = 37
            If componente = "Apreciacion" Then ColComponente = 51
            If componente = "Prueba"      Then ColComponente = 62
        Case 3
            If componente = "Parciales"   Then ColComponente = 71
            If componente = "Apreciacion" Then ColComponente = 85
            If componente = "Prueba"      Then ColComponente = 96
    End Select
End Function

Function MaxNotas(componente As String) As Integer
    If componente = "Parciales"   Then MaxNotas = 13
    If componente = "Apreciacion" Then MaxNotas = 9
    If componente = "Prueba"      Then MaxNotas = 2
End Function

Function FilaEst(n As Integer) As Integer
    FilaEst = n + 4
End Function

Function NombreEst(n As Integer, grado As String) As String
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("MAESTRO")
    Dim col As Integer
    If grado = "7" Then col = 2
    If grado = "8" Then col = 4
    If grado = "9" Then col = 6
    NombreEst = Trim(CStr(ws.Cells(n + 4, col).Value))
End Function

Function ContarEst(grado As String) As Integer
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("MAESTRO")
    Dim col As Integer
    If grado = "7" Then col = 2
    If grado = "8" Then col = 4
    If grado = "9" Then col = 6
    Dim t As Integer: t = 0
    Dim i As Integer
    For i = 5 To 44
        If Trim(CStr(ws.Cells(i, col).Value)) <> "" Then t = t + 1
    Next i
    ContarEst = t
End Function

Function LeerNota(hoja As String, numEst As Integer, tri As Integer, comp As String, numNota As Integer) As Double
    Dim ws As Worksheet
    On Error Resume Next: Set ws = ThisWorkbook.Sheets(hoja): On Error GoTo 0
    If ws Is Nothing Then LeerNota = 0: Exit Function
    Dim v As Variant: v = ws.Cells(FilaEst(numEst), ColComponente(tri, comp) + numNota - 1).Value
    If IsNumeric(v) Then LeerNota = CDbl(v) Else LeerNota = 0
End Function

Sub EscribirNota(hoja As String, numEst As Integer, tri As Integer, comp As String, numNota As Integer, nota As Double)
    Dim ws As Worksheet
    On Error Resume Next: Set ws = ThisWorkbook.Sheets(hoja): On Error GoTo 0
    If ws Is Nothing Then MsgBox "Hoja no encontrada: " & hoja, vbExclamation, "Error": Exit Sub
    ws.Cells(FilaEst(numEst), ColComponente(tri, comp) + numNota - 1).Value = nota
End Sub

Function LeerFinal(hoja As String, numEst As Integer, tri As Integer) As Double
    Dim ws As Worksheet
    On Error Resume Next: Set ws = ThisWorkbook.Sheets(hoja): On Error GoTo 0
    If ws Is Nothing Then LeerFinal = 0: Exit Function
    Dim col As Integer
    If tri = 1 Then col = 31
    If tri = 2 Then col = 65
    If tri = 3 Then col = 99
    Dim v As Variant: v = ws.Cells(FilaEst(numEst), col).Value
    If IsNumeric(v) Then LeerFinal = CDbl(v) Else LeerFinal = 0
End Function

Function Materias(grado As String) As String()
    Dim arr() As String
    If grado = "7" Then
        ReDim arr(2)
        arr(0)="Ingles": arr(1)="Agropecuaria": arr(2)="Hogar y Desarrollo"
    ElseIf grado = "8" Then
        ReDim arr(2)
        arr(0)="Ingles": arr(1)="Agropecuaria": arr(2)="Comercio"
    ElseIf grado = "9" Then
        ReDim arr(3)
        arr(0)="Ingles": arr(1)="Agropecuaria": arr(2)="Comercio": arr(3)="Ed. Artistica"
    End If
    Materias = arr
End Function

Sub AbrirMenu()
    frmMenu.Show
End Sub

Sub GuardarRespaldo()
    Dim r As String
    r = ThisWorkbook.Path & "\\Respaldo_" & Format(Now,"YYYYMMDD_HHMM") & ".xlsm"
    ThisWorkbook.SaveCopyAs r
    MsgBox "Respaldo guardado:" & Chr(10) & r, vbInformation, "OK"
End Sub
""")
print("      RD_CORE OK")

# ---- THISWORKBOOK ----
vba.VBComponents("ThisWorkbook").CodeModule.AddFromString("""
Private Sub Workbook_Open()
    Application.Wait Now + TimeValue("00:00:01")
    frmMenu.Show
End Sub
""")

# ---- FORMS ----
forms = {
"frmMenu": """
Private Sub UserForm_Initialize()
    Me.Caption = "RegistroDoc Multigrado v2.0"
    Me.BackColor = RGB(26,26,46): Me.Width=460: Me.Height=560
End Sub
Private Sub btnMaestro_Click():    Me.Hide: frmMaestro.Show:    Me.Show: End Sub
Private Sub btnNotas_Click():      Me.Hide: frmNotas.Show:      Me.Show: End Sub
Private Sub btnAsistencia_Click(): Me.Hide: frmAsistencia.Show: Me.Show: End Sub
Private Sub btnResumenes_Click():  Me.Hide: frmResumenes.Show:  Me.Show: End Sub
Private Sub btnDashboard_Click():  Me.Hide: frmDashboard.Show:  Me.Show: End Sub
Private Sub btnReportes_Click():   Me.Hide: frmReportes.Show:   Me.Show: End Sub
Private Sub btnRespaldo_Click():   Call GuardarRespaldo: End Sub
Private Sub btnCerrar_Click()
    If MsgBox("Cerrar RegistroDoc?",vbYesNo+vbQuestion,"Cerrar")=vbYes Then
        ThisWorkbook.Save: ThisWorkbook.Close
    End If
End Sub
""",
"frmMaestro": """
Private Sub UserForm_Initialize()
    Me.Caption = "Estudiantes - Hoja MAESTRO"
    Me.BackColor = RGB(26,26,46): Me.Width=520: Me.Height=520
    cboGrado.AddItem "7": cboGrado.AddItem "8": cboGrado.AddItem "9"
    lblInfo.Caption = "Los nombres se escriben directamente en la hoja MAESTRO"
End Sub
Private Sub cboGrado_Change()
    lstEstudiantes.Clear
    Dim i As Integer
    For i = 1 To 40
        Dim nom As String: nom = NombreEst(i, cboGrado.Value)
        If nom <> "" Then
            lstEstudiantes.AddItem i & " | " & nom
        Else
            lstEstudiantes.AddItem i & " | (vacio)"
        End If
    Next i
    lblInfo.Caption = "Grado " & cboGrado.Value & " - " & ContarEst(cboGrado.Value) & " estudiantes registrados"
End Sub
Private Sub btnIrMaestro_Click()
    Me.Hide
    ThisWorkbook.Sheets("MAESTRO").Visible = True
    ThisWorkbook.Sheets("MAESTRO").Activate
    MsgBox "Estas en la hoja MAESTRO." & Chr(10) & Chr(10) & _
           "Escribe los nombres en las celdas AMARILLAS." & Chr(10) & _
           "Formato: APELLIDO, Nombre" & Chr(10) & Chr(10) & _
           "Columna B = Grado 7" & Chr(10) & _
           "Columna D = Grado 8" & Chr(10) & _
           "Columna F = Grado 9" & Chr(10) & Chr(10) & _
           "Cuando termines regresa al menu con el boton AbrirMenu.", _
           vbInformation, "Hoja MAESTRO"
End Sub
Private Sub btnCerrar_Click(): Me.Hide: End Sub
""",
"frmNotas": """
Private Sub UserForm_Initialize()
    Me.Caption = "Ingreso de Calificaciones - MEDUCA Panama"
    Me.BackColor = RGB(26,26,46): Me.Width=580: Me.Height=660
    cboGrado.AddItem "7": cboGrado.AddItem "8": cboGrado.AddItem "9"
    cboTrimestre.AddItem "1 - Primer Trimestre"
    cboTrimestre.AddItem "2 - Segundo Trimestre"
    cboTrimestre.AddItem "3 - Tercer Trimestre"
    cboComponente.AddItem "Parciales (hasta 13)"
    cboComponente.AddItem "Apreciacion (hasta 9)"
    cboComponente.AddItem "Prueba (hasta 2)"
    lblInfo.Caption = "Selecciona grado -> materia -> trimestre -> componente"
    lblFinal.Caption = ""
End Sub
Private Sub cboGrado_Change()
    cboMateria.Clear
    Dim mats() As String: mats = Materias(cboGrado.Value)
    Dim m As Variant
    For Each m In mats: cboMateria.AddItem CStr(m): Next m
    lstEstudiantes.Clear: lstNotas.Clear
    lblInfo.Caption = "Grado " & cboGrado.Value & " -> Selecciona materia"
End Sub
Private Sub cboMateria_Change():    Call Refrescar: End Sub
Private Sub cboTrimestre_Change():  Call Refrescar: End Sub
Private Sub cboComponente_Change(): Call Refrescar: End Sub
Private Sub lstEstudiantes_Click(): Call Refrescar: End Sub
Private Sub Refrescar()
    If cboGrado.Value="" Or cboMateria.Value="" Or cboTrimestre.Value="" Or cboComponente.Value="" Then Exit Sub
    Dim g As String: g = cboGrado.Value
    Dim mat As String: mat = cboMateria.Value
    Dim tri As Integer: tri = CInt(Left(cboTrimestre.Value,1))
    Dim comp As String
    If InStr(cboComponente.Value,"Parc") > 0 Then comp = "Parciales"
    If InStr(cboComponente.Value,"Aprec") > 0 Then comp = "Apreciacion"
    If InStr(cboComponente.Value,"Prueba") > 0 Then comp = "Prueba"
    Dim hoja As String: hoja = NombreHojaPROM(mat, g)
    If hoja = "" Then lblInfo.Caption = "ERROR: materia/grado no encontrado": Exit Sub
    Dim selEst As Integer: selEst = lstEstudiantes.ListIndex
    lstEstudiantes.Clear
    Dim i As Integer
    For i = 1 To 40
        Dim nom As String: nom = NombreEst(i, g)
        If nom <> "" Then lstEstudiantes.AddItem i & " | " & nom
    Next i
    If selEst >= 0 And selEst < lstEstudiantes.ListCount Then lstEstudiantes.ListIndex = selEst
    lstNotas.Clear
    If lstEstudiantes.ListIndex = -1 Then
        lblInfo.Caption = mat & " | Grado " & g & " | T" & tri & " | " & comp & " | Selecciona estudiante"
        Exit Sub
    End If
    Dim numEst As Integer: numEst = CInt(Split(lstEstudiantes.Value," | ")(0))
    Dim maxN As Integer: maxN = MaxNotas(comp)
    Dim cnt As Integer: cnt = 0
    For i = 1 To maxN
        Dim nota As Double: nota = LeerNota(hoja, numEst, tri, comp, i)
        If nota > 0 Then
            lstNotas.AddItem "Nota " & i & ":  " & Format(nota,"0.0")
            cnt = cnt + 1
        Else
            lstNotas.AddItem "Nota " & i & ":  (vacia)"
        End If
    Next i
    Dim pf As Double: pf = LeerFinal(hoja, numEst, tri)
    If pf > 0 Then
        Dim estado As String: If pf >= 3 Then estado = "APROBADO" Else estado = "REPROBADO"
        lblFinal.Caption = "PROMEDIO FINAL T" & tri & ": " & Format(pf,"0.0") & "  ->  " & estado
        If pf >= 3 Then lblFinal.ForeColor = RGB(100,220,100) Else lblFinal.ForeColor = RGB(255,100,100)
    Else
        lblFinal.Caption = "Promedio T" & tri & ": (sin notas aun)"
        lblFinal.ForeColor = RGB(150,150,150)
    End If
    lblInfo.Caption = mat & " | Grado " & g & " | T" & tri & " | " & comp & " | " & cnt & "/" & maxN & " ingresadas"
End Sub
Private Sub btnGuardar_Click()
    If cboGrado.Value="" Or cboMateria.Value="" Or cboTrimestre.Value="" Or cboComponente.Value="" Then _
        MsgBox "Selecciona grado, materia, trimestre y componente.",vbExclamation,"Error": Exit Sub
    If lstEstudiantes.ListIndex = -1 Then MsgBox "Selecciona un estudiante.",vbExclamation,"Error": Exit Sub
    If lstNotas.ListIndex = -1 Then MsgBox "Selecciona cual nota ingresar (Nota 1, 2...).",vbExclamation,"Error": Exit Sub
    If txtNota.Value = "" Then MsgBox "Escribe la nota (1.0 a 5.0).",vbExclamation,"Error": Exit Sub
    Dim nota As Double
    On Error GoTo ErrN: nota = CDbl(txtNota.Value): On Error GoTo 0
    If nota < 1 Or nota > 5 Then MsgBox "Nota entre 1.0 y 5.0",vbExclamation,"Error": Exit Sub
    Dim g As String: g = cboGrado.Value
    Dim mat As String: mat = cboMateria.Value
    Dim tri As Integer: tri = CInt(Left(cboTrimestre.Value,1))
    Dim comp As String
    If InStr(cboComponente.Value,"Parc") > 0 Then comp = "Parciales"
    If InStr(cboComponente.Value,"Aprec") > 0 Then comp = "Apreciacion"
    If InStr(cboComponente.Value,"Prueba") > 0 Then comp = "Prueba"
    Dim numNota As Integer: numNota = lstNotas.ListIndex + 1
    Dim numEst As Integer: numEst = CInt(Split(lstEstudiantes.Value," | ")(0))
    Dim hoja As String: hoja = NombreHojaPROM(mat, g)
    Call EscribirNota(hoja, numEst, tri, comp, numNota, nota)
    Call Refrescar
    txtNota.Value = ""
    Exit Sub
ErrN: MsgBox "Numero invalido. Ejemplo: 3.5",vbExclamation,"Error"
End Sub
Private Sub btnBorrarNota_Click()
    If lstEstudiantes.ListIndex=-1 Or lstNotas.ListIndex=-1 Then _
        MsgBox "Selecciona estudiante y nota.",vbExclamation,"Error": Exit Sub
    Dim g As String: g = cboGrado.Value
    Dim mat As String: mat = cboMateria.Value
    Dim tri As Integer: tri = CInt(Left(cboTrimestre.Value,1))
    Dim comp As String
    If InStr(cboComponente.Value,"Parc") > 0 Then comp = "Parciales"
    If InStr(cboComponente.Value,"Aprec") > 0 Then comp = "Apreciacion"
    If InStr(cboComponente.Value,"Prueba") > 0 Then comp = "Prueba"
    Dim hoja As String: hoja = NombreHojaPROM(mat, g)
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(hoja)
    Dim numEst As Integer: numEst = CInt(Split(lstEstudiantes.Value," | ")(0))
    Dim numNota As Integer: numNota = lstNotas.ListIndex + 1
    ws.Cells(FilaEst(numEst), ColComponente(tri, comp) + numNota - 1).ClearContents
    Call Refrescar
End Sub
Private Sub btnCerrar_Click(): Me.Hide: End Sub
""",
"frmAsistencia": """
Private Sub UserForm_Initialize()
    Me.Caption = "Control de Asistencia Diaria"
    Me.BackColor = RGB(26,26,46): Me.Width=520: Me.Height=560
    cboGrado.AddItem "7": cboGrado.AddItem "8": cboGrado.AddItem "9"
    txtFecha.Value = Format(Now,"DD/MM/YYYY")
    lblUltimo.Caption = ""
End Sub
Private Sub cboGrado_Change(): Call CargarLista: End Sub
Private Sub CargarLista()
    lstEstudiantes.Clear
    If cboGrado.Value = "" Then Exit Sub
    Dim i As Integer, t As Integer: t = 0
    For i = 1 To 40
        Dim nom As String: nom = NombreEst(i, cboGrado.Value)
        If nom <> "" Then
            Dim g As String: g = cboGrado.Value
            Dim asistNom As String
            If g="7" Then asistNom="Asistencia (7" & Chr(176) & ")"
            If g="8" Then asistNom="Asistencia (8" & Chr(176) & ")"
            If g="9" Then asistNom="Asistencia (9" & Chr(176) & ")"
            Dim wsA As Worksheet
            On Error Resume Next: Set wsA = ThisWorkbook.Sheets(asistNom): On Error GoTo 0
            Dim aus As Long: aus = 0
            Dim tar As Long: tar = 0
            If Not wsA Is Nothing Then
                Dim filaA As Integer: filaA = i + 2
                Dim vA As Variant: vA = wsA.Cells(filaA, 61).Value
                If IsNumeric(vA) Then aus = CLng(vA)
            End If
            Dim alerta As String: alerta = ""
            If aus > 5 Then alerta = " *** ALERTA: " & aus & " ausencias"
            lstEstudiantes.AddItem i & " | " & nom & "  [A:" & aus & "]" & alerta
            t = t + 1
        End If
    Next i
    lblContador.Caption = t & " estudiantes - Grado " & cboGrado.Value
End Sub
Private Sub btnP_Click(): Call Marcar("P"): End Sub
Private Sub btnA_Click(): Call Marcar("A"): End Sub
Private Sub btnT_Click(): Call Marcar("T"): End Sub
Private Sub Marcar(estado As String)
    If lstEstudiantes.ListIndex=-1 Then MsgBox "Selecciona un estudiante.",vbExclamation,"Error": Exit Sub
    If txtFecha.Value="" Then MsgBox "Ingresa la fecha.",vbExclamation,"Error": Exit Sub
    Dim g As String: g = cboGrado.Value
    Dim asistNom As String
    If g="7" Then asistNom="Asistencia (7" & Chr(176) & ")"
    If g="8" Then asistNom="Asistencia (8" & Chr(176) & ")"
    If g="9" Then asistNom="Asistencia (9" & Chr(176) & ")"
    Dim ws As Worksheet
    On Error Resume Next: Set ws = ThisWorkbook.Sheets(asistNom): On Error GoTo 0
    If ws Is Nothing Then MsgBox "Hoja no encontrada: " & asistNom,vbExclamation,"Error": Exit Sub
    Dim f As Date
    On Error GoTo ErrF: f = CDate(txtFecha.Value): On Error GoTo 0
    Dim numEst As Integer: numEst = CInt(Split(lstEstudiantes.Value," | ")(0))
    Dim filaA As Integer: filaA = numEst + 2
    Dim colF As Integer: colF = 0
    Dim c As Integer
    For c = 3 To 62
        Dim vF As Variant: vF = ws.Cells(2, c).Value
        If IsDate(vF) Then
            If CDate(vF) = f Then colF = c: Exit For
        End If
    Next c
    If colF = 0 Then
        For c = 3 To 62
            If IsEmpty(ws.Cells(2,c)) Or ws.Cells(2,c).Value="" Then
                colF = c: ws.Cells(2,c).Value = f: Exit For
            End If
        Next c
    End If
    If colF = 0 Then MsgBox "Sin columnas disponibles.",vbExclamation,"Error": Exit Sub
    ws.Cells(filaA, colF).Value = estado
    Dim nom As String: nom = Split(Split(lstEstudiantes.Value," | ")(1),"  [")(0)
    Dim es(2) As String: es(0)="PRESENTE": es(1)="AUSENTE": es(2)="TARDANZA"
    Dim idx As Integer: If estado="P" Then idx=0 Else If estado="A" Then idx=1 Else idx=2
    If estado="A" Then lblUltimo.ForeColor=RGB(255,100,100)
    If estado="P" Then lblUltimo.ForeColor=RGB(100,220,100)
    If estado="T" Then lblUltimo.ForeColor=RGB(255,200,50)
    lblUltimo.Caption = nom & " -> " & es(idx)
    Call CargarLista
    Exit Sub
ErrF: MsgBox "Fecha invalida. Formato: DD/MM/YYYY",vbExclamation,"Error"
End Sub
Private Sub btnTodosP_Click()
    If cboGrado.Value="" Or txtFecha.Value="" Then _
        MsgBox "Selecciona grado y fecha.",vbExclamation,"Error": Exit Sub
    If MsgBox("Marcar TODOS Presentes el " & txtFecha.Value & "?",vbYesNo+vbQuestion,"Confirmar")=vbNo Then Exit Sub
    Dim i As Integer
    For i = 0 To lstEstudiantes.ListCount-1
        lstEstudiantes.ListIndex = i
        If Not InStr(lstEstudiantes.Value,"vacio") > 0 Then Call Marcar("P")
    Next i
    MsgBox "Todos marcados Presentes.",vbInformation,"Listo"
End Sub
Private Sub btnCerrar_Click(): Me.Hide: End Sub
""",
"frmResumenes": """
Private Sub UserForm_Initialize()
    Me.Caption = "Resumenes de Notas Finales"
    Me.BackColor = RGB(26,26,46): Me.Width=620: Me.Height=570
    cboGrado.AddItem "7": cboGrado.AddItem "8": cboGrado.AddItem "9"
    cboTrimestre.AddItem "1 - Primer Trimestre"
    cboTrimestre.AddItem "2 - Segundo Trimestre"
    cboTrimestre.AddItem "3 - Tercer Trimestre"
    cboTrimestre.AddItem "Final - Promedio 3 Trimestres"
    lblStats.Caption = "Selecciona grado y trimestre para ver los resultados"
End Sub
Private Sub cboGrado_Change():     Call CargarResumen: End Sub
Private Sub cboTrimestre_Change(): Call CargarResumen: End Sub
Private Sub CargarResumen()
    If cboGrado.Value="" Or cboTrimestre.Value="" Then Exit Sub
    lstResumen.Clear
    Dim g As String: g = cboGrado.Value
    Dim mats() As String: mats = Materias(g)
    Dim apro As Long, repr As Long: apro=0: repr=0
    Dim i As Integer
    For i = 1 To 40
        Dim nom As String: nom = NombreEst(i, g)
        If nom = "" Then GoTo SigEst
        Dim linea As String: linea = i & " | " & nom & "   "
        If InStr(cboTrimestre.Value,"Final") > 0 Then
            Dim sumF As Double, cntF As Integer: sumF=0: cntF=0
            Dim m As Variant
            For Each m In mats
                Dim hp As String: hp = NombreHojaPROM(CStr(m), g)
                If hp <> "" Then
                    Dim t As Integer, sumT As Double, cntT As Integer: sumT=0: cntT=0
                    For t = 1 To 3
                        Dim pf As Double: pf = LeerFinal(hp, i, t)
                        If pf > 0 Then sumT=sumT+pf: cntT=cntT+1
                    Next t
                    If cntT > 0 Then sumF=sumF+(sumT/cntT): cntF=cntF+1
                End If
            Next m
            If cntF > 0 Then
                Dim prom As Double: prom = sumF/cntF
                Dim est As String: If prom>=3 Then est="APROBADO": apro=apro+1 Else est="REPROBADO": repr=repr+1
                linea = linea & "Prom: " & Format(prom,"0.0") & "  ->  " & est
            Else
                linea = linea & "Sin notas"
            End If
        Else
            Dim tri As Integer: tri = CInt(Left(cboTrimestre.Value,1))
            Dim m2 As Variant
            For Each m2 In mats
                Dim hp2 As String: hp2 = NombreHojaPROM(CStr(m2), g)
                If hp2 <> "" Then
                    Dim pf2 As Double: pf2 = LeerFinal(hp2, i, tri)
                    Dim ns As String: If pf2>0 Then ns=Format(pf2,"0.0") Else ns="---"
                    linea = linea & Left(CStr(m2),8) & ":" & ns & "  "
                End If
            Next m2
        End If
        lstResumen.AddItem linea
SigEst:
    Next i
    lblStats.Caption = "Grado " & g & "  |  Aprobados: " & apro & "  |  Reprobados: " & repr & "  |  Total: " & (apro+repr)
End Sub
Private Sub btnCerrar_Click(): Me.Hide: End Sub
""",
"frmDashboard": """
Private Sub UserForm_Initialize()
    Me.Caption = "Dashboard - Estadisticas y Alertas"
    Me.BackColor = RGB(26,26,46): Me.Width=520: Me.Height=520
    Call Refrescar
End Sub
Private Sub Refrescar()
    Dim t7 As Long, t8 As Long, t9 As Long
    t7=ContarEst("7"): t8=ContarEst("8"): t9=ContarEst("9")
    lblNum7.Caption=CStr(t7): lblNum8.Caption=CStr(t8)
    lblNum9.Caption=CStr(t9): lblNumTotal.Caption=CStr(t7+t8+t9)
    lstAlertas.Clear
    Dim g As Integer
    For g = 7 To 9
        Dim gStr As String: gStr = CStr(g)
        Dim asistNom As String
        asistNom = "Asistencia (" & gStr & Chr(176) & ")"
        Dim wsA As Worksheet
        On Error Resume Next: Set wsA = ThisWorkbook.Sheets(asistNom): On Error GoTo 0
        If Not wsA Is Nothing Then
            Dim i As Integer
            For i = 1 To 40
                Dim nom As String: nom = NombreEst(i, gStr)
                If nom <> "" Then
                    Dim filaA As Integer: filaA = i + 2
                    Dim aus As Variant: aus = wsA.Cells(filaA, 61).Value
                    If IsNumeric(aus) And CDbl(aus) > 5 Then
                        lstAlertas.AddItem "Grado " & gStr & " - " & nom & " -> " & aus & " ausencias"
                    End If
                End If
            Next i
        End If
    Next g
    lblNumAlertas.Caption = lstAlertas.ListCount & " estudiantes con mas de 5 ausencias"
End Sub
Private Sub btnRefrescar_Click(): Call Refrescar: End Sub
Private Sub btnCerrar_Click(): Me.Hide: End Sub
""",
"frmReportes": """
Private Sub UserForm_Initialize()
    Me.Caption = "Reportes Oficiales"
    Me.BackColor = RGB(26,26,46): Me.Width=480: Me.Height=380
End Sub
Private Sub btnReporteDir_Click()
    Dim t7 As Long, t8 As Long, t9 As Long
    t7=ContarEst("7"): t8=ContarEst("8"): t9=ContarEst("9")
    MsgBox "MINISTERIO DE EDUCACION DE PANAMA" & Chr(10) & _
           "SISTEMA EDUCATIVO PANAMENIO" & Chr(10) & String(50,"-") & Chr(10) & _
           "Centro Educativo: Escuela Cerro Cacicon" & Chr(10) & _
           "Docente: Elmer Tugri" & Chr(10) & _
           "Modalidad: Premedia Multigrado Flexible" & Chr(10) & _
           "Ano Lectivo: 2026" & Chr(10) & String(50,"-") & Chr(10) & _
           "MATRICULA ACTUAL:" & Chr(10) & _
           "  7 Grado: " & t7 & " estudiantes" & Chr(10) & _
           "  8 Grado: " & t8 & " estudiantes" & Chr(10) & _
           "  9 Grado: " & t9 & " estudiantes" & Chr(10) & _
           String(50,"-") & Chr(10) & "  TOTAL: " & (t7+t8+t9) & " estudiantes", _
           vbInformation, "Reporte para la Direccion"
End Sub
Private Sub btnVerPlanillas_Click()
    MsgBox "Las Planillas oficiales estan en las pestanas del Excel:" & Chr(10) & Chr(10) & _
           "Planilla (Ingles 7°) / (8°) / (9°)" & Chr(10) & _
           "Planilla (Hogar y Desarrollo 7)" & Chr(10) & _
           "Planilla (Comercio 8°) / (9°)" & Chr(10) & _
           "Planilla (Artistica 9°)" & Chr(10) & _
           "Planilla (Agropecuaria 7°) / (8°) / (9°)" & Chr(10) & Chr(10) & _
           "Para imprimir: abre la planilla y presiona Ctrl+P", _
           vbInformation, "Planillas Oficiales MEDUCA"
End Sub
Private Sub btnCerrar_Click(): Me.Hide: End Sub
""",
}

for nombre, codigo in forms.items():
    uf = vba.VBComponents.Add(3)
    uf.Name = nombre
    uf.CodeModule.AddFromString(codigo)
    print(f"      {nombre} OK")

# ---- MACRO PARA CONSTRUIR CONTROLES VISUALES ----
m_build = vba.VBComponents.Add(1)
m_build.Name = "RD_Build"
m_build.CodeModule.AddFromString("""
Sub ConstruirTodosLosControles()
    Call BMenu: Call BMaestro: Call BNotas
    Call BAsistencia: Call BResumenes
    Call BDashboard: Call BReportes
    MsgBox "Forms construidos correctamente." & Chr(10) & "El sistema esta listo.", vbInformation, "Listo"
End Sub

Private Function L(f As Object, nm As String, txt As String, l%, t%, w%, h%, fs%, fc As Long, bc As Long, al%) As Object
    Dim o As Object: Set o=f.Controls.Add("Forms.Label.1",nm,True)
    o.Caption=txt: o.Left=l: o.Top=t: o.Width=w: o.Height=h
    o.Font.Size=fs: o.ForeColor=fc: o.BackColor=bc: o.TextAlign=al
    Set L=o
End Function
Private Function B(f As Object, nm As String, txt As String, l%, t%, w%, h%, col As Long) As Object
    Dim o As Object: Set o=f.Controls.Add("Forms.CommandButton.1",nm,True)
    o.Caption=txt: o.Left=l: o.Top=t: o.Width=w: o.Height=h
    o.BackColor=col: o.ForeColor=RGB(255,255,255): o.Font.Bold=True: o.Font.Size=10
    Set B=o
End Function
Private Function T(f As Object, nm As String, l%, t%, w%, h%) As Object
    Dim o As Object: Set o=f.Controls.Add("Forms.TextBox.1",nm,True)
    o.Left=l: o.Top=t: o.Width=w: o.Height=h
    o.BackColor=RGB(46,46,70): o.ForeColor=RGB(255,255,255): o.Font.Size=10
    Set T=o
End Function
Private Function C(f As Object, nm As String, l%, t%, w%, h%) As Object
    Dim o As Object: Set o=f.Controls.Add("Forms.ComboBox.1",nm,True)
    o.Left=l: o.Top=t: o.Width=w: o.Height=h
    o.BackColor=RGB(46,46,70): o.ForeColor=RGB(255,255,255): o.Font.Size=10: o.Style=2
    Set C=o
End Function
Private Function LB(f As Object, nm As String, l%, t%, w%, h%) As Object
    Dim o As Object: Set o=f.Controls.Add("Forms.ListBox.1",nm,True)
    o.Left=l: o.Top=t: o.Width=w: o.Height=h
    o.BackColor=RGB(36,36,56): o.ForeColor=RGB(200,230,255): o.Font.Size=10
    Set LB=o
End Function
Private Sub Tarj(f As Object, nL As String, nN As String, tit As String, l%, t%, w%, h%, col As Long)
    Dim lb As Object: Set lb=f.Controls.Add("Forms.Label.1",nL,True)
    lb.Caption=tit: lb.Left=l: lb.Top=t: lb.Width=w: lb.Height=18
    lb.Font.Size=8: lb.Font.Bold=True: lb.ForeColor=RGB(255,255,255): lb.BackColor=col: lb.TextAlign=2
    Dim nb As Object: Set nb=f.Controls.Add("Forms.Label.1",nN,True)
    nb.Caption="0": nb.Left=l: nb.Top=t+18: nb.Width=w: nb.Height=h-18
    nb.Font.Size=22: nb.Font.Bold=True: nb.ForeColor=RGB(255,255,255): nb.BackColor=col: nb.TextAlign=2
End Sub
Private Sub Lim(f As Object)
    On Error Resume Next
    Dim c As Object: For Each c In f.Controls: f.Controls.Remove c.Name: Next c
    On Error GoTo 0
End Sub

Private Sub BMenu()
    Dim f As Object: Set f=ThisWorkbook.VBProject.VBComponents("frmMenu").Designer
    Call Lim(f)
    Call L(f,"lb0","RD",16,12,48,48,22,RGB(31,56,100),RGB(255,192,0),2)
    Call L(f,"lb1","RegistroDoc Multigrado",72,12,360,24,16,RGB(255,192,0),RGB(26,26,46),1)
    Call L(f,"lb2","Sistema de Registro Escolar - Panama 2026",72,36,360,14,9,RGB(189,215,238),RGB(26,26,46),1)
    Call L(f,"lbL","",0,64,460,4,8,RGB(255,192,0),RGB(255,192,0),1)
    Call L(f,"lb3","INGRESAR DATOS",0,76,460,14,8,RGB(255,192,0),RGB(26,26,46),2)
    Call B(f,"btnMaestro","Estudiantes / MAESTRO",10,92,210,46,RGB(46,117,182))
    Call B(f,"btnNotas","Ingresar Notas",240,92,200,46,RGB(112,48,160))
    Call B(f,"btnAsistencia","Asistencia Diaria",10,142,210,46,RGB(55,86,35))
    Call L(f,"lb4","VER RESULTADOS",0,196,460,14,8,RGB(255,192,0),RGB(26,26,46),2)
    Call B(f,"btnResumenes","Resumenes de Notas",10,212,210,44,RGB(46,117,182))
    Call B(f,"btnDashboard","Dashboard / Alertas",240,212,200,44,RGB(197,90,17))
    Call B(f,"btnReportes","Reportes Oficiales",10,260,210,44,RGB(55,86,35))
    Call L(f,"lb5","HERRAMIENTAS",0,312,460,14,8,RGB(255,192,0),RGB(26,26,46),2)
    Call B(f,"btnRespaldo","Guardar Respaldo",10,328,210,36,RGB(197,90,17))
    Call B(f,"btnCerrar","X  Cerrar Sistema",240,328,200,36,RGB(192,0,0))
    Call L(f,"lbC","c 2026 RegistroDoc Multigrado - Elmer Tugri - Escuela Cerro Cacicon",0,378,460,12,7,RGB(100,100,100),RGB(26,26,46),2)
End Sub

Private Sub BMaestro()
    Dim f As Object: Set f=ThisWorkbook.VBProject.VBComponents("frmMaestro").Designer
    Call Lim(f)
    Call L(f,"lbT","Estudiantes Registrados",0,8,520,22,13,RGB(255,192,0),RGB(26,26,46),2)
    Call L(f,"lbSub","Los nombres se escriben directamente en la hoja MAESTRO del Excel",0,30,520,14,9,RGB(189,215,238),RGB(26,26,46),2)
    Call L(f,"lbG","Ver grado:",10,52,75,18,9,RGB(189,215,238),RGB(26,26,46),1)
    Dim cbG As Object: Set cbG=C(f,"cboGrado",90,50,80,22)
    cbG.AddItem "7": cbG.AddItem "8": cbG.AddItem "9"
    Call L(f,"lblInfo","Selecciona un grado",186,52,320,18,9,RGB(100,220,100),RGB(26,26,46),1)
    Call LB(f,"lstEstudiantes",10,76,500,330)
    Call L(f,"lbInst","Para agregar o editar estudiantes usa la hoja MAESTRO:",10,414,340,14,9,RGB(255,192,0),RGB(26,26,46),1)
    Call B(f,"btnIrMaestro","Ir a Hoja MAESTRO",10,432,260,36,RGB(46,117,182))
    Call B(f,"btnCerrar","Cerrar",284,432,140,36,RGB(197,90,17))
End Sub

Private Sub BNotas()
    Dim f As Object: Set f=ThisWorkbook.VBProject.VBComponents("frmNotas").Designer
    Call Lim(f)
    Call L(f,"lbT","Ingreso de Calificaciones - Sistema MEDUCA Panama",0,8,580,22,13,RGB(255,192,0),RGB(26,26,46),2)
    Call L(f,"lbSub","Grado -> Materia -> Trimestre -> Componente -> Estudiante -> Nota",0,30,580,12,8,RGB(150,150,150),RGB(26,26,46),2)
    Call L(f,"lbG","Grado:",10,50,48,18,9,RGB(189,215,238),RGB(26,26,46),1)
    Call C(f,"cboGrado",62,48,70,22)
    Call L(f,"lbM","Materia:",145,50,55,18,9,RGB(189,215,238),RGB(26,26,46),1)
    Call C(f,"cboMateria",204,48,185,22)
    Call L(f,"lbTr","Trimestre:",403,50,68,18,9,RGB(189,215,238),RGB(26,26,46),1)
    Call C(f,"cboTrimestre",475,48,95,22)
    Call L(f,"lbCo","Componente:",10,76,80,18,9,RGB(189,215,238),RGB(26,26,46),1)
    Call C(f,"cboComponente",95,74,310,22)
    Call L(f,"lblInfo","Selecciona todos los campos para comenzar",0,100,580,16,9,RGB(100,200,255),RGB(15,15,35),2)
    Call L(f,"lbE","Estudiante:",10,120,80,14,8,RGB(255,192,0),RGB(26,26,46),1)
    Call L(f,"lbN","Notas del componente:",300,120,230,14,8,RGB(255,192,0),RGB(26,26,46),1)
    Call LB(f,"lstEstudiantes",10,136,282,350)
    Call LB(f,"lstNotas",300,136,268,350)
    Call L(f,"lblFinal","Promedio final aparece aqui",10,492,420,18,10,RGB(100,220,100),RGB(26,26,46),1)
    Call L(f,"lbNL","Nota (1.0-5.0):",10,518,108,18,9,RGB(189,215,238),RGB(26,26,46),1)
    Call T(f,"txtNota",122,516,82,24)
    Call B(f,"btnGuardar","Guardar Nota",218,514,165,28,RGB(55,86,35))
    Call B(f,"btnBorrarNota","Borrar Nota",396,514,115,28,RGB(197,90,17))
    Call B(f,"btnCerrar","Cerrar",490,550,82,28,RGB(192,0,0))
    Dim cbG2 As Object: Set cbG2=f.Controls("cboGrado")
    cbG2.AddItem "7": cbG2.AddItem "8": cbG2.AddItem "9"
    cbG2.BackColor=RGB(46,46,70): cbG2.ForeColor=RGB(255,255,255): cbG2.Style=2
    Dim cbTr As Object: Set cbTr=f.Controls("cboTrimestre")
    cbTr.AddItem "1 - Primer Trimestre": cbTr.AddItem "2 - Segundo Trimestre": cbTr.AddItem "3 - Tercer Trimestre"
    cbTr.BackColor=RGB(46,46,70): cbTr.ForeColor=RGB(255,255,255): cbTr.Style=2
    Dim cbCo As Object: Set cbCo=f.Controls("cboComponente")
    cbCo.AddItem "Parciales (hasta 13)": cbCo.AddItem "Apreciacion (hasta 9)": cbCo.AddItem "Prueba (hasta 2)"
    cbCo.BackColor=RGB(46,46,70): cbCo.ForeColor=RGB(255,255,255): cbCo.Style=2
End Sub

Private Sub BAsistencia()
    Dim f As Object: Set f=ThisWorkbook.VBProject.VBComponents("frmAsistencia").Designer
    Call Lim(f)
    Call L(f,"lbT","Control de Asistencia Diaria",0,8,520,22,13,RGB(255,192,0),RGB(26,26,46),2)
    Call L(f,"lbG","Grado:",10,40,48,18,9,RGB(189,215,238),RGB(26,26,46),1)
    Dim cbG3 As Object: Set cbG3=C(f,"cboGrado",62,38,80,22)
    cbG3.AddItem "7": cbG3.AddItem "8": cbG3.AddItem "9"
    cbG3.BackColor=RGB(46,46,70): cbG3.ForeColor=RGB(255,255,255): cbG3.Style=2
    Call L(f,"lbF","Fecha (DD/MM/YYYY):",158,40,135,18,9,RGB(189,215,238),RGB(26,26,46),1)
    Call T(f,"txtFecha",298,38,134,22)
    Call LB(f,"lstEstudiantes",10,68,500,300)
    Call L(f,"lblContador","0 estudiantes",10,376,300,16,9,RGB(100,200,255),RGB(26,26,46),1)
    Call L(f,"lblUltimo","",10,394,500,18,10,RGB(100,220,100),RGB(26,26,46),1)
    Call B(f,"btnP","P  PRESENTE",10,416,158,50,RGB(55,86,35))
    Call B(f,"btnA","A  AUSENTE",178,416,158,50,RGB(192,0,0))
    Call B(f,"btnT","T  TARDANZA",346,416,158,50,RGB(184,134,11))
    Call B(f,"btnTodosP","Marcar Todos Presentes hoy",10,474,250,30,RGB(46,117,182))
    Call B(f,"btnCerrar","Cerrar",276,474,150,30,RGB(197,90,17))
End Sub

Private Sub BResumenes()
    Dim f As Object: Set f=ThisWorkbook.VBProject.VBComponents("frmResumenes").Designer
    Call Lim(f)
    Call L(f,"lbT","Resumenes de Notas - Conectado con PROM MEDUCA",0,8,620,22,13,RGB(255,192,0),RGB(26,26,46),2)
    Call L(f,"lbG","Grado:",10,40,48,18,9,RGB(189,215,238),RGB(26,26,46),1)
    Dim cbG4 As Object: Set cbG4=C(f,"cboGrado",62,38,80,22)
    cbG4.AddItem "7": cbG4.AddItem "8": cbG4.AddItem "9"
    cbG4.BackColor=RGB(46,46,70): cbG4.ForeColor=RGB(255,255,255): cbG4.Style=2
    Call L(f,"lbTr","Trimestre:",156,40,70,18,9,RGB(189,215,238),RGB(26,26,46),1)
    Dim cbTr4 As Object: Set cbTr4=C(f,"cboTrimestre",230,38,200,22)
    cbTr4.AddItem "1 - Primer Trimestre": cbTr4.AddItem "2 - Segundo Trimestre"
    cbTr4.AddItem "3 - Tercer Trimestre": cbTr4.AddItem "Final - Promedio 3 Trimestres"
    cbTr4.BackColor=RGB(46,46,70): cbTr4.ForeColor=RGB(255,255,255): cbTr4.Style=2
    Call L(f,"lblStats","Selecciona grado y trimestre",440,40,168,18,9,RGB(100,200,255),RGB(26,26,46),1)
    Call LB(f,"lstResumen",10,66,600,438)
    Call B(f,"btnCerrar","Cerrar",230,512,160,34,RGB(46,117,182))
End Sub

Private Sub BDashboard()
    Dim f As Object: Set f=ThisWorkbook.VBProject.VBComponents("frmDashboard").Designer
    Call Lim(f)
    Call L(f,"lbT","Dashboard - Estadisticas y Alertas",0,8,520,22,13,RGB(255,192,0),RGB(26,26,46),2)
    Call L(f,"lbS1","MATRICULA ACTUAL POR GRADO",0,36,520,14,8,RGB(255,192,0),RGB(26,26,46),2)
    Call Tarj(f,"lb7t","lblNum7","7 Grado",10,52,110,54,RGB(46,117,182))
    Call Tarj(f,"lb8t","lblNum8","8 Grado",130,52,110,54,RGB(112,48,160))
    Call Tarj(f,"lb9t","lblNum9","9 Grado",250,52,110,54,RGB(197,90,17))
    Call Tarj(f,"lbGt","lblNumTotal","TOTAL",370,52,122,54,RGB(55,86,35))
    Call L(f,"lbS2","ALERTAS - Estudiantes con mas de 5 ausencias",0,116,520,14,8,RGB(255,150,50),RGB(26,26,46),2)
    Call LB(f,"lstAlertas",10,134,500,256)
    Call L(f,"lblNumAlertas","0 alertas",10,398,300,16,9,RGB(255,150,50),RGB(26,26,46),1)
    Call B(f,"btnRefrescar","Refrescar datos",10,420,185,32,RGB(55,86,35))
    Call B(f,"btnCerrar","Cerrar",210,420,165,32,RGB(46,117,182))
End Sub

Private Sub BReportes()
    Dim f As Object: Set f=ThisWorkbook.VBProject.VBComponents("frmReportes").Designer
    Call Lim(f)
    Call L(f,"lbT","Reportes Oficiales MEDUCA Panama",0,8,480,22,13,RGB(255,192,0),RGB(26,26,46),2)
    Call L(f,"lbS","Genera reportes usando los datos actuales del sistema:",10,38,460,16,9,RGB(189,215,238),RGB(26,26,46),1)
    Call B(f,"btnReporteDir","Reporte de Matricula para la Direccion",10,64,460,52,RGB(31,56,100))
    Call B(f,"btnVerPlanillas","Ver Planillas Oficiales (imprimir desde Excel)",10,128,460,52,RGB(55,86,35))
    Call L(f,"lbN","Las Planillas se imprimen directamente desde las pestanas del Excel (Ctrl+P).",10,194,460,14,9,RGB(150,150,150),RGB(26,26,46),2)
    Call B(f,"btnCerrar","Cerrar",160,220,160,32,RGB(197,90,17))
End Sub
""")

print("\n     Ejecutando ConstruirTodosLosControles...")
try:
    app.screen_updating = True
    app.api.Run("ConstruirTodosLosControles")
    print("     Controles construidos OK")
except Exception as e:
    print(f"     Nota: {e}")

time.sleep(2)

print(f"\nGuardando como: {SALIDA}")
wb.api.SaveAs(SALIDA, FileFormat=52)
print("OK")
time.sleep(1)
wb.close()
app.quit()

print("\n" + "=" * 60)
print("  COMPLETADO - RegistroDoc v2.0")
print(f"  Archivo: {SALIDA}")
print("=" * 60)
print("""
  CORRECCIONES APLICADAS AL EXCEL ORIGINAL:
  + Hoja MAESTRO creada (nombres en un solo lugar)
  + MAESTRO conectado a todas las hojas PROM y Asistencia
  + Planillas: nombres y ausencias corregidos
  + Hogar 7 T3: apuntaba a Ingles, ahora corregido
  + PROM Ingles 8: T1+T2+T3 completos con 3 componentes
  + PROM Hogar 7, Comercio 8, Artistica 9: T2 completo

  FORMS CONECTADOS CON EL EXCEL ORIGINAL:
  + frmMenu      -> menu principal con navegacion
  + frmMaestro   -> ve y edita la hoja MAESTRO
  + frmNotas     -> escribe en celdas PROM reales de MEDUCA
  + frmAsistencia-> escribe en hojas Asistencia reales
  + frmResumenes -> lee promedios finales de PROM
  + frmDashboard -> estadisticas y alertas en tiempo real
  + frmReportes  -> reportes y acceso a planillas

  SIGUIENTE PASO:
  Abre RegistroDoc_v2.xlsm y habilita macros.
  El menu aparece automaticamente.
""")
input("Presiona Enter para cerrar...")
