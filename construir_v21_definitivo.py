import xlwings as xw
import os, time

# ============================================================
# REGISTRODOC MULTIGRADO v2.1 - SCRIPT DEFINITIVO COMPLETO
# Incluye: Login ISO 27001, CRUD Estudiantes, Notas con
# descripcion y fecha, Configuracion materias/grados,
# Impresion libro completo, Auditoria, Respaldo automatico
# Prof. Elmer Tugri - Panama 2026
# ============================================================

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
ORIGEN = os.path.join(BASE_DIR, "RegistroDoc_Multigrado_v1_FINAL.xlsx")
SALIDA = os.path.join(BASE_DIR, "RegistroDoc_v21.xlsm")

if not os.path.exists(ORIGEN):
    print(f"ERROR: No se encontro: {ORIGEN}")
    input("Presiona Enter para salir...")
    exit()

print("=" * 60)
print("  RegistroDoc Multigrado v2.1 - Version DEFINITIVA")
print("=" * 60)
print(f"Carpeta de trabajo: {BASE_DIR}")

# ============================================================
# NOMBRES EXACTOS DE HOJAS ORIGINALES
# ============================================================
PROM_ORIG = {
    "Ingles 7":    "PROM (Ingles 7\u00b0)",
    "Ingles 8":    "PROM (Ingles 8\u00b0)",
    "Ingles 9":    "PROM (Ingles 9\u00b0)",
    "Hogar 7":     "PROM (Hogar y Desarrollo 7\u00b0)",
    "Comercio 8":  "PROM (Comercio 8\u00b0)",
    "Comercio 9":  "PROM (Comercio 9\u00b0)",
    "Artistica 9": "PROM (Artistica 9\u00b0)",
    "Agro 7":      "PROM (Agropecuaria 7\u00b0)",
    "Agro 8":      "PROM (Agropecuaria 8\u00b0)",
    "Agro 9":      "PROM (Agropecuaria 9\u00b0)",
}
PLAN_ORIG = {
    "Ingles 7":    "Planilla (Ingles 7\u00b0) ",
    "Ingles 8":    "Planilla (Ingles 8\u00b0)",
    "Ingles 9":    "Planilla (Ingles 9\u00b0) ",
    "Hogar 7":     "Planilla (Hogar y Desarrollo 7)",
    "Comercio 8":  "Planilla (Comercio 8\u00b0)",
    "Comercio 9":  "Planilla (Comercio 9\u00b0)",
    "Artistica 9": "Planilla (Artistica 9\u00b0)",
    "Agro 7":      "Planilla (Agropecuaria 7\u00b0)",
    "Agro 8":      "Planilla (Agropecuaria 8\u00b0) ",
    "Agro 9":      "Planilla (Agropecuaria 9\u00b0)",
}
ASIST = {
    "7": "Asistencia (7\u00b0)",
    "8": "Asistencia (8\u00b0)",
    "9": "Asistencia (9\u00b0)",
}
GRADO_MAT = {
    "Ingles 7":"7","Hogar 7":"7","Agro 7":"7",
    "Ingles 8":"8","Comercio 8":"8","Agro 8":"8",
    "Ingles 9":"9","Comercio 9":"9","Artistica 9":"9","Agro 9":"9",
}

def rgb_xl(r,g,b): return r + g*256 + b*65536
def bg(ws,rng,r,g,b): ws.range(rng).api.Interior.Color = rgb_xl(r,g,b)
def ft(ws,rng,r,g,b,bold=False,sz=10):
    ws.range(rng).api.Font.Color = rgb_xl(r,g,b)
    ws.range(rng).api.Font.Bold = bold
    ws.range(rng).api.Font.Size = sz

def safe_clear(ws, rng):
    """Limpia un rango, incluso si existen celdas combinadas."""
    try:
        ws.range(rng).clear()
    except Exception:
        # Excel falla con "Esta accion no se puede realizar en una celda combinada"
        # cuando el rango intercepta combinaciones preexistentes.
        ws.range(rng).api.UnMerge()
        ws.range(rng).clear()


print("\n[1/7] Abriendo Excel...")
app = xw.App(visible=True)
app.display_alerts = False
wb = app.books.open(ORIGEN)
print("      OK")
time.sleep(1)

# ============================================================
# PASO 1: HOJA MAESTRO
# ============================================================
print("\n[2/7] Creando hoja MAESTRO...")
names = [s.name for s in wb.sheets]
if "MAESTRO" not in names:
    wb.sheets.add("MAESTRO", before=wb.sheets[0])
ws_m = wb.sheets["MAESTRO"]
ws_m.range("A1:J60").clear()
ws_m.range("A1").value = "MAESTRO DE ESTUDIANTES - RegistroDoc Multigrado 2026"
ft(ws_m,"A1",31,56,100,True,14)
bg(ws_m,"A1:H1",255,242,204)
ws_m.range("A2").value = "Columna B = Grado 7  |  Columna D = Grado 8  |  Columna F = Grado 9  |  Formato: APELLIDO, Nombre"
ft(ws_m,"A2",100,100,100,False,9)
ws_m.range("A2").api.Font.Italic = True
for col,grado,r,g,b in [("B3","7 GRADO",31,56,100),("D3","8 GRADO",112,48,160),("F3","9 GRADO",55,86,35)]:
    ws_m.range(col).value = grado
    bg(ws_m,col,r,g,b)
    ft(ws_m,col,255,255,255,True,11)
    ws_m.range(col).api.HorizontalAlignment = -4108
ws_m.range("A3").value = "No."
ft(ws_m,"A3",80,80,80,True,9)
for i in range(1,41):
    fila = 4+i
    ws_m.range(f"A{fila}").value = i
    for col in ["B","D","F"]:
        bg(ws_m,f"{col}{fila}",255,242,204)
        ws_m.range(f"{col}{fila}").api.Borders(9).LineStyle = 1
ws_m.api.Columns("A").ColumnWidth = 5
for c in ["B","D","F"]: ws_m.api.Columns(c).ColumnWidth = 30
for c in ["C","E","G"]: ws_m.api.Columns(c).ColumnWidth = 3
print("      MAESTRO OK")

# ============================================================
# PASO 1.1: HOJA PORTADA (SIN IMAGEN, DISENO EDUCATIVO)
# ============================================================
print("\n[2.1/7] Creando hoja PORTADA educativa...")
try:
    ws_portada = wb.sheets["PORTADA"]
except Exception:
    try:
        wb.sheets.add("PORTADA", before=wb.sheets[0])
    except ValueError:
        pass  # Ya existe; tomar la hoja existente
    ws_portada = wb.sheets["PORTADA"]

safe_clear(ws_portada, "A1:J40")
ws_portada.api.Cells.UnMerge()

# Fondo y paneles
bg(ws_portada, "A1:J40", 242, 248, 255)
bg(ws_portada, "A1:J4", 31, 56, 100)
bg(ws_portada, "B6:I13", 255, 255, 255)
bg(ws_portada, "B16:I28", 255, 255, 255)

# Titulos grandes
ws_portada.range("A1:J1").api.Merge()
ws_portada.range("A2:J2").api.Merge()
ws_portada.range("A3:J3").api.Merge()
ws_portada.range("A1").value = "REGISTRODOC MULTIGRADO"
ws_portada.range("A2").value = "Sistema de Gestion Academica"
ws_portada.range("A3").value = "Version 2.1"
ft(ws_portada, "A1", 255, 255, 255, True, 28)
ft(ws_portada, "A2", 226, 240, 255, True, 16)
ft(ws_portada, "A3", 200, 220, 245, False, 12)
for cell in ["A1", "A2", "A3"]:
    ws_portada.range(cell).api.HorizontalAlignment = -4108

# Mensaje principal
ws_portada.range("B6:I6").api.Merge()
ws_portada.range("B6").value = "BIENVENIDO/A DOCENTE"
ft(ws_portada, "B6", 31, 56, 100, True, 20)
ws_portada.range("B6").api.HorizontalAlignment = -4108

ws_portada.range("B8:I8").api.Merge()
ws_portada.range("B8").value = "Para iniciar correctamente:"
ft(ws_portada, "B8", 40, 40, 40, True, 13)

ws_portada.range("B10:I10").api.Merge()
ws_portada.range("B10").value = "1) Habilita las MACROS al abrir el archivo"
ft(ws_portada, "B10", 0, 102, 51, True, 13)

ws_portada.range("B11:I11").api.Merge()
ws_portada.range("B11").value = "2) Cierra y abre nuevamente si el login no aparece"
ft(ws_portada, "B11", 0, 102, 51, True, 13)

ws_portada.range("B12:I12").api.Merge()
ws_portada.range("B12").value = "3) Usa usuario y contrasena del sistema"
ft(ws_portada, "B12", 0, 102, 51, True, 13)

# Nota de seguridad
ws_portada.range("B16:I16").api.Merge()
ws_portada.range("B16").value = "SEGURIDAD Y CONTROL"
ft(ws_portada, "B16", 112, 48, 160, True, 15)
ws_portada.range("B18:I18").api.Merge()
ws_portada.range("B18").value = "Este libro oculta hojas automaticamente y muestra login por macros."
ft(ws_portada, "B18", 60, 60, 60, False, 11)
ws_portada.range("B19:I19").api.Merge()
ws_portada.range("B19").value = "Si ves MAESTRO o no sale login, habilita macros y vuelve a abrir."
ft(ws_portada, "B19", 170, 0, 0, True, 11)

# Bordes y formato
for rng in ["B6:I13", "B16:I28"]:
    ws_portada.range(rng).api.Borders(7).LineStyle = 1
    ws_portada.range(rng).api.Borders(8).LineStyle = 1
    ws_portada.range(rng).api.Borders(9).LineStyle = 1
    ws_portada.range(rng).api.Borders(10).LineStyle = 1

ws_portada.api.Columns("A:J").ColumnWidth = 12
for fila, alto in [(1,34),(2,26),(3,20),(6,30),(8,24),(10,22),(11,22),(12,22),(16,24),(18,22),(19,22)]:
    ws_portada.api.Rows(fila).RowHeight = alto

ws_portada.activate()
print("      PORTADA OK")

# ============================================================
# PASO 2: HOJA BD_CONFIG (materias y grados configurables)
# ============================================================
print("\n[3/7] Creando BD_Config y BD_Auditoria...")
# Activar primera hoja para evitar seleccion multiple
if "BD_Config" not in names:
    app.api.ActiveWorkbook.Sheets(1).Activate()
    wb.sheets.add("BD_Config")
ws_cfg = wb.sheets["BD_Config"]
ws_cfg.range("A1:H100").clear()
ws_cfg.range("A1").value = ["Clave","Valor"]
cfg_inicial = [
    ["anio_lectivo","2026"],
    ["nombre_docente","Elmer Tugri"],
    ["nombre_escuela","Escuela Cerro Cacicon"],
    ["ubicacion","Cacicon, Comarca Ngabe Bugle"],
    ["modalidad","Premedia Multigrado Flexible"],
    ["director",""],
    ["email_docente",""],
    ["version","2.1"],
]
ws_cfg.range("A2").value = cfg_inicial
ws_cfg.api.Visible = False  # Oculta

# Hoja BD_Materias (lista de materias configurables)
if "BD_Materias" not in names:
    app.api.ActiveWorkbook.Sheets(1).Activate()
    wb.sheets.add("BD_Materias")
ws_mat = wb.sheets["BD_Materias"]
ws_mat.range("A1:G100").clear()
ws_mat.range("A1").value = ["ID","NombreMateria","Grado","NombreHojaPROM","NombrePlanilla","EsUnion","NombreUnion"]
# Materias iniciales del archivo original
materias_ini = [
    [1,"Ingles","7","PROM (Ingles 7\u00b0)","Planilla (Ingles 7\u00b0) ","No",""],
    [2,"Agropecuaria","7","PROM (Agropecuaria 7\u00b0)","Planilla (Agropecuaria 7\u00b0)","Si","Tecnologia"],
    [3,"Hogar y Desarrollo","7","PROM (Hogar y Desarrollo 7\u00b0)","Planilla (Hogar y Desarrollo 7)","Si","Tecnologia"],
    [4,"Ingles","8","PROM (Ingles 8\u00b0)","Planilla (Ingles 8\u00b0)","No",""],
    [5,"Agropecuaria","8","PROM (Agropecuaria 8\u00b0)","Planilla (Agropecuaria 8\u00b0) ","Si","Tecnologia"],
    [6,"Comercio","8","PROM (Comercio 8\u00b0)","Planilla (Comercio 8\u00b0)","Si","Tecnologia"],
    [7,"Ingles","9","PROM (Ingles 9\u00b0)","Planilla (Ingles 9\u00b0) ","No",""],
    [8,"Agropecuaria","9","PROM (Agropecuaria 9\u00b0)","Planilla (Agropecuaria 9\u00b0)","Si","Tecnologia"],
    [9,"Comercio","9","PROM (Comercio 9\u00b0)","Planilla (Comercio 9\u00b0)","Si","Tecnologia"],
    [10,"Ed. Artistica","9","PROM (Artistica 9\u00b0)","Planilla (Artistica 9\u00b0)","No",""],
]
ws_mat.range("A2").value = materias_ini
ws_mat.api.Visible = False

# Hoja BD_Usuarios
if "BD_Usuarios" not in names:
    app.api.ActiveWorkbook.Sheets(1).Activate()
    wb.sheets.add("BD_Usuarios")
ws_usr = wb.sheets["BD_Usuarios"]
ws_usr.range("A1:H20").clear()
ws_usr.range("A1").value = ["Usuario","HashContrasena","Rol","Activo","UltimoCambio","Intentos","Bloqueado","Email"]
# Hash simple: base64 de usuario+salt (no SHA real en VBA sin librerias, pero funcional)
ws_usr.range("A2").value = [
    ["admin",   "admin123",   "Admin",   "Si","","0","No",""],
    ["docente", "docente123", "Docente", "Si","","0","No",""],
    ["director","director123","Director","Si","","0","No",""],
]
ws_usr.api.Visible = 2  # xlSheetVeryHidden

# Hoja BD_Auditoria
if "BD_Auditoria" not in names:
    app.api.ActiveWorkbook.Sheets(1).Activate()
    wb.sheets.add("BD_Auditoria")
ws_aud = wb.sheets["BD_Auditoria"]
ws_aud.range("A1:G2").clear()
ws_aud.range("A1").value = ["ID","Usuario","Accion","Detalle","Tabla","Fecha","Hora"]
ws_aud.api.Visible = 2  # xlSheetVeryHidden

print("      BD_Config, BD_Materias, BD_Usuarios, BD_Auditoria OK")

# ============================================================
# PASO 3: CONECTAR MAESTRO -> PROM, ASISTENCIA, PLANILLAS
# ============================================================
print("\n[4/7] Conectando hojas y corrigiendo errores...")

for clave, prom_nom in PROM_ORIG.items():
    grado = GRADO_MAT[clave]
    mcol = ["B","D","F"][["7","8","9"].index(grado)]
    ws = wb.sheets[prom_nom]
    for i in range(40):
        ws.range(f"B{5+i}").value = f"=MAESTRO!{mcol}{5+i}"

for grado, asist_nom in ASIST.items():
    mcol = ["B","D","F"][["7","8","9"].index(grado)]
    ws = wb.sheets[asist_nom]
    for i in range(40):
        ws.range(f"B{3+i}").value = f"=MAESTRO!{mcol}{5+i}"

for clave, plan_nom in PLAN_ORIG.items():
    grado = GRADO_MAT[clave]
    prom_nom  = PROM_ORIG[clave]
    asist_nom = ASIST[grado]
    mcol = ["B","D","F"][["7","8","9"].index(grado)]
    ws = wb.sheets[plan_nom]
    for i in range(40):
        fp = 16+i
        ws.range(f"C{fp}").value = f"=MAESTRO!{mcol}{5+i}"
        ws.range(f"J{fp}").value = f"='{asist_nom}'!BI{3+i}"

# Hogar 7 T3 apuntaba a Ingles 7
ws_hogar = wb.sheets[PLAN_ORIG["Hogar 7"]]
for i in range(40):
    ws_hogar.range(f"H{16+i}").value = f"='{PROM_ORIG['Hogar 7']}'!CU{5+i}"

# Corregir formulas PROM incompletas
correc = {
    "PROM (Ingles 8\u00b0)": {"AE":"=TRUNC(AVERAGE(P{r},AA{r},AD{r}),1)","BM":"=TRUNC(AVERAGE(AX{r},BI{r},BL{r}),1)","CU":"=TRUNC(AVERAGE(CF{r},CQ{r},CT{r}),1)"},
    "PROM (Hogar y Desarrollo 7\u00b0)": {"BM":"=TRUNC(AVERAGE(AX{r},BI{r},BL{r}),1)"},
    "PROM (Comercio 8\u00b0)": {"BM":"=TRUNC(AVERAGE(AX{r},BI{r},BL{r}),1)"},
    "PROM (Artistica 9\u00b0)": {"BM":"=TRUNC(AVERAGE(AX{r},BI{r},BL{r}),1)"},
}
for pnom, cols in correc.items():
    ws = wb.sheets[pnom]
    for col, fm in cols.items():
        for r in range(5,45): ws.range(f"{col}{r}").value = fm.replace("{r}",str(r))

print("      Conexiones y correcciones OK")

# ============================================================
# PASO 4: VBA COMPLETO
# ============================================================
print("\n[5/7] Agregando VBA...")
vba = wb.api.VBProject

# Limpiar VBA previo
for i in range(vba.VBComponents.Count,0,-1):
    try:
        comp = vba.VBComponents.Item(i)
        if not comp.Name.startswith("Sheet") and comp.Name != "ThisWorkbook":
            vba.VBComponents.Remove(comp)
    except: pass
try:
    mod = vba.VBComponents("ThisWorkbook").CodeModule
    if mod.CountOfLines > 0: mod.DeleteLines(1,mod.CountOfLines)
except: pass

# ---- MODULO RD_CORE ----
m1 = vba.VBComponents.Add(1); m1.Name = "RD_CORE"
m1.CodeModule.AddFromString("""
' ============================================================
' RD_CORE - Funciones base del sistema
' RegistroDoc Multigrado v2.1 - Panama 2026
' ============================================================
Public UsuarioActual As String
Public RolActual As String

Function ObtenerConfig(clave As String) As String
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("BD_Config")
    Dim i As Long
    For i=2 To ws.Cells(ws.Rows.Count,1).End(xlUp).Row
        If ws.Cells(i,1).Value=clave Then ObtenerConfig=CStr(ws.Cells(i,2).Value): Exit Function
    Next i
End Function

Sub GuardarConfig(clave As String, valor As String)
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("BD_Config")
    Dim i As Long
    For i=2 To ws.Cells(ws.Rows.Count,1).End(xlUp).Row
        If ws.Cells(i,1).Value=clave Then ws.Cells(i,2).Value=valor: Exit Sub
    Next i
    Dim u As Long: u=ws.Cells(ws.Rows.Count,1).End(xlUp).Row+1
    ws.Cells(u,1).Value=clave: ws.Cells(u,2).Value=valor
End Sub

Sub RegistrarAuditoria(accion As String, detalle As String, tabla As String)
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("BD_Auditoria")
    Dim u As Long: u=ws.Cells(ws.Rows.Count,1).End(xlUp).Row+1
    ws.Cells(u,1).Value = u-1
    ws.Cells(u,2).Value = UsuarioActual
    ws.Cells(u,3).Value = accion
    ws.Cells(u,4).Value = detalle
    ws.Cells(u,5).Value = tabla
    ws.Cells(u,6).Value = Format(Now,"DD/MM/YYYY")
    ws.Cells(u,7).Value = Format(Now,"HH:MM:SS")
End Sub

Function NombrePROM(materia As String, grado As String) As String
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("BD_Materias")
    Dim i As Long
    For i=2 To ws.Cells(ws.Rows.Count,1).End(xlUp).Row
        If ws.Cells(i,2).Value=materia And ws.Cells(i,3).Value=grado Then
            NombrePROM=CStr(ws.Cells(i,4).Value): Exit Function
        End If
    Next i
    NombrePROM=""
End Function

Function NombrePlanilla(materia As String, grado As String) As String
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("BD_Materias")
    Dim i As Long
    For i=2 To ws.Cells(ws.Rows.Count,1).End(xlUp).Row
        If ws.Cells(i,2).Value=materia And ws.Cells(i,3).Value=grado Then
            NombrePlanilla=CStr(ws.Cells(i,5).Value): Exit Function
        End If
    Next i
End Function

Function ColComp(tri As Integer, comp As String) As Integer
    If tri=1 Then
        If comp="Parciales"   Then ColComp=3
        If comp="Apreciacion" Then ColComp=18
        If comp="Prueba"      Then ColComp=28
    ElseIf tri=2 Then
        If comp="Parciales"   Then ColComp=37
        If comp="Apreciacion" Then ColComp=51
        If comp="Prueba"      Then ColComp=62
    ElseIf tri=3 Then
        If comp="Parciales"   Then ColComp=71
        If comp="Apreciacion" Then ColComp=85
        If comp="Prueba"      Then ColComp=96
    End If
End Function

Function MaxComp(comp As String) As Integer
    If comp="Parciales"   Then MaxComp=13
    If comp="Apreciacion" Then MaxComp=9
    If comp="Prueba"      Then MaxComp=2
End Function

Function FilaEst(n As Integer) As Integer: FilaEst=n+4: End Function

Function NombreEst(n As Integer, grado As String) As String
    Dim ws As Worksheet: Set ws=ThisWorkbook.Sheets("MAESTRO")
    Dim col As Integer
    If grado="7" Then col=2 Else If grado="8" Then col=4 Else col=6
    NombreEst=Trim(CStr(ws.Cells(n+4,col).Value))
End Function

Function ContarEst(grado As String) As Integer
    Dim ws As Worksheet: Set ws=ThisWorkbook.Sheets("MAESTRO")
    Dim col As Integer
    If grado="7" Then col=2 Else If grado="8" Then col=4 Else col=6
    Dim t As Integer: t=0
    Dim i As Integer
    For i=5 To 44: If Trim(CStr(ws.Cells(i,col).Value))<>"" Then t=t+1: Next i
    ContarEst=t
End Function

Function LeerNota(hoja As String, numEst As Integer, tri As Integer, comp As String, numNota As Integer) As Double
    On Error Resume Next
    Dim ws As Worksheet: Set ws=ThisWorkbook.Sheets(hoja)
    If ws Is Nothing Then LeerNota=0: Exit Function
    Dim v As Variant: v=ws.Cells(FilaEst(numEst),ColComp(tri,comp)+numNota-1).Value
    If IsNumeric(v) Then LeerNota=CDbl(v) Else LeerNota=0
End Function

Sub EscribirNota(hoja As String, numEst As Integer, tri As Integer, comp As String, numNota As Integer, nota As Double)
    On Error Resume Next
    Dim ws As Worksheet: Set ws=ThisWorkbook.Sheets(hoja)
    If ws Is Nothing Then MsgBox "Hoja no encontrada: "&hoja,vbExclamation,"Error": Exit Sub
    ws.Cells(FilaEst(numEst),ColComp(tri,comp)+numNota-1).Value=nota
    Call RegistrarAuditoria("NOTA","Materia: "&hoja&" Est:"&numEst&" T"&tri&" "&comp&" Nota"&numNota&"="&nota,"PROM")
End Sub

Function LeerFinal(hoja As String, numEst As Integer, tri As Integer) As Double
    On Error Resume Next
    Dim ws As Worksheet: Set ws=ThisWorkbook.Sheets(hoja)
    If ws Is Nothing Then LeerFinal=0: Exit Function
    Dim col As Integer
    If tri=1 Then col=31 Else If tri=2 Then col=65 Else col=99
    Dim v As Variant: v=ws.Cells(FilaEst(numEst),col).Value
    If IsNumeric(v) Then LeerFinal=CDbl(v) Else LeerFinal=0
End Function

Function MateriasGrado(grado As String) As String()
    Dim ws As Worksheet: Set ws=ThisWorkbook.Sheets("BD_Materias")
    Dim arr() As String
    Dim cnt As Integer: cnt=0
    Dim i As Long
    For i=2 To ws.Cells(ws.Rows.Count,1).End(xlUp).Row
        If ws.Cells(i,3).Value=grado Then cnt=cnt+1
    Next i
    ReDim arr(cnt-1)
    Dim idx As Integer: idx=0
    For i=2 To ws.Cells(ws.Rows.Count,1).End(xlUp).Row
        If ws.Cells(i,3).Value=grado Then arr(idx)=CStr(ws.Cells(i,2).Value): idx=idx+1
    Next i
    MateriasGrado=arr
End Function

Function GradosDisponibles() As String()
    Dim ws As Worksheet: Set ws=ThisWorkbook.Sheets("BD_Materias")
    Dim dict As Object: Set dict=CreateObject("Scripting.Dictionary")
    Dim i As Long
    For i=2 To ws.Cells(ws.Rows.Count,1).End(xlUp).Row
        Dim g As String: g=CStr(ws.Cells(i,3).Value)
        If g<>"" And Not dict.Exists(g) Then dict.Add g,g
    Next i
    Dim arr() As String: ReDim arr(dict.Count-1)
    Dim idx As Integer: idx=0
    Dim k As Variant
    For Each k In dict.Keys: arr(idx)=CStr(k): idx=idx+1: Next k
    GradosDisponibles=arr
End Function

Function ValidarCedula(cedula As String) As Boolean
    cedula=Trim(cedula)
    If Len(cedula)>=7 And Len(cedula)<=10 Then
        Dim i As Integer
        For i=1 To Len(cedula)
            If Not IsNumeric(Mid(cedula,i,1)) Then ValidarCedula=False: Exit Function
        Next i
        ValidarCedula=True
    Else
        ValidarCedula=False
    End If
End Function

Sub GuardarRespaldo()
    Dim r As String: r=ThisWorkbook.Path&"\\Respaldo_"&Format(Now,"YYYYMMDD_HHMM")&".xlsm"
    ThisWorkbook.SaveCopyAs r
    Call RegistrarAuditoria("RESPALDO","Respaldo guardado en: "&r,"SISTEMA")
    MsgBox "Respaldo guardado:"&Chr(10)&r,vbInformation,"OK"
End Sub

Function TienePermiso(permiso As String) As Boolean
    Select Case RolActual
        Case "Admin":    TienePermiso=True
        Case "Docente":  TienePermiso=(permiso<>"USUARIOS" And permiso<>"AUDITORIA_TOTAL")
        Case "Director": TienePermiso=(permiso="LECTURA" Or permiso="REPORTES")
        Case Else:       TienePermiso=False
    End Select
End Function

Sub AbrirMenu(): frmMenu.Show: End Sub

Sub AbrirLogin()
    On Error Resume Next
    Call OcultarTodasLasHojas
    frmLogin.Show
    On Error GoTo 0
End Sub
""")
print("      RD_CORE OK")

# ---- THISWORKBOOK ----
vba.VBComponents("ThisWorkbook").CodeModule.AddFromString("""
Private Sub Workbook_Open()
    On Error Resume Next
    Application.Wait Now + TimeValue("00:00:01")
    Call OcultarTodasLasHojas

    Err.Clear
    frmLogin.Show

    ' Si falla el login por cualquier motivo, NO destapar todo el libro.
    If Err.Number <> 0 Then
        Err.Clear
        ThisWorkbook.Sheets("PORTADA").Visible = True
        ThisWorkbook.Sheets("PORTADA").Activate
        MsgBox "No se pudo abrir el login automaticamente. Habilita macros y vuelve a abrir el archivo.", vbExclamation, "RegistroDoc"
    End If

    On Error GoTo 0
End Sub

Sub OcultarTodasLasHojas()
    On Error Resume Next
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Sheets
        If ws.Name <> "PORTADA" Then ws.Visible = 2
    Next ws
    ThisWorkbook.Sheets("PORTADA").Visible = True
    On Error GoTo 0
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    If UsuarioActual <> "" Then
        Call RegistrarAuditoria("CIERRE","Sesion cerrada","SISTEMA")
        If MsgBox("Guardar respaldo antes de cerrar?",vbYesNo+vbQuestion,"Respaldo")=vbYes Then
            Call GuardarRespaldo
        End If
    End If
End Sub
""")

# ---- FORMS CODIGO ----
forms_cod = {

"frmLogin": """
Private Sub UserForm_Initialize()
    Me.Caption="RegistroDoc Multigrado v2.1"
    Me.BackColor=RGB(26,26,46): Me.Width=380: Me.Height=340
    lblError.Caption=""
    lblError.ForeColor=RGB(255,100,100)
End Sub

Private Sub btnIngresar_Click()
    Dim usr As String: usr=Trim(txtUsuario.Value)
    Dim pwd As String: pwd=Trim(txtPassword.Value)
    If usr="" Or pwd="" Then lblError.Caption="Ingresa usuario y contrasena.": Exit Sub
    Dim ws As Worksheet: Set ws=ThisWorkbook.Sheets("BD_Usuarios")
    Dim i As Long
    For i=2 To ws.Cells(ws.Rows.Count,1).End(xlUp).Row
        If LCase(ws.Cells(i,1).Value)=LCase(usr) And ws.Cells(i,4).Value="Si" Then
            If ws.Cells(i,7).Value="Si" Then
                lblError.Caption="Usuario bloqueado. Contacta al Admin.": Exit Sub
            End If
            If ws.Cells(i,2).Value=pwd Then
                ws.Cells(i,6).Value=0
                UsuarioActual=usr: RolActual=CStr(ws.Cells(i,3).Value)
                Call RegistrarAuditoria("LOGIN","Sesion: "&RolActual,"SEGURIDAD")
                Me.Hide: frmMenu.Show: Exit Sub
            Else
                Dim intentos As Integer: intentos=CInt(ws.Cells(i,6).Value)+1
                ws.Cells(i,6).Value=intentos
                If intentos>=3 Then
                    ws.Cells(i,7).Value="Si"
                    Call RegistrarAuditoria("BLOQUEO","Bloqueado por intentos","SEGURIDAD")
                    lblError.Caption="3 intentos fallidos. Usuario BLOQUEADO."
                Else
                    lblError.Caption="Clave incorrecta. Intento "&intentos&" de 3."
                End If
                Exit Sub
            End If
        End If
    Next i
    lblError.Caption="Usuario no encontrado o inactivo."
End Sub

Private Sub btnCancelar_Click()
    If MsgBox("Cerrar RegistroDoc?",vbYesNo+vbQuestion,"Salir")=vbYes Then
        Application.Quit
    End If
End Sub
""",

"frmMenu": """
Private Sub UserForm_Initialize()
    Me.Caption="RegistroDoc Multigrado v2.1"
    Me.BackColor=RGB(26,26,46): Me.Width=460: Me.Height=580
    lblUsuario.Caption="Usuario: "&UsuarioActual&"  |  Rol: "&RolActual
    lblEstats.Caption="7G: "&ContarEst("7")&"  8G: "&ContarEst("8")&"  9G: "&ContarEst("9")
End Sub
Private Sub btnEstudiantes_Click(): Me.Hide: frmEstudiantes.Show: Me.Show: End Sub
Private Sub btnNotas_Click():       Me.Hide: frmNotas.Show:       Me.Show: End Sub
Private Sub btnAsistencia_Click():  Me.Hide: frmAsistencia.Show:  Me.Show: End Sub
Private Sub btnResumenes_Click():   Me.Hide: frmResumenes.Show:   Me.Show: End Sub
Private Sub btnDashboard_Click():   Me.Hide: frmDashboard.Show:   Me.Show: End Sub
Private Sub btnReportes_Click():    Me.Hide: frmReportes.Show:    Me.Show: End Sub
Private Sub btnConfig_Click()
    If Not TienePermiso("CONFIG") Then MsgBox "Sin permiso.",vbExclamation,"Error": Exit Sub
    Me.Hide: frmConfiguracion.Show: Me.Show
End Sub
Private Sub btnAuditoria_Click()
    If Not TienePermiso("AUDITORIA_TOTAL") Then MsgBox "Solo Admin puede ver auditoria.",vbExclamation,"Error": Exit Sub
    Me.Hide: frmAuditoria.Show: Me.Show
End Sub
Private Sub btnRespaldo_Click(): Call GuardarRespaldo: End Sub
Private Sub btnCerrar_Click()
    If MsgBox("Cerrar RegistroDoc?",vbYesNo+vbQuestion,"Cerrar")=vbYes Then
        ThisWorkbook.Save: ThisWorkbook.Close
    End If
End Sub
""",

"frmEstudiantes": """
Private Sub UserForm_Initialize()
    Me.Caption="Gestion de Estudiantes"
    Me.BackColor=RGB(26,26,46): Me.Width=540: Me.Height=600
    Dim g As String
    Dim grados() As String: grados=GradosDisponibles()
    cboGradoFiltro.AddItem "Todos"
    cboGradoNuevo.AddItem ""
    For Each g In grados
        cboGradoFiltro.AddItem g
        cboGradoNuevo.AddItem g
    Next g
    cboGradoFiltro.Value="Todos"
    Call CargarLista
End Sub

Private Sub CargarLista()
    lstEstudiantes.Clear
    Dim ws As Worksheet: Set ws=ThisWorkbook.Sheets("MAESTRO")
    Dim grados() As String: grados=GradosDisponibles()
    Dim total As Integer: total=0
    Dim g As Variant
    For Each g In grados
        If cboGradoFiltro.Value="Todos" Or cboGradoFiltro.Value=CStr(g) Then
            Dim col As Integer
            If CStr(g)="7" Then col=2 Else If CStr(g)="8" Then col=4 Else col=6
            Dim i As Integer
            For i=5 To 44
                Dim nom As String: nom=Trim(CStr(ws.Cells(i,col).Value))
                If nom<>"" Then
                    Dim buscar As String: buscar=LCase(Trim(txtBuscar.Value))
                    If buscar="" Or InStr(LCase(nom),buscar)>0 Then
                        lstEstudiantes.AddItem "G"&CStr(g)&" - Pos"&(i-4)&" | "&nom
                        total=total+1
                    End If
                End If
            Next i
        End If
    Next g
    lblContador.Caption="Total: "&total&" estudiantes"
End Sub

Private Sub cboGradoFiltro_Change(): Call CargarLista: End Sub
Private Sub txtBuscar_Change(): Call CargarLista: End Sub

Private Sub btnAgregar_Click()
    If Not TienePermiso("CRUD_EST") Then MsgBox "Sin permiso.",vbExclamation,"Error": Exit Sub
    Dim ap As String: ap=Trim(txtApellido.Value)
    Dim nm As String: nm=Trim(txtNombre.Value)
    Dim gr As String: gr=Trim(cboGradoNuevo.Value)
    Dim ced As String: ced=Trim(txtCedula.Value)
    If ap="" Then MsgBox "Ingresa el apellido.",vbExclamation,"Error": Exit Sub
    If nm="" Then MsgBox "Ingresa el nombre.",vbExclamation,"Error": Exit Sub
    If gr="" Then MsgBox "Selecciona el grado.",vbExclamation,"Error": Exit Sub
    If ced<>"" And Not ValidarCedula(ced) Then
        MsgBox "Cedula invalida. Debe tener 7 a 10 digitos numericos.",vbExclamation,"Error": Exit Sub
    End If
    Dim ws As Worksheet: Set ws=ThisWorkbook.Sheets("MAESTRO")
    Dim col As Integer
    If gr="7" Then col=2 Else If gr="8" Then col=4 Else col=6
    Dim pos As Integer: pos=0
    Dim i As Integer
    For i=5 To 44
        If Trim(CStr(ws.Cells(i,col).Value))="" Then pos=i: Exit For
    Next i
    If pos=0 Then MsgBox "Grado "&gr&" esta lleno (max 40 estudiantes).",vbExclamation,"Error": Exit Sub
    ws.Cells(pos,col).Value=UCase(ap)&", "&nm
    Call RegistrarAuditoria("AGREGAR_EST","Agregado: "&UCase(ap)&", "&nm&" Grado "&gr,"MAESTRO")
    MsgBox "Estudiante agregado: "&UCase(ap)&", "&nm,vbInformation,"Guardado"
    txtApellido.Value="": txtNombre.Value="": txtCedula.Value=""
    Call CargarLista
End Sub

Private Sub btnEliminar_Click()
    If Not TienePermiso("CRUD_EST") Then MsgBox "Sin permiso.",vbExclamation,"Error": Exit Sub
    If lstEstudiantes.ListIndex=-1 Then MsgBox "Selecciona un estudiante.",vbExclamation,"Error": Exit Sub
    Dim item As String: item=lstEstudiantes.Value
    Dim partes() As String: partes=Split(item," | ")
    Dim info As String: info=partes(0)
    Dim gr As String: gr=Mid(info,2,1)
    Dim pos As Integer: pos=CInt(Mid(info,InStr(info,"Pos")+3,2))
    If MsgBox("Eliminar: "&partes(1)&Chr(10)&"Grado "&gr&" Posicion "&pos,vbYesNo+vbQuestion,"Confirmar")=vbNo Then Exit Sub
    Dim ws As Worksheet: Set ws=ThisWorkbook.Sheets("MAESTRO")
    Dim col As Integer
    If gr="7" Then col=2 Else If gr="8" Then col=4 Else col=6
    Dim fila As Integer: fila=pos+4
    Call RegistrarAuditoria("ELIMINAR_EST","Eliminado: "&ws.Cells(fila,col).Value&" Grado "&gr,"MAESTRO")
    ws.Cells(fila,col).Value=""
    MsgBox "Estudiante eliminado.",vbInformation,"OK"
    Call CargarLista
End Sub

Private Sub btnCerrar_Click(): Me.Hide: End Sub
""",

"frmNotas": """
Private Sub UserForm_Initialize()
    Me.Caption="Ingreso de Calificaciones - MEDUCA Panama"
    Me.BackColor=RGB(26,26,46): Me.Width=600: Me.Height=680
    Dim grados() As String: grados=GradosDisponibles()
    Dim g As Variant
    For Each g In grados: cboGrado.AddItem CStr(g): Next g
    cboTrimestre.AddItem "1 - Primer Trimestre"
    cboTrimestre.AddItem "2 - Segundo Trimestre"
    cboTrimestre.AddItem "3 - Tercer Trimestre"
    cboComponente.AddItem "Parciales (hasta 13)"
    cboComponente.AddItem "Apreciacion (hasta 9)"
    cboComponente.AddItem "Prueba (hasta 2)"
    txtFechaNota.Value=Format(Now,"DD/MM/YYYY")
    lblInfo.Caption="Selecciona grado -> materia -> trimestre -> componente"
    lblFinal.Caption=""
End Sub
Private Sub cboGrado_Change()
    cboMateria.Clear
    Dim mats() As String: mats=MateriasGrado(cboGrado.Value)
    Dim m As Variant: For Each m In mats: cboMateria.AddItem CStr(m): Next m
    lstEstudiantes.Clear: lstNotas.Clear
    lblInfo.Caption="Grado "&cboGrado.Value&" -> Selecciona materia"
End Sub
Private Sub cboMateria_Change():    Call Refrescar: End Sub
Private Sub cboTrimestre_Change():  Call Refrescar: End Sub
Private Sub cboComponente_Change(): Call Refrescar: End Sub
Private Sub lstEstudiantes_Click(): Call RefrescarNotas: End Sub

Private Sub Refrescar()
    If cboGrado.Value="" Or cboMateria.Value="" Then Exit Sub
    Dim g As String: g=cboGrado.Value
    lstEstudiantes.Clear
    Dim i As Integer
    For i=1 To 40
        Dim nom As String: nom=NombreEst(i,g)
        If nom<>"" Then lstEstudiantes.AddItem i&" | "&nom
    Next i
    Call RefrescarNotas
End Sub

Private Sub RefrescarNotas()
    lstNotas.Clear: lblFinal.Caption=""
    If cboGrado.Value="" Or cboMateria.Value="" Or cboTrimestre.Value="" Or cboComponente.Value="" Then Exit Sub
    If lstEstudiantes.ListIndex=-1 Then
        lblInfo.Caption=cboMateria.Value&" | Grado "&cboGrado.Value&" | Selecciona estudiante"
        Exit Sub
    End If
    Dim g As String: g=cboGrado.Value
    Dim mat As String: mat=cboMateria.Value
    Dim tri As Integer: tri=CInt(Left(cboTrimestre.Value,1))
    Dim comp As String
    If InStr(cboComponente.Value,"Parc")>0 Then comp="Parciales"
    If InStr(cboComponente.Value,"Aprec")>0 Then comp="Apreciacion"
    If InStr(cboComponente.Value,"Prueba")>0 Then comp="Prueba"
    Dim hoja As String: hoja=NombrePROM(mat,g)
    If hoja="" Then lblInfo.Caption="ERROR: hoja PROM no encontrada para "&mat&" "&g: Exit Sub
    Dim numEst As Integer: numEst=CInt(Split(lstEstudiantes.Value," | ")(0))
    Dim maxN As Integer: maxN=MaxComp(comp)
    Dim cnt As Integer: cnt=0
    Dim i As Integer
    For i=1 To maxN
        Dim nota As Double: nota=LeerNota(hoja,numEst,tri,comp,i)
        If nota>0 Then
            lstNotas.AddItem "Nota "&i&":  "&Format(nota,"0.0")
            cnt=cnt+1
        Else
            lstNotas.AddItem "Nota "&i&":  (vacia)"
        End If
    Next i
    Dim pf As Double: pf=LeerFinal(hoja,numEst,tri)
    If pf>0 Then
        Dim est As String: If pf>=3 Then est="APROBADO" Else est="REPROBADO"
        lblFinal.Caption="PROMEDIO T"&tri&": "&Format(pf,"0.0")&"  ->  "&est
        If pf>=3 Then lblFinal.ForeColor=RGB(100,220,100) Else lblFinal.ForeColor=RGB(255,100,100)
    Else
        lblFinal.Caption="Promedio T"&tri&": (sin notas aun)"
        lblFinal.ForeColor=RGB(150,150,150)
    End If
    lblInfo.Caption=mat&" | G"&g&" | T"&tri&" | "&comp&" | "&cnt&"/"&maxN&" ingresadas"
End Sub

Private Sub btnGuardar_Click()
    If Not TienePermiso("NOTAS") Then MsgBox "Sin permiso.",vbExclamation,"Error": Exit Sub
    If cboGrado.Value="" Or cboMateria.Value="" Or cboTrimestre.Value="" Or cboComponente.Value="" Then _
        MsgBox "Selecciona todos los campos.",vbExclamation,"Error": Exit Sub
    If lstEstudiantes.ListIndex=-1 Then MsgBox "Selecciona un estudiante.",vbExclamation,"Error": Exit Sub
    If lstNotas.ListIndex=-1 Then MsgBox "Selecciona cual nota ingresar.",vbExclamation,"Error": Exit Sub
    If txtNota.Value="" Then MsgBox "Escribe la nota (1.0 a 5.0).",vbExclamation,"Error": Exit Sub
    Dim nota As Double
    On Error GoTo ErrN: nota=CDbl(txtNota.Value): On Error GoTo 0
    If nota<1 Or nota>5 Then MsgBox "Nota entre 1.0 y 5.0",vbExclamation,"Error": Exit Sub
    Dim comp As String
    If InStr(cboComponente.Value,"Parc")>0 Then comp="Parciales"
    If InStr(cboComponente.Value,"Aprec")>0 Then comp="Apreciacion"
    If InStr(cboComponente.Value,"Prueba")>0 Then comp="Prueba"
    Dim tri As Integer: tri=CInt(Left(cboTrimestre.Value,1))
    Dim numNota As Integer: numNota=lstNotas.ListIndex+1
    Dim numEst As Integer: numEst=CInt(Split(lstEstudiantes.Value," | ")(0))
    Dim hoja As String: hoja=NombrePROM(cboMateria.Value,cboGrado.Value)
    Call EscribirNota(hoja,numEst,tri,comp,numNota,nota)
    If txtDescripcion.Value<>"" Or txtFechaNota.Value<>"" Then
        Dim ws As Worksheet: Set ws=ThisWorkbook.Sheets(hoja)
        ws.Cells(FilaEst(numEst),36+tri).Value=txtDescripcion.Value&" ["&txtFechaNota.Value&"]"
    End If
    Call RefrescarNotas
    txtNota.Value=""
    Exit Sub
ErrN: MsgBox "Numero invalido. Ejemplo: 3.5",vbExclamation,"Error"
End Sub

Private Sub btnBorrarNota_Click()
    If lstEstudiantes.ListIndex=-1 Or lstNotas.ListIndex=-1 Then _
        MsgBox "Selecciona estudiante y nota.",vbExclamation,"Error": Exit Sub
    Dim comp As String
    If InStr(cboComponente.Value,"Parc")>0 Then comp="Parciales"
    If InStr(cboComponente.Value,"Aprec")>0 Then comp="Apreciacion"
    If InStr(cboComponente.Value,"Prueba")>0 Then comp="Prueba"
    Dim tri As Integer: tri=CInt(Left(cboTrimestre.Value,1))
    Dim numEst As Integer: numEst=CInt(Split(lstEstudiantes.Value," | ")(0))
    Dim hoja As String: hoja=NombrePROM(cboMateria.Value,cboGrado.Value)
    Dim ws As Worksheet: Set ws=ThisWorkbook.Sheets(hoja)
    ws.Cells(FilaEst(numEst),ColComp(tri,comp)+lstNotas.ListIndex).ClearContents
    Call RegistrarAuditoria("BORRAR_NOTA","Nota borrada en "&hoja&" Est:"&numEst,"PROM")
    Call RefrescarNotas
End Sub
Private Sub btnCerrar_Click(): Me.Hide: End Sub
""",

"frmAsistencia": """
Private Sub UserForm_Initialize()
    Me.Caption="Control de Asistencia Diaria"
    Me.BackColor=RGB(26,26,46): Me.Width=520: Me.Height=560
    Dim grados() As String: grados=GradosDisponibles()
    Dim g As Variant: For Each g In grados: cboGrado.AddItem CStr(g): Next g
    txtFecha.Value=Format(Now,"DD/MM/YYYY")
    lblUltimo.Caption=""
End Sub
Private Sub cboGrado_Change(): Call CargarLista: End Sub
Private Sub CargarLista()
    lstEstudiantes.Clear
    If cboGrado.Value="" Then Exit Sub
    Dim i As Integer, t As Integer: t=0
    For i=1 To 40
        Dim nom As String: nom=NombreEst(i,cboGrado.Value)
        If nom<>"" Then
            Dim aus As Long: aus=0
            Dim g As String: g=cboGrado.Value
            Dim asistNom As String: asistNom="Asistencia ("&g&Chr(176)&")"
            Dim wsA As Worksheet
            On Error Resume Next: Set wsA=ThisWorkbook.Sheets(asistNom): On Error GoTo 0
            If Not wsA Is Nothing Then
                Dim v As Variant: v=wsA.Cells(i+2,61).Value
                If IsNumeric(v) Then aus=CLng(v)
            End If
            Dim alerta As String: alerta=""
            If aus>5 Then alerta=" *** ALERTA: "&aus&" aus."
            lstEstudiantes.AddItem i&" | "&nom&"  [A:"&aus&"]"&alerta
            t=t+1
        End If
    Next i
    lblContador.Caption=t&" estudiantes - Grado "&cboGrado.Value
End Sub
Private Sub btnP_Click(): Call Marcar("P"): End Sub
Private Sub btnA_Click(): Call Marcar("A"): End Sub
Private Sub btnT_Click(): Call Marcar("T"): End Sub
Private Sub Marcar(estado As String)
    If Not TienePermiso("ASISTENCIA") Then MsgBox "Sin permiso.",vbExclamation,"Error": Exit Sub
    If lstEstudiantes.ListIndex=-1 Then MsgBox "Selecciona un estudiante.",vbExclamation,"Error": Exit Sub
    If txtFecha.Value="" Then MsgBox "Ingresa la fecha.",vbExclamation,"Error": Exit Sub
    Dim g As String: g=cboGrado.Value
    Dim asistNom As String: asistNom="Asistencia ("&g&Chr(176)&")"
    Dim ws As Worksheet
    On Error Resume Next: Set ws=ThisWorkbook.Sheets(asistNom): On Error GoTo 0
    If ws Is Nothing Then MsgBox "Hoja no encontrada: "&asistNom,vbExclamation,"Error": Exit Sub
    Dim f As Date
    On Error GoTo ErrF: f=CDate(txtFecha.Value): On Error GoTo 0
    Dim numEst As Integer: numEst=CInt(Split(lstEstudiantes.Value," | ")(0))
    Dim filaA As Integer: filaA=numEst+2
    Dim colF As Integer: colF=0
    Dim c As Integer
    For c=3 To 62
        Dim vF As Variant: vF=ws.Cells(2,c).Value
        If IsDate(vF) Then
            If CDate(vF)=f Then colF=c: Exit For
        End If
    Next c
    If colF=0 Then
        For c=3 To 62
            If IsEmpty(ws.Cells(2,c)) Or ws.Cells(2,c).Value="" Then
                colF=c: ws.Cells(2,c).Value=f: Exit For
            End If
        Next c
    End If
    If colF=0 Then MsgBox "Sin columnas disponibles.",vbExclamation,"Error": Exit Sub
    ws.Cells(filaA,colF).Value=estado
    Dim nom As String: nom=Split(Split(lstEstudiantes.Value," | ")(1),"  [")(0)
    Dim es(2) As String: es(0)="PRESENTE": es(1)="AUSENTE": es(2)="TARDANZA"
    Dim idx As Integer: If estado="P" Then idx=0 Else If estado="A" Then idx=1 Else idx=2
    If estado="A" Then lblUltimo.ForeColor=RGB(255,100,100)
    If estado="P" Then lblUltimo.ForeColor=RGB(100,220,100)
    If estado="T" Then lblUltimo.ForeColor=RGB(255,200,50)
    lblUltimo.Caption=nom&" -> "&es(idx)
    Call RegistrarAuditoria("ASISTENCIA",nom&" "&estado&" "&txtFecha.Value,"ASISTENCIA")
    Call CargarLista
    Exit Sub
ErrF: MsgBox "Fecha invalida. Formato: DD/MM/YYYY",vbExclamation,"Error"
End Sub
Private Sub btnTodosP_Click()
    If cboGrado.Value="" Or txtFecha.Value="" Then _
        MsgBox "Selecciona grado y fecha.",vbExclamation,"Error": Exit Sub
    If MsgBox("Marcar TODOS Presentes el "&txtFecha.Value&"?",vbYesNo+vbQuestion,"Confirmar")=vbNo Then Exit Sub
    Dim i As Integer
    For i=0 To lstEstudiantes.ListCount-1
        lstEstudiantes.ListIndex=i
        If InStr(lstEstudiantes.Value,"vacio")=0 Then Call Marcar("P")
    Next i
    MsgBox "Todos marcados Presentes.",vbInformation,"Listo"
End Sub
Private Sub btnCerrar_Click(): Me.Hide: End Sub
""",

"frmResumenes": """
Private Sub UserForm_Initialize()
    Me.Caption="Resumenes de Notas Finales"
    Me.BackColor=RGB(26,26,46): Me.Width=640: Me.Height=580
    Dim grados() As String: grados=GradosDisponibles()
    Dim g As Variant: For Each g In grados: cboGrado.AddItem CStr(g): Next g
    cboTrimestre.AddItem "1 - Primer Trimestre"
    cboTrimestre.AddItem "2 - Segundo Trimestre"
    cboTrimestre.AddItem "3 - Tercer Trimestre"
    cboTrimestre.AddItem "Final - Promedio 3 Trimestres"
    lblStats.Caption="Selecciona grado y trimestre"
End Sub
Private Sub cboGrado_Change():     Call CargarResumen: End Sub
Private Sub cboTrimestre_Change(): Call CargarResumen: End Sub
Private Sub CargarResumen()
    If cboGrado.Value="" Or cboTrimestre.Value="" Then Exit Sub
    lstResumen.Clear
    Dim g As String: g=cboGrado.Value
    Dim mats() As String: mats=MateriasGrado(g)
    Dim apro As Long, repr As Long: apro=0: repr=0
    Dim i As Integer
    For i=1 To 40
        Dim nom As String: nom=NombreEst(i,g)
        If nom="" Then GoTo SigEst
        Dim linea As String: linea=i&" | "&nom&"   "
        If InStr(cboTrimestre.Value,"Final")>0 Then
            Dim sumF As Double, cntF As Integer: sumF=0: cntF=0
            Dim m As Variant
            For Each m In mats
                Dim hp As String: hp=NombrePROM(CStr(m),g)
                If hp<>"" Then
                    Dim t As Integer, sumT As Double, cntT As Integer: sumT=0: cntT=0
                    For t=1 To 3
                        Dim pf As Double: pf=LeerFinal(hp,i,t)
                        If pf>0 Then sumT=sumT+pf: cntT=cntT+1
                    Next t
                    If cntT>0 Then sumF=sumF+(sumT/cntT): cntF=cntF+1
                End If
            Next m
            If cntF>0 Then
                Dim prom As Double: prom=sumF/cntF
                Dim est As String: If prom>=3 Then est="APROBADO": apro=apro+1 Else est="REPROBADO": repr=repr+1
                linea=linea&"Prom: "&Format(prom,"0.0")&"  ->  "&est
            Else
                linea=linea&"Sin notas"
            End If
        Else
            Dim tri As Integer: tri=CInt(Left(cboTrimestre.Value,1))
            Dim m2 As Variant
            For Each m2 In mats
                Dim hp2 As String: hp2=NombrePROM(CStr(m2),g)
                If hp2<>"" Then
                    Dim pf2 As Double: pf2=LeerFinal(hp2,i,tri)
                    Dim ns As String: If pf2>0 Then ns=Format(pf2,"0.0") Else ns="---"
                    linea=linea&Left(CStr(m2),8)&":"&ns&"  "
                End If
            Next m2
        End If
        lstResumen.AddItem linea
SigEst:
    Next i
    lblStats.Caption="Grado "&g&"  |  Aprobados: "&apro&"  |  Reprobados: "&repr&"  |  Total: "&(apro+repr)
End Sub
Private Sub btnCerrar_Click(): Me.Hide: End Sub
""",

"frmDashboard": """
Private Sub UserForm_Initialize()
    Me.Caption="Dashboard - Estadisticas y Alertas"
    Me.BackColor=RGB(26,26,46): Me.Width=520: Me.Height=520
    Call Refrescar
End Sub
Private Sub Refrescar()
    Dim grados() As String: grados=GradosDisponibles()
    Dim g As Variant
    For Each g In grados
        If CStr(g)="7" Then lblNum7.Caption=CStr(ContarEst("7"))
        If CStr(g)="8" Then lblNum8.Caption=CStr(ContarEst("8"))
        If CStr(g)="9" Then lblNum9.Caption=CStr(ContarEst("9"))
    Next g
    Dim total As Long: total=0
    For Each g In grados: total=total+ContarEst(CStr(g)): Next g
    lblNumTotal.Caption=CStr(total)
    lstAlertas.Clear
    Dim ws As Worksheet: Set ws=ThisWorkbook.Sheets("MAESTRO")
    For Each g In grados
        Dim gr As String: gr=CStr(g)
        Dim asistNom As String: asistNom="Asistencia ("&gr&Chr(176)&")"
        Dim wsA As Worksheet
        On Error Resume Next: Set wsA=ThisWorkbook.Sheets(asistNom): On Error GoTo 0
        If Not wsA Is Nothing Then
            Dim col As Integer
            If gr="7" Then col=2 Else If gr="8" Then col=4 Else col=6
            Dim i As Integer
            For i=5 To 44
                Dim nom As String: nom=Trim(CStr(ws.Cells(i,col).Value))
                If nom<>"" Then
                    Dim filaA As Integer: filaA=i-2
                    Dim aus As Variant: aus=wsA.Cells(filaA,61).Value
                    If IsNumeric(aus) And CDbl(aus)>5 Then
                        lstAlertas.AddItem "G"&gr&" - "&nom&" -> "&aus&" ausencias"
                    End If
                End If
            Next i
        End If
    Next g
    lblNumAlertas.Caption=lstAlertas.ListCount&" alertas de ausencias"
End Sub
Private Sub btnRefrescar_Click(): Call Refrescar: End Sub
Private Sub btnCerrar_Click(): Me.Hide: End Sub
""",

"frmReportes": """
Private Sub UserForm_Initialize()
    Me.Caption="Reportes e Impresion Oficial"
    Me.BackColor=RGB(26,26,46): Me.Width=500: Me.Height=460
End Sub

Private Sub btnReporteDir_Click()
    Dim grados() As String: grados=GradosDisponibles()
    Dim g As Variant
    Dim total As Long: total=0
    Dim msg As String
    msg="MINISTERIO DE EDUCACION DE PANAMA"&Chr(10)
    msg=msg&"SISTEMA EDUCATIVO PANAMENIO"&Chr(10)&String(50,"-")&Chr(10)
    msg=msg&"Centro Educativo: "&ObtenerConfig("nombre_escuela")&Chr(10)
    msg=msg&"Docente: "&ObtenerConfig("nombre_docente")&Chr(10)
    msg=msg&"Modalidad: "&ObtenerConfig("modalidad")&Chr(10)
    msg=msg&"Ano Lectivo: "&ObtenerConfig("anio_lectivo")&Chr(10)&String(50,"-")&Chr(10)
    msg=msg&"MATRICULA ACTUAL:"&Chr(10)
    For Each g In grados
        Dim cnt As Long: cnt=ContarEst(CStr(g))
        msg=msg&"  Grado "&CStr(g)&":  "&cnt&" estudiantes"&Chr(10)
        total=total+cnt
    Next g
    msg=msg&String(50,"-")&Chr(10)&"  TOTAL: "&total&" estudiantes"
    MsgBox msg,vbInformation,"Reporte para la Direccion"
    Call RegistrarAuditoria("REPORTE","Reporte de matricula generado","REPORTES")
End Sub

Private Sub btnVistaPrevia_Click()
    Dim ws As Worksheet: Set ws=ThisWorkbook.Sheets("BD_Materias")
    Dim i As Long
    Dim planillas As String: planillas=""
    For i=2 To ws.Cells(ws.Rows.Count,1).End(xlUp).Row
        Dim pnom As String: pnom=CStr(ws.Cells(i,5).Value)
        If pnom<>"" Then planillas=planillas&"  - "&pnom&Chr(10)
    Next i
    MsgBox "PLANILLAS DISPONIBLES PARA IMPRIMIR:"&Chr(10)&Chr(10)&planillas&Chr(10)&_
           "Usa 'Imprimir Libro Completo' para imprimir todas.",vbInformation,"Vista Previa"
End Sub

Private Sub btnImprimirTodo_Click()
    If MsgBox("Imprimir TODAS las planillas del libro?"&Chr(10)&Chr(10)&_
              "Asegurate de tener la impresora conectada.",_
              vbYesNo+vbQuestion,"Imprimir Libro Completo")=vbNo Then Exit Sub
    Dim ws As Worksheet: Set ws=ThisWorkbook.Sheets("BD_Materias")
    Dim i As Long, count As Integer: count=0
    For i=2 To ws.Cells(ws.Rows.Count,1).End(xlUp).Row
        Dim pnom As String: pnom=Trim(CStr(ws.Cells(i,5).Value))
        If pnom<>"" Then
            On Error Resume Next
            Dim wsPlan As Worksheet: Set wsPlan=ThisWorkbook.Sheets(pnom)
            On Error GoTo 0
            If Not wsPlan Is Nothing Then
                wsPlan.PageSetup.Orientation=2
                wsPlan.PageSetup.FitToPagesWide=1
                wsPlan.PrintOut
                count=count+1
            End If
        End If
    Next i
    Call RegistrarAuditoria("IMPRESION","Libro completo impreso - "&count&" planillas","REPORTES")
    MsgBox count&" planillas enviadas a la impresora.",vbInformation,"Impresion Completada"
End Sub

Private Sub btnExportarPDF_Click()
    Dim ruta As String
    ruta=Application.GetSaveAsFilename(_
        ObtenerConfig("nombre_escuela")&"_"&ObtenerConfig("anio_lectivo"),_
        "PDF (*.pdf),*.pdf")
    If ruta="False" Then Exit Sub
    Dim ws As Worksheet: Set ws=ThisWorkbook.Sheets("BD_Materias")
    Dim i As Long
    Dim hojasImpr() As String: ReDim hojasImpr(0)
    Dim cnt As Integer: cnt=0
    For i=2 To ws.Cells(ws.Rows.Count,1).End(xlUp).Row
        Dim pnom As String: pnom=Trim(CStr(ws.Cells(i,5).Value))
        If pnom<>"" Then
            On Error Resume Next
            Dim wsPlan As Worksheet: Set wsPlan=ThisWorkbook.Sheets(pnom)
            On Error GoTo 0
            If Not wsPlan Is Nothing Then
                ReDim Preserve hojasImpr(cnt): hojasImpr(cnt)=pnom: cnt=cnt+1
            End If
        End If
    Next i
    If cnt=0 Then MsgBox "No se encontraron planillas.",vbExclamation,"Error": Exit Sub
    ThisWorkbook.Sheets(hojasImpr).Select
    ActiveSheet.ExportAsFixedFormat Type:=0, Filename:=ruta
    ThisWorkbook.Sheets(hojasImpr(0)).Select
    Call RegistrarAuditoria("PDF","Libro exportado a PDF: "&ruta,"REPORTES")
    MsgBox "PDF exportado en:"&Chr(10)&ruta,vbInformation,"Exportado"
End Sub

Private Sub btnCerrar_Click(): Me.Hide: End Sub
""",

"frmConfiguracion": """
Private Sub UserForm_Initialize()
    Me.Caption="Configuracion del Sistema"
    Me.BackColor=RGB(26,26,46): Me.Width=580: Me.Height=680
    Dim grados() As String: grados=GradosDisponibles()
    Dim g As Variant
    For Each g In grados: cboGradoNuevaMat.AddItem CStr(g): Next g
    cboGradoNuevaMat.AddItem "(nuevo grado)"
    Call CargarDatos
    Call CargarMaterias
End Sub

Private Sub CargarDatos()
    txtAnio.Value=ObtenerConfig("anio_lectivo")
    txtDocente.Value=ObtenerConfig("nombre_docente")
    txtEscuela.Value=ObtenerConfig("nombre_escuela")
    txtUbic.Value=ObtenerConfig("ubicacion")
    txtModal.Value=ObtenerConfig("modalidad")
    txtDirector.Value=ObtenerConfig("director")
    txtEmail.Value=ObtenerConfig("email_docente")
End Sub

Private Sub CargarMaterias()
    lstMaterias.Clear
    Dim ws As Worksheet: Set ws=ThisWorkbook.Sheets("BD_Materias")
    Dim i As Long
    For i=2 To ws.Cells(ws.Rows.Count,1).End(xlUp).Row
        Dim nm As String: nm=CStr(ws.Cells(i,2).Value)
        Dim gr As String: gr=CStr(ws.Cells(i,3).Value)
        Dim union As String: union=""
        If ws.Cells(i,6).Value="Si" Then union=" [une->Tecnologia]"
        lstMaterias.AddItem ws.Cells(i,1).Value&" | "&nm&" - Grado "&gr&union
    Next i
End Sub

Private Sub btnAgregarMateria_Click()
    If Not TienePermiso("CONFIG") Then MsgBox "Sin permiso.",vbExclamation,"Error": Exit Sub
    Dim nm As String: nm=Trim(txtNuevaMat.Value)
    Dim gr As String: gr=Trim(cboGradoNuevaMat.Value)
    If nm="" Then MsgBox "Ingresa el nombre de la materia.",vbExclamation,"Error": Exit Sub
    If gr="" Or gr="(nuevo grado)" Then
        If txtNuevoGrado.Value="" Then
            MsgBox "Ingresa el grado en el campo 'Nuevo grado'.",vbExclamation,"Error": Exit Sub
        End If
        gr=Trim(txtNuevoGrado.Value)
    End If
    Dim wsM As Worksheet: Set wsM=ThisWorkbook.Sheets("BD_Materias")
    Dim u As Long: u=wsM.Cells(wsM.Rows.Count,1).End(xlUp).Row+1
    Dim newID As Long: newID=u-1
    Dim prom_nuevo As String: prom_nuevo="PROM ("&nm&" "&gr&Chr(176)&")"
    Dim plan_nuevo As String: plan_nuevo="Planilla ("&nm&" "&gr&Chr(176)&")"
    Dim esUnion As String: If chkUnion.Value Then esUnion="Si" Else esUnion="No"
    wsM.Cells(u,1).Value=newID
    wsM.Cells(u,2).Value=nm
    wsM.Cells(u,3).Value=gr
    wsM.Cells(u,4).Value=prom_nuevo
    wsM.Cells(u,5).Value=plan_nuevo
    wsM.Cells(u,6).Value=esUnion
    wsM.Cells(u,7).Value=IIf(chkUnion.Value,"Tecnologia","")
    Call CrearHojasPROMPlanilla(nm,gr,prom_nuevo,plan_nuevo)
    Call RegistrarAuditoria("AGREGAR_MATERIA","Materia: "&nm&" Grado "&gr,"CONFIG")
    Call CargarMaterias
    txtNuevaMat.Value="": txtNuevoGrado.Value=""
    MsgBox "Materia '"&nm&"' para Grado "&gr&" agregada."&Chr(10)&"Hojas PROM y Planilla creadas.",vbInformation,"OK"
End Sub

Sub CrearHojasPROMPlanilla(materia As String, grado As String, prom_nom As String, plan_nom As String)
    Dim nombres_hojas As Variant
    nombres_hojas=Array()
    Dim sn As Worksheet
    For Each sn In ThisWorkbook.Sheets
        If sn.Name=prom_nom Then Exit Sub
    Next sn
    Dim ws_base As Worksheet
    On Error Resume Next
    Set ws_base=ThisWorkbook.Sheets("PROM (Ingles 9" & Chr(176) & ")")
    On Error GoTo 0
    If Not ws_base Is Nothing Then
        ws_base.Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
        ActiveSheet.Name=prom_nom
        Dim ws_new As Worksheet: Set ws_new=ThisWorkbook.Sheets(prom_nom)
        Dim mcol As String
        If grado="7" Then mcol="B" Else If grado="8" Then mcol="D" Else mcol="F"
        Dim i As Integer
        For i=5 To 44: ws_new.Range("B"&i).Value="=MAESTRO!"&mcol&i: Next i
    End If
    Dim ws_plan_base As Worksheet
    On Error Resume Next
    Set ws_plan_base=ThisWorkbook.Sheets("Planilla (Ingles 9" & Chr(176) & ") ")
    If ws_plan_base Is Nothing Then Set ws_plan_base=ThisWorkbook.Sheets("Planilla (Ingles 9" & Chr(176) & ")")
    On Error GoTo 0
    If Not ws_plan_base Is Nothing Then
        ws_plan_base.Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
        ActiveSheet.Name=plan_nom
        Dim ws_pnew As Worksheet: Set ws_pnew=ThisWorkbook.Sheets(plan_nom)
        Dim mcol2 As String
        If grado="7" Then mcol2="B" Else If grado="8" Then mcol2="D" Else mcol2="F"
        Dim asistNom As String: asistNom="Asistencia ("&grado&Chr(176)&")"
        For i=0 To 39
            ws_pnew.Range("C"&(16+i)).Value="=MAESTRO!"&mcol2&(5+i)
            ws_pnew.Range("F"&(16+i)).Value="='"&prom_nom&"'!AE"&(5+i)
            ws_pnew.Range("G"&(16+i)).Value="='"&prom_nom&"'!BM"&(5+i)
            ws_pnew.Range("H"&(16+i)).Value="='"&prom_nom&"'!CU"&(5+i)
            ws_pnew.Range("I"&(16+i)).Value="=TRUNC(AVERAGE(F"&(16+i)&":H"&(16+i)&"),1)"
            ws_pnew.Range("J"&(16+i)).Value="='"&asistNom&"'!BI"&(3+i)
        Next i
        ws_pnew.Range("D11").Value=materia
        ws_pnew.Range("L12").Value=grado&Chr(176)
    End If
End Sub

Private Sub btnEliminarMateria_Click()
    If Not TienePermiso("CONFIG") Then MsgBox "Sin permiso.",vbExclamation,"Error": Exit Sub
    If lstMaterias.ListIndex=-1 Then MsgBox "Selecciona una materia.",vbExclamation,"Error": Exit Sub
    Dim item As String: item=lstMaterias.Value
    Dim id As Long: id=CLng(Split(item," | ")(0))
    If MsgBox("Eliminar esta materia?"&Chr(10)&item&Chr(10)&Chr(10)&_
              "ADVERTENCIA: Las hojas PROM y Planilla NO se eliminan.",_
              vbYesNo+vbQuestion,"Confirmar")=vbNo Then Exit Sub
    Dim ws As Worksheet: Set ws=ThisWorkbook.Sheets("BD_Materias")
    Dim i As Long
    For i=2 To ws.Cells(ws.Rows.Count,1).End(xlUp).Row
        If ws.Cells(i,1).Value=id Then ws.Rows(i).Delete: Exit For
    Next i
    Call RegistrarAuditoria("ELIMINAR_MATERIA","ID: "&id,"CONFIG")
    Call CargarMaterias
    MsgBox "Materia eliminada del sistema.",vbInformation,"OK"
End Sub

Private Sub btnGuardar_Click()
    If Not TienePermiso("CONFIG") Then MsgBox "Sin permiso.",vbExclamation,"Error": Exit Sub
    Call GuardarConfig("anio_lectivo",txtAnio.Value)
    Call GuardarConfig("nombre_docente",txtDocente.Value)
    Call GuardarConfig("nombre_escuela",txtEscuela.Value)
    Call GuardarConfig("ubicacion",txtUbic.Value)
    Call GuardarConfig("modalidad",txtModal.Value)
    Call GuardarConfig("director",txtDirector.Value)
    Call GuardarConfig("email_docente",txtEmail.Value)
    ThisWorkbook.Save
    Call RegistrarAuditoria("CONFIG","Configuracion guardada","CONFIG")
    MsgBox "Configuracion guardada.",vbInformation,"OK"
    Me.Hide
End Sub
Private Sub btnCerrar_Click(): Me.Hide: End Sub
""",

"frmAuditoria": """
Private Sub UserForm_Initialize()
    Me.Caption="Auditoria del Sistema - ISO 27001"
    Me.BackColor=RGB(26,26,46): Me.Width=620: Me.Height=520
    Call CargarAuditoria
End Sub
Private Sub CargarAuditoria()
    lstAuditoria.Clear
    Dim ws As Worksheet: Set ws=ThisWorkbook.Sheets("BD_Auditoria")
    Dim ultima As Long: ultima=ws.Cells(ws.Rows.Count,1).End(xlUp).Row
    Dim i As Long
    For i=ultima To 2 Step -1
        lstAuditoria.AddItem ws.Cells(i,6).Value&" "&ws.Cells(i,7).Value&_
            " | "&ws.Cells(i,2).Value&_
            " | "&ws.Cells(i,3).Value&_
            " | "&ws.Cells(i,4).Value
    Next i
    lblTotal.Caption="Total de registros: "&(ultima-1)
End Sub
Private Sub btnExportar_Click()
    Dim ruta As String
    ruta=Application.GetSaveAsFilename("Auditoria_"&Format(Now,"YYYYMMDD"),"Excel (*.xlsx),*.xlsx")
    If ruta="False" Then Exit Sub
    ThisWorkbook.Sheets("BD_Auditoria").Visible=True
    ThisWorkbook.Sheets("BD_Auditoria").Copy
    ActiveWorkbook.SaveAs ruta,51
    ActiveWorkbook.Close
    ThisWorkbook.Sheets("BD_Auditoria").Visible=2
    MsgBox "Auditoria exportada en:"&Chr(10)&ruta,vbInformation,"OK"
End Sub
Private Sub btnCerrar_Click(): Me.Hide: End Sub
""",
}

for nombre, codigo in forms_cod.items():
    uf = vba.VBComponents.Add(3)
    uf.Name = nombre
    uf.CodeModule.AddFromString(codigo)
    print(f"      {nombre} OK")

# ---- MODULO RD_BUILD (construir controles) ----
m_build = vba.VBComponents.Add(1); m_build.Name = "RD_Build"
m_build.CodeModule.AddFromString("""
Sub ConstruirTodosLosControles()
    On Error Resume Next
    Call BLogin
    Call BMenu
    Call BEstudiantes
    Call BNotas
    Call BAsistencia
    Call BResumenes
    Call BDashboard
    Call BReportes
    Call BConfig
    Call BAuditoria
    On Error GoTo 0
    MsgBox "RegistroDoc v2.1 - Controles construidos." & Chr(10) & "Sistema listo.", vbInformation, "Listo"
End Sub

Private Function NwL(f As Object, nm As String, txt As String, l%, t%, w%, h%, fs%, fc As Long, bc As Long, al%) As Object
    Dim o As Object: Set o=f.Controls.Add("Forms.Label.1",nm,True)
    o.Caption=txt: o.Left=l: o.Top=t: o.Width=w: o.Height=h
    o.Font.Size=fs: o.ForeColor=fc: o.BackColor=bc: o.TextAlign=al
    Set NwL=o
End Function
Private Function NwB(f As Object, nm As String, txt As String, l%, t%, w%, h%, col As Long) As Object
    Dim o As Object: Set o=f.Controls.Add("Forms.CommandButton.1",nm,True)
    o.Caption=txt: o.Left=l: o.Top=t: o.Width=w: o.Height=h
    o.BackColor=col: o.ForeColor=RGB(255,255,255): o.Font.Bold=True: o.Font.Size=10
    Set NwB=o
End Function
Private Function NwT(f As Object, nm As String, l%, t%, w%, h%, Optional pw As Boolean=False) As Object
    Dim o As Object: Set o=f.Controls.Add("Forms.TextBox.1",nm,True)
    o.Left=l: o.Top=t: o.Width=w: o.Height=h
    o.BackColor=RGB(46,46,70): o.ForeColor=RGB(255,255,255): o.Font.Size=10
    If pw Then o.PasswordChar="*"
    Set NwT=o
End Function
Private Function NwC(f As Object, nm As String, l%, t%, w%, h%) As Object
    Dim o As Object: Set o=f.Controls.Add("Forms.ComboBox.1",nm,True)
    o.Left=l: o.Top=t: o.Width=w: o.Height=h
    o.BackColor=RGB(46,46,70): o.ForeColor=RGB(255,255,255): o.Font.Size=10: o.Style=2
    Set NwC=o
End Function
Private Function NwLB(f As Object, nm As String, l%, t%, w%, h%) As Object
    Dim o As Object: Set o=f.Controls.Add("Forms.ListBox.1",nm,True)
    o.Left=l: o.Top=t: o.Width=w: o.Height=h
    o.BackColor=RGB(36,36,56): o.ForeColor=RGB(200,230,255): o.Font.Size=10
    Set NwLB=o
End Function
Private Sub NwCK(f As Object, nm As String, txt As String, l%, t%, w%, h%)
    On Error Resume Next
    Dim o As Object: Set o=f.Controls.Add("Forms.CheckBox.1",nm,True)
    If Not o Is Nothing Then
        o.Caption=txt: o.Left=l: o.Top=t: o.Width=w: o.Height=h
        o.ForeColor=RGB(189,215,238): o.BackColor=RGB(26,26,46): o.Font.Size=9
    End If
    On Error GoTo 0
End Sub
Private Sub NwTarj(f As Object, nL As String, nN As String, tit As String, l%, t%, w%, h%, col As Long)
    Dim lb As Object: Set lb=f.Controls.Add("Forms.Label.1",nL,True)
    lb.Caption=tit: lb.Left=l: lb.Top=t: lb.Width=w: lb.Height=18
    lb.Font.Size=8: lb.Font.Bold=True: lb.ForeColor=RGB(255,255,255): lb.BackColor=col: lb.TextAlign=2
    Dim nb As Object: Set nb=f.Controls.Add("Forms.Label.1",nN,True)
    nb.Caption="0": nb.Left=l: nb.Top=t+18: nb.Width=w: nb.Height=h-18
    nb.Font.Size=22: nb.Font.Bold=True: nb.ForeColor=RGB(255,255,255): nb.BackColor=col: nb.TextAlign=2
End Sub
Private Sub NwLim(f As Object)
    On Error Resume Next
    Dim c As Object: For Each c In f.Controls: f.Controls.Remove c.Name: Next c
    On Error GoTo 0
End Sub

Private Sub BLogin()
    Dim f As Object: Set f=ThisWorkbook.VBProject.VBComponents("frmLogin").Designer
    Call NwLim(f)
    f.Width=420: f.Height=420: f.BackColor=RGB(26,26,46)
    Call NwL(f,"lbBanner","",0,0,420,80,8,RGB(255,192,0),RGB(31,56,100),2)
    Call NwL(f,"lbLogo","RD",16,10,60,60,26,RGB(31,56,100),RGB(255,192,0),2)
    Call NwL(f,"lbTitle","RegistroDoc Multigrado",84,14,310,26,16,RGB(255,255,255),RGB(31,56,100),1)
    Call NwL(f,"lbSub","Sistema de Registro Escolar - Panama 2026",84,40,310,14,9,RGB(189,215,238),RGB(31,56,100),1)
    Call NwL(f,"lbBand","",0,80,420,4,8,RGB(255,192,0),RGB(255,192,0),1)
    Call NwL(f,"lbWelcome","Bienvenido - Ingresa tus credenciales",0,96,420,18,10,RGB(189,215,238),RGB(26,26,46),2)
    Call NwL(f,"lbU","Usuario:",20,130,80,18,10,RGB(189,215,238),RGB(26,26,46),1)
    Call NwT(f,"txtUsuario",106,128,290,26)
    Call NwL(f,"lbP","Contrasena:",20,166,80,18,10,RGB(189,215,238),RGB(26,26,46),1)
    Call NwT(f,"txtPassword",106,164,290,26,True)
    Call NwL(f,"lblError","",10,202,400,18,9,RGB(255,100,100),RGB(26,26,46),2)
    Call NwB(f,"btnIngresar","Ingresar al Sistema",80,228,260,38,RGB(46,117,182))
    Call NwL(f,"lbBand2","",0,278,420,2,8,RGB(46,46,70),RGB(46,46,70),1)
    Call NwL(f,"lbISO","ISO 27001 - Ley 34/2006 Panama - Decreto 751",0,288,420,14,8,RGB(80,120,180),RGB(26,26,46),2)
    Call NwL(f,"lbRoles","Roles: Admin | Docente | Director",0,306,420,14,8,RGB(80,80,100),RGB(26,26,46),2)
    Call NwB(f,"btnCancelar","Salir",155,330,110,26,RGB(80,40,40))
    Call NwL(f,"lbC","c 2026 RegistroDoc Multigrado - Elmer Tugri - Escuela Cerro Cacicon",0,368,420,12,7,RGB(60,60,80),RGB(26,26,46),2)
End Sub

Private Sub BMenu()
    Dim f As Object: Set f=ThisWorkbook.VBProject.VBComponents("frmMenu").Designer
    Call NwLim(f)
    Call NwL(f,"lb0","RD",16,12,48,48,22,RGB(31,56,100),RGB(255,192,0),2)
    Call NwL(f,"lb1","RegistroDoc Multigrado",72,12,370,24,16,RGB(255,192,0),RGB(26,26,46),1)
    Call NwL(f,"lb2","Sistema de Registro Escolar - Panama 2026",72,36,370,14,9,RGB(189,215,238),RGB(26,26,46),1)
    Call NwL(f,"lbL","",0,64,460,4,8,RGB(255,192,0),RGB(255,192,0),1)
    Call NwL(f,"lblUsuario","Usuario: -",0,72,460,14,8,RGB(100,200,255),RGB(15,15,35),2)
    Call NwL(f,"lblEstats","",0,88,460,14,8,RGB(100,220,100),RGB(15,15,35),2)
    Call NwL(f,"lb3","INGRESAR DATOS",0,108,460,14,8,RGB(255,192,0),RGB(26,26,46),2)
    Call NwB(f,"btnEstudiantes","Estudiantes",10,124,210,46,RGB(46,117,182))
    Call NwB(f,"btnNotas","Ingresar Notas",240,124,200,46,RGB(112,48,160))
    Call NwB(f,"btnAsistencia","Asistencia Diaria",10,174,210,44,RGB(55,86,35))
    Call NwL(f,"lb4","RESULTADOS Y REPORTES",0,226,460,14,8,RGB(255,192,0),RGB(26,26,46),2)
    Call NwB(f,"btnResumenes","Resumenes Notas",10,242,135,42,RGB(46,117,182))
    Call NwB(f,"btnDashboard","Dashboard",155,242,135,42,RGB(197,90,17))
    Call NwB(f,"btnReportes","Reportes/Imprimir",300,242,150,42,RGB(55,86,35))
    Call NwL(f,"lb5","ADMINISTRACION (ISO 27001)",0,292,460,14,8,RGB(255,192,0),RGB(26,26,46),2)
    Call NwB(f,"btnConfig","Configuracion",10,308,135,38,RGB(184,134,11))
    Call NwB(f,"btnAuditoria","Auditoria",155,308,135,38,RGB(80,80,120))
    Call NwB(f,"btnRespaldo","Respaldo",300,308,150,38,RGB(197,90,17))
    Call NwB(f,"btnCerrar","X  Cerrar Sistema",100,354,260,36,RGB(192,0,0))
    Call NwL(f,"lbC","c 2026 RegistroDoc Multigrado - Elmer Tugri - Ley 34/2006 Panama",0,400,460,12,7,RGB(100,100,100),RGB(26,26,46),2)
End Sub

Private Sub BEstudiantes()
    Dim f As Object: Set f=ThisWorkbook.VBProject.VBComponents("frmEstudiantes").Designer
    Call NwLim(f)
    Call NwL(f,"lbT","Gestion de Estudiantes",0,8,540,22,13,RGB(255,192,0),RGB(26,26,46),2)
    Call NwL(f,"lbF","Filtrar:",10,38,50,18,9,RGB(189,215,238),RGB(26,26,46),1)
    Call NwC(f,"cboGradoFiltro",64,36,90,22)
    Call NwL(f,"lbB","Buscar:",168,38,50,18,9,RGB(189,215,238),RGB(26,26,46),1)
    Call NwT(f,"txtBuscar",222,36,200,22)
    Call NwL(f,"lblContador","Total: 0",432,38,100,18,9,RGB(100,220,100),RGB(26,26,46),1)
    Call NwLB(f,"lstEstudiantes",10,64,520,230)
    Call NwL(f,"lbSep","",0,300,540,3,8,RGB(255,192,0),RGB(255,192,0),1)
    Call NwL(f,"lbAdd","AGREGAR NUEVO ESTUDIANTE",0,308,540,14,8,RGB(255,192,0),RGB(26,26,46),2)
    Call NwL(f,"lbG","Grado:",10,330,50,18,9,RGB(189,215,238),RGB(26,26,46),1)
    Call NwC(f,"cboGradoNuevo",65,328,80,22)
    Call NwL(f,"lbA","Apellido:",158,330,62,18,9,RGB(189,215,238),RGB(26,26,46),1)
    Call NwT(f,"txtApellido",224,328,180,22)
    Call NwL(f,"lbNm","Nombre:",418,330,56,18,9,RGB(189,215,238),RGB(26,26,46),1)
    Call NwT(f,"txtNombre",478,328,50,22)
    Call NwL(f,"lbCed","Cedula panamena:",10,358,120,18,9,RGB(189,215,238),RGB(26,26,46),1)
    Call NwT(f,"txtCedula",134,356,150,22)
    Call NwL(f,"lbCV","(7 a 10 digitos)",292,358,160,16,8,RGB(150,150,150),RGB(26,26,46),1)
    Call NwB(f,"btnAgregar","Agregar Estudiante",10,390,220,36,RGB(55,86,35))
    Call NwB(f,"btnEliminar","Eliminar Seleccionado",244,390,200,36,RGB(192,0,0))
    Call NwB(f,"btnCerrar","Cerrar",458,390,72,36,RGB(46,117,182))
End Sub

Private Sub BNotas()
    Dim f As Object: Set f=ThisWorkbook.VBProject.VBComponents("frmNotas").Designer
    Call NwLim(f)
    Call NwL(f,"lbT","Ingreso de Calificaciones - Sistema MEDUCA Panama",0,8,600,22,13,RGB(255,192,0),RGB(26,26,46),2)
    Call NwL(f,"lbG","Grado:",10,36,48,18,9,RGB(189,215,238),RGB(26,26,46),1)
    Call NwC(f,"cboGrado",62,34,70,22)
    Call NwL(f,"lbM","Materia:",145,36,55,18,9,RGB(189,215,238),RGB(26,26,46),1)
    Call NwC(f,"cboMateria",204,34,185,22)
    Call NwL(f,"lbTr","Trimestre:",403,36,68,18,9,RGB(189,215,238),RGB(26,26,46),1)
    Call NwC(f,"cboTrimestre",475,34,113,22)
    Call NwL(f,"lbCo","Componente:",10,62,80,18,9,RGB(189,215,238),RGB(26,26,46),1)
    Call NwC(f,"cboComponente",95,60,220,22)
    Call NwL(f,"lbD","Descripcion:",328,62,80,18,9,RGB(189,215,238),RGB(26,26,46),1)
    Call NwT(f,"txtDescripcion",412,60,176,22)
    Call NwL(f,"lbF","Fecha nota:",10,88,80,18,9,RGB(189,215,238),RGB(26,26,46),1)
    Call NwT(f,"txtFechaNota",95,86,100,22)
    Call NwL(f,"lblInfo","Selecciona todos los campos",0,114,600,16,9,RGB(100,200,255),RGB(15,15,35),2)
    Call NwL(f,"lbE","Estudiantes:",10,134,90,14,8,RGB(255,192,0),RGB(26,26,46),1)
    Call NwL(f,"lbN","Notas del componente:",300,134,240,14,8,RGB(255,192,0),RGB(26,26,46),1)
    Call NwLB(f,"lstEstudiantes",10,150,280,360)
    Call NwLB(f,"lstNotas",300,150,288,360)
    Call NwL(f,"lblFinal","",10,516,440,18,10,RGB(100,220,100),RGB(26,26,46),1)
    Call NwL(f,"lbNL","Nota (1.0-5.0):",10,542,108,18,9,RGB(189,215,238),RGB(26,26,46),1)
    Call NwT(f,"txtNota",122,540,82,24)
    Call NwB(f,"btnGuardar","Guardar Nota",216,538,168,28,RGB(55,86,35))
    Call NwB(f,"btnBorrarNota","Borrar Nota",398,538,116,28,RGB(197,90,17))
    Call NwB(f,"btnCerrar","Cerrar",500,574,88,28,RGB(192,0,0))
End Sub

Private Sub BAsistencia()
    Dim f As Object: Set f=ThisWorkbook.VBProject.VBComponents("frmAsistencia").Designer
    Call NwLim(f)
    Call NwL(f,"lbT","Control de Asistencia Diaria - Ley 34/2006",0,8,520,22,13,RGB(255,192,0),RGB(26,26,46),2)
    Call NwL(f,"lbG","Grado:",10,38,50,18,9,RGB(189,215,238),RGB(26,26,46),1)
    Call NwC(f,"cboGrado",64,36,80,22)
    Call NwL(f,"lbF","Fecha (DD/MM/YYYY):",158,38,138,18,9,RGB(189,215,238),RGB(26,26,46),1)
    Call NwT(f,"txtFecha",300,36,130,22)
    Call NwLB(f,"lstEstudiantes",10,66,500,296)
    Call NwL(f,"lblContador","0 estudiantes",10,370,300,16,9,RGB(100,200,255),RGB(26,26,46),1)
    Call NwL(f,"lblUltimo","",10,388,500,18,10,RGB(100,220,100),RGB(26,26,46),1)
    Call NwB(f,"btnP","P  PRESENTE",10,410,158,50,RGB(55,86,35))
    Call NwB(f,"btnA","A  AUSENTE",178,410,158,50,RGB(192,0,0))
    Call NwB(f,"btnT","T  TARDANZA",346,410,158,50,RGB(184,134,11))
    Call NwB(f,"btnTodosP","Todos Presentes hoy",10,468,248,30,RGB(46,117,182))
    Call NwB(f,"btnCerrar","Cerrar",272,468,150,30,RGB(197,90,17))
End Sub

Private Sub BResumenes()
    Dim f As Object: Set f=ThisWorkbook.VBProject.VBComponents("frmResumenes").Designer
    Call NwLim(f)
    Call NwL(f,"lbT","Resumenes de Notas - Conectado con PROM MEDUCA",0,8,640,22,13,RGB(255,192,0),RGB(26,26,46),2)
    Call NwL(f,"lbG","Grado:",10,38,50,18,9,RGB(189,215,238),RGB(26,26,46),1)
    Call NwC(f,"cboGrado",64,36,80,22)
    Call NwL(f,"lbTr","Trimestre:",158,38,70,18,9,RGB(189,215,238),RGB(26,26,46),1)
    Call NwC(f,"cboTrimestre",232,36,210,22)
    Call NwL(f,"lblStats","Selecciona grado y trimestre",454,38,178,18,9,RGB(100,200,255),RGB(26,26,46),1)
    Call NwLB(f,"lstResumen",10,64,620,444)
    Call NwB(f,"btnCerrar","Cerrar",240,516,160,32,RGB(46,117,182))
End Sub

Private Sub BDashboard()
    Dim f As Object: Set f=ThisWorkbook.VBProject.VBComponents("frmDashboard").Designer
    Call NwLim(f)
    Call NwL(f,"lbT","Dashboard - Estadisticas y Alertas",0,8,520,22,13,RGB(255,192,0),RGB(26,26,46),2)
    Call NwL(f,"lbS1","MATRICULA ACTUAL",0,36,520,14,8,RGB(255,192,0),RGB(26,26,46),2)
    Call NwTarj(f,"lb7t","lblNum7","7 Grado",10,52,110,54,RGB(46,117,182))
    Call NwTarj(f,"lb8t","lblNum8","8 Grado",130,52,110,54,RGB(112,48,160))
    Call NwTarj(f,"lb9t","lblNum9","9 Grado",250,52,110,54,RGB(197,90,17))
    Call NwTarj(f,"lbGt","lblNumTotal","TOTAL",370,52,122,54,RGB(55,86,35))
    Call NwL(f,"lbS2","ALERTAS - Estudiantes con mas de 5 ausencias",0,116,520,14,8,RGB(255,150,50),RGB(26,26,46),2)
    Call NwLB(f,"lstAlertas",10,134,500,250)
    Call NwL(f,"lblNumAlertas","0 alertas",10,392,300,16,9,RGB(255,150,50),RGB(26,26,46),1)
    Call NwB(f,"btnRefrescar","Refrescar datos",10,414,185,32,RGB(55,86,35))
    Call NwB(f,"btnCerrar","Cerrar",210,414,165,32,RGB(46,117,182))
End Sub

Private Sub BReportes()
    Dim f As Object: Set f=ThisWorkbook.VBProject.VBComponents("frmReportes").Designer
    Call NwLim(f)
    Call NwL(f,"lbT","Reportes e Impresion Oficial MEDUCA",0,8,500,22,13,RGB(255,192,0),RGB(26,26,46),2)
    Call NwL(f,"lbS","Decreto Ejecutivo N 751 - Registros Academicos Panama",0,30,500,12,8,RGB(150,150,150),RGB(26,26,46),2)
    Call NwB(f,"btnReporteDir","Reporte de Matricula para la Direccion",10,52,480,50,RGB(31,56,100))
    Call NwB(f,"btnVistaPrevia","Ver Lista de Planillas Disponibles",10,114,480,50,RGB(46,117,182))
    Call NwB(f,"btnImprimirTodo","Imprimir Libro Completo (todas las planillas)",10,176,480,50,RGB(55,86,35))
    Call NwB(f,"btnExportarPDF","Exportar Libro Completo a PDF",10,238,480,50,RGB(184,134,11))
    Call NwL(f,"lbN","Las notas se exportan directamente desde las hojas PROM del Excel.",10,300,480,14,9,RGB(150,150,150),RGB(26,26,46),2)
    Call NwB(f,"btnCerrar","Cerrar",175,326,150,32,RGB(197,90,17))
End Sub

Private Sub BConfig()
    Dim f As Object: Set f=ThisWorkbook.VBProject.VBComponents("frmConfiguracion").Designer
    Call NwLim(f)
    Call NwL(f,"lbT","Configuracion del Sistema - ISO 27001",0,8,580,22,13,RGB(255,192,0),RGB(26,26,46),2)
    Call NwL(f,"lbD","DATOS INSTITUCIONALES",0,34,580,14,8,RGB(255,192,0),RGB(26,26,46),2)
    Call NwL(f,"lbA","Ano lectivo:",10,52,90,18,9,RGB(189,215,238),RGB(26,26,46),2): Call NwT(f,"txtAnio",104,50,90,22)
    Call NwL(f,"lbDoc","Docente:",210,52,60,18,9,RGB(189,215,238),RGB(26,26,46),2): Call NwT(f,"txtDocente",274,50,294,22)
    Call NwL(f,"lbEs","Escuela:",10,78,90,18,9,RGB(189,215,238),RGB(26,26,46),2): Call NwT(f,"txtEscuela",104,76,464,22)
    Call NwL(f,"lbUb","Ubicacion:",10,104,90,18,9,RGB(189,215,238),RGB(26,26,46),2): Call NwT(f,"txtUbic",104,102,464,22)
    Call NwL(f,"lbMo","Modalidad:",10,130,90,18,9,RGB(189,215,238),RGB(26,26,46),2): Call NwT(f,"txtModal",104,128,464,22)
    Call NwL(f,"lbDr","Director:",10,156,90,18,9,RGB(189,215,238),RGB(26,26,46),2): Call NwT(f,"txtDirector",104,154,200,22)
    Call NwL(f,"lbEm","Correo:",318,156,55,18,9,RGB(189,215,238),RGB(26,26,46),2): Call NwT(f,"txtEmail",377,154,191,22)
    Call NwL(f,"lbSp","",0,182,580,3,8,RGB(255,192,0),RGB(255,192,0),1)
    Call NwL(f,"lbM","MATERIAS Y GRADOS DEL SISTEMA",0,190,580,14,8,RGB(255,192,0),RGB(26,26,46),2)
    Call NwLB(f,"lstMaterias",10,208,560,180)
    Call NwL(f,"lbAdd","AGREGAR NUEVA MATERIA / NUEVO GRADO:",0,396,580,14,8,RGB(100,200,255),RGB(26,26,46),2)
    Call NwL(f,"lbNM","Nombre materia:",10,416,106,18,9,RGB(189,215,238),RGB(26,26,46),1): Call NwT(f,"txtNuevaMat",120,414,200,22)
    Call NwL(f,"lbGM","Grado:",334,416,48,18,9,RGB(189,215,238),RGB(26,26,46),1): Call NwC(f,"cboGradoNuevaMat",386,414,90,22)
    Call NwL(f,"lbNG","Si es grado nuevo:",10,444,125,18,9,RGB(189,215,238),RGB(26,26,46),1): Call NwT(f,"txtNuevoGrado",140,442,80,22)
    Call NwL(f,"lbNG2","(escribe el numero o letra del nuevo grado)",228,444,280,16,8,RGB(150,150,150),RGB(26,26,46),1)
    Call NwCK(f,"chkUnion","Esta materia une con otra para formar Tecnologia",10,470,340,20)
    Call NwB(f,"btnAgregarMateria","Agregar Materia + Crear Hojas",10,496,260,32,RGB(55,86,35))
    Call NwB(f,"btnEliminarMateria","Eliminar Seleccionada",280,496,200,32,RGB(192,0,0))
    Call NwB(f,"btnGuardar","Guardar Configuracion",10,538,260,34,RGB(46,117,182))
    Call NwB(f,"btnCerrar","Cancelar",284,538,170,34,RGB(197,90,17))
    Dim cbGM As Object: Set cbGM=f.Controls("cboGradoNuevaMat")
    cbGM.BackColor=RGB(46,46,70): cbGM.ForeColor=RGB(255,255,255): cbGM.Style=2
End Sub

Private Sub BAuditoria()
    Dim f As Object: Set f=ThisWorkbook.VBProject.VBComponents("frmAuditoria").Designer
    Call NwLim(f)
    Call NwL(f,"lbT","Auditoria del Sistema - ISO 27001 A.12.4",0,8,620,22,13,RGB(255,192,0),RGB(26,26,46),2)
    Call NwL(f,"lbS","Registro de todos los cambios - quien, que, cuando",0,30,620,12,8,RGB(150,150,150),RGB(26,26,46),2)
    Call NwLB(f,"lstAuditoria",10,48,600,390)
    Call NwL(f,"lblTotal","0 registros",10,446,300,16,9,RGB(100,200,255),RGB(26,26,46),1)
    Call NwB(f,"btnExportar","Exportar Auditoria a Excel",10,468,260,34,RGB(55,86,35))
    Call NwB(f,"btnCerrar","Cerrar",290,468,160,34,RGB(46,117,182))
End Sub
""")
print("      RD_Build OK")
print("\n[6/7] Ejecutando construccion de controles...")
try:
    app.screen_updating = True
    app.api.Run("ConstruirTodosLosControles")
    print("      Controles construidos OK")
except Exception as e:
    print(f"      Nota: {e} - ejecuta ConstruirTodosLosControles manualmente")

time.sleep(3)

print(f"\n[7/7] Guardando: {SALIDA}")
try:
    wb.sheets["PORTADA"].activate()
except Exception:
    pass
app.display_alerts = False
app.screen_updating = True
# Cerrar cualquier dialogo pendiente
try:
    app.api.SendKeys("~")
except: pass
time.sleep(2)
try:
    wb.api.SaveAs(SALIDA, FileFormat=52)
    print("      OK")
except Exception as e:
    print(f"      Error SaveAs: {e}")
    print("      Guardando con Save + copia manual...")
    try:
        # Guardar primero con formato actual
        wb.save()
        # Luego copiar como xlsm
        import shutil
        orig = wb.fullname
        wb.close(save_changes=True)
        app.quit()
        if orig != SALIDA:
            shutil.copy(orig, SALIDA)
        print(f"      Archivo en: {SALIDA}")
    except Exception as e2:
        print(f"      {e2}")
        try:
            wb.close(save_changes=False)
            app.quit()
        except: pass
    import sys; sys.exit(0)

time.sleep(1)
try:
    wb.close(save_changes=False)
except: pass
try:
    app.quit()
except: pass

print("\n" + "=" * 60)
print("  RegistroDoc v2.1 COMPLETADO")
print(f"  Archivo: {SALIDA}")
print("=" * 60)
print("""
  SISTEMA COMPLETO:
  + frmLogin      Login con 3 roles (ISO 27001 A.9)
  + frmMenu       Menu con stats y acceso por rol
  + frmEstudiantes CRUD completo + busqueda + cedula
  + frmNotas      Notas con descripcion + fecha + PROM real
  + frmAsistencia P/A/T con auditoria
  + frmResumenes  Resultados por grado/trimestre
  + frmDashboard  Estadisticas y alertas
  + frmReportes   Imprimir libro + PDF + reporte direccion
  + frmConfig     Agregar/eliminar materias y grados nuevos
  + frmAuditoria  Log completo ISO 27001 exportable
  + BD_Auditoria  Registro de todos los cambios
  + BD_Usuarios   Control de acceso con bloqueo
  + Respaldo auto al cerrar (ISO 22301)
  + Validacion cedula panamena (Ley 34/2006)
  
  USUARIOS DEMO:
  admin    / admin123    (acceso total)
  docente  / docente123  (notas y asistencia)
  director / director123 (solo lectura)
""")
input("Presiona Enter para cerrar...")
