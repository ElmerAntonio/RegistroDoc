#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Crear VBScript para agregar macros
Se ejecuta directamente sin COM issues
"""

import os
import subprocess

def crear_y_ejecutar_vbscript():
    """Crea y ejecuta un VBScript que importa el módulo VBA"""
    
    vbscript_content = '''
' VBScript para agregar RD_Macros.bas a RegistroDoc_PRO.xlsm

Dim objExcel
Dim objWorkbook
Dim objVBProject
Dim strFilePath
Dim strMacroPath

' Rutas
strFilePath = "C:\\RegistroDoc\\RegistroDoc_PRO.xlsm"
strMacroPath = "C:\\RegistroDoc\\RD_Macros.bas"

' Crear objeto Excel
On Error Resume Next
Set objExcel = CreateObject("Excel.Application")

If Err.Number <> 0 Then
    WScript.Echo "Error: No se puede crear objeto Excel"
    WScript.Echo "Err: " & Err.Number & " - " & Err.Description
    WScript.Quit 1
End If

On Error GoTo 0

objExcel.Visible = False
objExcel.DisplayAlerts = False

' Abrir archivo
On Error Resume Next
Set objWorkbook = objExcel.Workbooks.Open(strFilePath, , , 1)

If Err.Number <> 0 Then
    WScript.Echo "Error: No se puede abrir " & strFilePath
    WScript.Echo "Err: " & Err.Number & " - " & Err.Description
    objExcel.Quit
    WScript.Quit 1
End If

On Error GoTo 0

WScript.Echo "Archivo abierto: " & strFilePath

' Obtener proyecto VBA
Set objVBProject = objWorkbook.VBProject

' Importar módulo
On Error Resume Next
objVBProject.VBComponents.Import(strMacroPath)

If Err.Number <> 0 Then
    WScript.Echo "Error: No se puede importar " & strMacroPath
    WScript.Echo "Err: " & Err.Number & " - " & Err.Description
    objWorkbook.Close
    objExcel.Quit
    WScript.Quit 1
End If

On Error GoTo 0

WScript.Echo "Módulo importado: RD_Macros.bas"

' Guardar archivo
objWorkbook.Save
WScript.Echo "Archivo guardado"

' Cerrar
objWorkbook.Close
objExcel.Quit

WScript.Echo "Completado"
WScript.Quit 0
'''
    
    # Guardar VBScript
    vbs_file = 'agregar_macros.vbs'
    
    with open(vbs_file, 'w', encoding='utf-8') as f:
        f.write(vbscript_content)
    
    print(f"[*] Script VBScript creado: {vbs_file}")
    print(f"[*] Ejecutando...")
    print()
    
    # Ejecutar VBScript
    try:
        result = subprocess.run(
            ['cscript.exe', vbs_file],
            capture_output=True,
            text=True
        )
        
        # Mostrar salida
        if result.stdout:
            print(result.stdout.strip())
        
        # Limpiar
        if os.path.exists(vbs_file):
            os.remove(vbs_file)
        
        return result.returncode == 0
        
    except Exception as e:
        print(f"[!] Error ejecutando VBScript: {e}")
        if os.path.exists(vbs_file):
            os.remove(vbs_file)
        return False

if __name__ == '__main__':
    print("="*60)
    print("AGREGAR MACROS CON VBSCRIPT")
    print("="*60)
    print()
    
    exito = crear_y_ejecutar_vbscript()
    
    print()
    if exito:
        print("="*60)
        print("[OK] MACROS AGREGADOS")
        print("="*60)
    else:
        print()
        print("[!] Intenta manualmente:")
        print("1. Abre RegistroDoc_PRO.xlsm en Excel")
        print("2. Alt+F11")  
        print("3. File > Import File > RD_Macros.bas")
        print("4. Ctrl+S para guardar")
