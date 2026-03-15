#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Agregar RD_Macros.bas al RegistroDoc_PRO.xlsm
Versión simplificada sin caracteres especiales
"""

import os
import subprocess
import sys

def agregar_macros_excel():
    """Script PowerShell simple para agregar macros"""
    
    ps_script = '''
$ErrorActionPreference = "SilentlyContinue"

try {
    $ExcelApp = New-Object -ComObject Excel.Application
    $ExcelApp.Visible = $false
    $ExcelApp.DisplayAlerts = $false
    
    $archivo = "C:\\RegistroDoc\\RegistroDoc_PRO.xlsm"
    $basfile = "C:\\RegistroDoc\\RD_Macros.bas"
    
    Write-Host "Abriendo Excel..."
    $wb = $ExcelApp.Workbooks.Open($archivo, $null, $false, 1)
    
    if ($wb) {
        Write-Host "Archivo abierto correctamente"
        
        $vbProject = $wb.VBProject
        
        if ($vbProject) {
            Write-Host "Importando modulo VBA..."
            
            try {
                $vbProject.VBComponents.Import($basfile)
                Write-Host "EXITO: Modulo VBA importado"
                
                Write-Host "Guardando archivo..."
                $wb.Save()
                Write-Host "EXITO: Archivo guardado con macros"
                $exitCode = 0
                
            } catch {
                Write-Host "ERROR: No se pudo importar"
                $exitCode = 1
            }
        }
        
        $wb.Close($false)
    }
    
    $ExcelApp.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ExcelApp) | Out-Null
    exit $exitCode
    
} catch {
    Write-Host "ERROR: $_"
    exit 1
}
'''
    
    # Guardar script
    ps_file = os.path.join(os.getcwd(), 'temp_add_macros_simple.ps1')
    
    try:
        with open(ps_file, 'w') as f:
            f.write(ps_script)
        
        print("[*] Ejecutando PowerShell para agregar macros...")
        
        # Ejecutar
        result = subprocess.run(
            ['powershell', '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', ps_file],
            capture_output=True,
            text=True
        )
        
        # Mostrar output
        if result.stdout:
            print(result.stdout.strip())
        
        # Limpiar
        if os.path.exists(ps_file):
            os.remove(ps_file)
        
        return result.returncode == 0
        
    except Exception as e:
        print(f"[!] Error: {str(e)}")
        if os.path.exists(ps_file):
            os.remove(ps_file)
        return False

if __name__ == '__main__':
    print("="*60)
    print("AGREGAR MACROS A RegistroDoc_PRO.xlsm")
    print("="*60)
    print()
    
    # Verificar archivos
    if not os.path.exists('RegistroDoc_PRO.xlsm'):
        print("[!] Archivo RegistroDoc_PRO.xlsm no encontrado")
        sys.exit(1)
    
    if not os.path.exists('RD_Macros.bas'):
        print("[!] Archivo RD_Macros.bas no encontrado")
        sys.exit(1)
    
    # Agregar macros
    exito = agregar_macros_excel()
    
    print()
    if exito:
        print("="*60)
        print("[OK] MACROS AGREGADOS CORRECTAMENTE")
        print("="*60)
        print()
        print("Archivo RegistroDoc_PRO.xlsm listo para usar")
    else:
        print("[!] Error al agregar macros")
        print()
        print("Intenta manualmente:")
        print("1. Abre RegistroDoc_PRO.xlsm en Excel")
        print("2. Alt+F11 para abrir Visual Basic")
        print("3. File > Import File...")
        print("4. Selecciona RD_Macros.bas")
        print("5. Guarda (Ctrl+S)")
