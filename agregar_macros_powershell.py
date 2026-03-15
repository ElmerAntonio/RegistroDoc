#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Agregar RD_Macros.bas al RegistroDoc_PRO.xlsm automáticamente
Usa PowerShell para interactuar con Excel COM
"""

import os
import subprocess

def agregar_macros_con_powershell():
    """Usa PowerShell para agregar .bas a través de Excel COM"""
    
    # Script PowerShell que agrega el módulo VBA
    ps_script = r'''
# Habilitar interacción con COM de Excel
$ErrorActionPreference = "SilentlyContinue"

try {
    $ExcelApp = New-Object -ComObject Excel.Application
    $ExcelApp.Visible = $false
    $ExcelApp.DisplayAlerts = $false
    
    $archivo = "C:\RegistroDoc\RegistroDoc_PRO.xlsm"
    $basfile = "C:\RegistroDoc\RD_Macros.bas"
    
    Write-Host "📂 Abriendo Excel..." -ForegroundColor Cyan
    $wb = $ExcelApp.Workbooks.Open($archivo, $null, $false, 1)
    
    if ($wb) {
        Write-Host "✅ Archivo abierto correctamente" -ForegroundColor Green
        
        # Acceder al proyecto VBA
        $vbProject = $wb.VBProject
        
        if ($vbProject) {
            Write-Host "📋 Proyecto VBA detectado" -ForegroundColor Cyan
            
            # Importar el módulo .bas
            Write-Host "📥 Importando módulo desde: $basfile" -ForegroundColor Cyan
            
            try {
                $vbProject.VBComponents.Import($basfile)
                Write-Host "✅ Módulo VBA importado exitosamente" -ForegroundColor Green
                
                # Guardar el archivo
                Write-Host "💾 Guardando archivo..." -ForegroundColor Cyan
                $wb.Save()
                Write-Host "✅ Archivo guardado con macros" -ForegroundColor Green
                
            } catch {
                Write-Host "❌ Error al importar: $_" -ForegroundColor Red
            }
        } else {
            Write-Host "❌ No se pudo acceder al proyecto VBA" -ForegroundColor Red
        }
        
        # Cerrar el archivo
        $wb.Close($false)
        Write-Host "✅ Archivo cerrado" -ForegroundColor Green
    } else {
        Write-Host "❌ No se pudo abrir el archivo" -ForegroundColor Red
    }
    
    # Cerrar Excel
    $ExcelApp.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ExcelApp) | Out-Null
    
} catch {
    Write-Host "❌ Error crítico: $_" -ForegroundColor Red
}

Write-Host "`n✅ Proceso completado" -ForegroundColor Green
'''
    
    # Guardar script en archivo temporal
    ps_file = 'temp_add_macros.ps1'
    with open(ps_file, 'w', encoding='utf-8') as f:
        f.write(ps_script)
    
    try:
        # Ejecutar PowerShell con permisos elevados si es necesario
        print("🔄 Ejecutando PowerShell para agregar macros...")
        print("=" * 60)
        
        result = subprocess.run(
            ['powershell', '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', ps_file],
            capture_output=False,
            text=True
        )
        
        print("=" * 60)
        
        # Limpiar archivo temporal
        if os.path.exists(ps_file):
            os.remove(ps_file)
        
        return result.returncode == 0
        
    except Exception as e:
        print(f"❌ Error al ejecutar PowerShell: {str(e)}")
        if os.path.exists(ps_file):
            os.remove(ps_file)
        return False

def verificar_archivos():
    """Verifica que existan los archivos necesarios"""
    
    requeridos = [
        'RegistroDoc_PRO.xlsm',
        'RD_Macros.bas'
    ]
    
    print("📋 Verificando archivos necesarios...")
    todos_existen = True
    
    for archivo in requeridos:
        if os.path.exists(archivo):
            tamaño = os.path.getsize(archivo)
            print(f"   ✅ {archivo} ({tamaño} bytes)")
        else:
            print(f"   ❌ {archivo} - NO ENCONTRADO")
            todos_existen = False
    
    print()
    return todos_existen

if __name__ == '__main__':
    print("=" * 60)
    print("AGREGAR MACROS A RegistroDoc_PRO.xlsm")
    print("=" * 60)
    print()
    
    # Verificar archivos
    if not verificar_archivos():
        print("❌ Faltan archivos requeridos")
        exit(1)
    
    # Agregar macros
    exito = agregar_macros_con_powershell()
    
    if exito:
        print()
        print("=" * 60)
        print("✅ MACROS AGREGADOS CORRECTAMENTE")
        print("=" * 60)
        print()
        print("📌 Próximo paso:")
        print("   python agregar_formularios.py")
        print()
    else:
        print()
        print("❌ No se pudieron agregar los macros automáticamente")
        print()
        print("📌 Hacerlo manualmente:")
        print("   1. Abre RegistroDoc_PRO.xlsm en Excel")
        print("   2. Presiona Alt+F11 para abrir Visual Basic")
        print("   3. File > Import File...")
        print("   4. Selecciona RD_Macros.bas")
        print("   5. Haz clic en Abrir")
        print("   6. Guarda el archivo (Ctrl+S)")
        print()
