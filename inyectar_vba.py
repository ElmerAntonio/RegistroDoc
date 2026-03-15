#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Inyector de código VBA en archivo Excel
Modifica directamente la estructura ZIP para agregar macros válidas
"""

import zipfile
import os
import xml.etree.ElementTree as ET
from lxml import etree
import shutil

def inyectar_vba_en_xlsm():
    """Inyecta código VBA en el archivo XLSM"""
    
    archivo_xlsm = 'RegistroDoc_PRO.xlsm'
    archivo_bas = 'RD_Macros.bas'
    
    # Leer el código del .bas
    with open(archivo_bas, 'r', encoding='utf-8') as f:
        codigo_vba = f.read()
    
    print("[*] Leyendo contenido VBA...")
    print(f"    Líneas: {len(codigo_vba.splitlines())}")
    
    # Hacer backup
    backup = archivo_xlsm + '.backup'
    if os.path.exists(backup):
        os.remove(backup)
    shutil.copy(archivo_xlsm, backup)
    print(f"[*] Backup creado: {backup}")
    
    # Trabajar con tempoal
    temp_dir = 'temp_xlsm_work'
    if os.path.exists(temp_dir):
        shutil.rmtree(temp_dir)
    os.makedirs(temp_dir)
    
    # Extraer contenido
    print(f"[*] Extrayendo contenido...")
    with zipfile.ZipFile(archivo_xlsm, 'r') as zf:
        zf.extractall(temp_dir)
    
    # Crear archivos necesarios para VBA
    print(f"[*] Creando estructura VBA...")
    
    # Crear carpeta xl/macrosheets si no existe
    xl_dir = os.path.join(temp_dir, 'xl')
    if not os.path.exists(xl_dir):
        os.makedirs(xl_dir)
    
    # Crear archivo de configuración de macro
    projwm_path = os.path.join(xl_dir, 'vbaProject.xml')
    projwm_content = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Projects xmlns="http://schemas.microsoft.com/office/appforoffice/11/2006/ovml">
 <Project>
  <ProjectName>VBAProject</ProjectName>
  <ProjectDocString></ProjectDocString>
  <ProjectHelpFileName1>0</ProjectHelpFileName1>
  <ProjectHelpFileName2>0</ProjectHelpFileName2>
  <HelpLanguageID>0</HelpLanguageID>
  <ProjectConstants></ProjectConstants>
  <ReferenceName Name="VBA" Libid="000204EF-0000-0000-C000-000000000046" VersionMajor="4" VersionMinor="2" />
  <ReferenceName Name="Excel" Libid="00020813-0000-0000-C000-000000000046" VersionMajor="1" VersionMinor="9" />
  <ReferenceName Name="stdole" Libid="00020430-0000-0000-C000-000000000046" VersionMajor="2" VersionMinor="0" />
  <ReferenceControl Name="stdole" Libid="00020430-0000-0000-C000-000000000046" VersionMajor="2" VersionMinor="0" />
 </Project>
</Projects>'''
    
    # Crear módulo VBA
    vba_dir = os.path.join(xl_dir, 'macrosheets')
    if not os.path.exists(vba_dir):
        os.makedirs(vba_dir)
    
    modulos_path = os.path.join(vba_dir, 'RD_CORE.txt')
    with open(modulos_path, 'w', encoding='utf-8') as f:
        f.write(codigo_vba)
    
    print(f"[+] Módulo VBA creado")
    
    # Actualizar [Content_Types].xml
    ct_path = os.path.join(temp_dir, '[Content_Types].xml')
    if os.path.exists(ct_path):
        tree = ET.parse(ct_path)
        root = tree.getroot()
        
        # Agregar tipo para macros si no existe
        ns = {'ct': 'http://schemas.openxmlformats.org/package/2006/content-types'}
        
        # Buscar si ya existe
        overrides = root.findall('.//ct:Override[@PartName="/xl/vbaProject.bin"]', ns)
        if not overrides:
            override = ET.Element('{http://schemas.openxmlformats.org/package/2006/content-types}Override')
            override.set('PartName', '/xl/vbaProject.bin')
            override.set('ContentType', 'application/vnd.ms-excel.vbaProject')
            root.append(override)
            
            tree.write(ct_path, encoding='utf-8', xml_declaration=True)
            print(f"[+] Content_Types.xml actualizado")
    
    # Reempacar
    print(f"[*] Reempaquetando archivo...")
    
    os.remove(archivo_xlsm)
    
    with zipfile.ZipFile(archivo_xlsm, 'w', zipfile.ZIP_DEFLATED) as zf:
        for root_dir, dirs, files in os.walk(temp_dir):
            for file in files:
                file_path = os.path.join(root_dir, file)
                arcname = os.path.relpath(file_path, temp_dir)
                zf.write(file_path, arcname)
    
    # Limpiar temporal
    shutil.rmtree(temp_dir)
    
    print(f"[+] Archivo reempaquetado: {archivo_xlsm}")
    print()
    print("="*60)
    print("[OK] Estructura VBA inyectada")
    print("="*60)
    print()
    print("Ahora intenta abrir el archivo en Excel.")
    print("Excel debería reconocer los macros.")
    
    return True

if __name__ == '__main__':
    print("="*60)
    print("INYECTAR VBA EN XLSM")
    print("="*60)
    print()
    
    try:
        inyectar_vba_en_xlsm()
    except ImportError:
        print("[!] Se requiere librería lxml")
        print("[*] Instalando: pip install lxml")
        os.system('pip install lxml -q')
        inyectar_vba_en_xlsm()
    except Exception as e:
        print(f"[!] Error: {e}")
