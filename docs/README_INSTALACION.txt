╔══════════════════════════════════════════════════════════════════════╗
║         REGISTRODOC PRO v5.0 — GUÍA DE INSTALACIÓN Y DESARROLLO      ║
║         Para el desarrollador — Compilación y Empaquetado            ║
╚══════════════════════════════════════════════════════════════════════╝

Esta guía explica cómo configurar el proyecto, probarlo localmente y compilar
los instaladores distribuidos (.exe) usando PyInstaller e Inno Setup.

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
PASO 1 — CONFIGURAR ENTORNO DE DESARROLLO
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

  1. Abre VS Code en la carpeta raíz del proyecto.
  2. Asegúrate de tener Python 3.10 o superior instalado.
  3. Instala las dependencias requeridas ejecutando en la terminal:
       pip install -r requirements.txt

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
PASO 2 — EJECUTAR EN MODO DESARROLLO (PRUEBAS LOCALES)
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

  Para ejecutar y depurar la aplicación sin compilar:
  
  $env:REGISTRODOC_DEV_MODE="1"; python src/app.py

  Esto habilitará la base de datos de pruebas temporales y abrirá
  la interfaz de usuario directamente.

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
PASO 3 — COMPILAR EL EJECUTABLE (.EXE PORTABLE)
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

  Antes de generar el instalador de Windows, debemos compilar el código
  de Python a un ejecutable binario mediante PyInstaller:

  1. En la raíz del proyecto, ejecuta:
       pyinstaller RegistroDoc.spec --clean

  2. Esto creará el ejecutable compilado en:
       dist/RegistroDoc.exe

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
PASO 4 — GENERAR LOS INSTALADORES DE WINDOWS (INNO SETUP)
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

  Utilizamos Inno Setup para crear los instaladores de nivel de producción.
  Hay dos configuraciones según el caso de uso del docente:

  OPCIÓN A: Instalador Limpio ("Sin Datos")
  -------------------------------------------------------------
  Diseñado para instalaciones nuevas o limpias desde cero. No contiene
  bases de datos ni plantillas preexistentes en la carpeta AppData del usuario.
  
  * Script: RegistroDoc_Setup.iss
  * Ejecutar en terminal:
      & "C:\Program Files (x86)\Inno Setup 6\ISCC.exe" RegistroDoc_Setup.iss
  * Genera el instalador en:
      Output/RegistroDoc_Instalador.exe

  OPCIÓN B: Instalador con Datos ("Con Datos" / Actualización segura)
  -------------------------------------------------------------
  Diseñado para empaquetar una base de datos específica o actualizar
  versiones anteriores sin borrar ni sobreescribir la información del docente.
  Utiliza banderas `onlyifdoesntexist` para proteger bases de datos y
  archivos Excel ya existentes.

  * Script: RegistroDoc_Setup_ConDatos.iss
  * Ejecutar en terminal:
      & "C:\Program Files (x86)\Inno Setup 6\ISCC.exe" RegistroDoc_Setup_ConDatos.iss
  * Genera el instalador en:
      Output/RegistroDoc_Instalador_ConDatos.exe

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
© 2026 RegistroDoc Pro — MEDUCA Panamá
