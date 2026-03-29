╔══════════════════════════════════════════════════════════════════════╗
║         REGISTRODOC PRO — GUÍA DE INSTALACIÓN Y DESARROLLO          ║
║         Para el desarrollador — Paso a paso en VS Code              ║
╚══════════════════════════════════════════════════════════════════════╝

Esta guía explica cómo configurar el proyecto en VS Code,
probarlo y generar el instalador .exe para distribuir.

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
PASO 1 — CLONAR EL REPOSITORIO
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

  1. Abre VS Code
  2. Presiona Ctrl + ` para abrir la terminal integrada
  3. Navega a donde quieres el proyecto:
       cd C:\Users\TuUsuario\Documents
  4. Clona el repositorio:
       git clone https://github.com/ElmerAntonio/RegistroDoc.git
  5. Entra a la carpeta:
       cd RegistroDoc

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
PASO 2 — INSTALAR DEPENDENCIAS (SOLO UNA VEZ)
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

  En la terminal de VS Code ejecuta:

    pip install openpyxl cryptography pillow pyinstaller

  Verifica que todo esté instalado:
    python -c "import openpyxl, cryptography; print('OK')"

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
PASO 3 — ESTRUCTURA DE ARCHIVOS
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

  Tu carpeta debe verse así:

  RegistroDoc/
  ├── src/
  │   ├── app.py                  ← Programa principal del docente
  │   ├── rdsecurity.py           ← Módulo de seguridad AES-256
  │   ├── rdprint.py              ← Módulo de impresión
  │   ├── generador_codigos.py    ← Tu herramienta privada
  │   └── Registro_2026.xlsx      ← El Excel oficial
  ├── docs/
  │   ├── MANUAL_USUARIO.txt
  │   ├── PREGUNTAS_FRECUENTES.txt
  │   └── LICENCIA_Y_TERMINOS.txt
  ├── assets/
  │   └── icon.ico                ← Icono del programa (opcional)
  ├── setup.py                    ← Genera el .exe
  └── README_INSTALACION.txt      ← Este archivo

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
PASO 4 — PROBAR EN VS CODE (SIN GENERAR .EXE)
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

  Para probar el programa sin activación (modo desarrollo):

  1. En VS Code, abre el archivo: src/app.py
  2. En la terminal:
       cd src
       python app.py

  La primera vez pedirá activación. Para pruebas usa:
    Cédula:   4-785-823
    Código:   (genera uno con generador_codigos.py primero)

  Para probar el GENERADOR DE CÓDIGOS:
    python generador_codigos.py

  NOTA: Asegúrate de estar en la carpeta src/ antes de ejecutar.

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
PASO 5 — INTEGRAR EL MÓDULO DE IMPRESIÓN
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

  En app.py, agrega la importación al inicio:

    from rdprint import PanelImpresion

  Y en la función _construir_ui, agrega la pestaña:

    nb.add(PanelImpresion(nb, self), text="🖨️  Imprimir")

  (Agrégala después de PanelResumen y antes de PanelAuditoria)

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
PASO 6 — AGREGAR ICONO (OPCIONAL PERO RECOMENDADO)
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

  1. Crea o descarga un archivo icon.ico (tamaño 256x256)
  2. Colócalo en la carpeta assets/
  3. El setup.py lo tomará automáticamente

  Para convertir un PNG a ICO gratis:
  → https://convertio.co/png-ico/

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
PASO 7 — GENERAR EL EJECUTABLE .EXE
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

  IMPORTANTE: Ejecutar desde la raíz del proyecto (no desde src/).

  En la terminal de VS Code:

    python setup.py

  El proceso toma 2-5 minutos. Al terminar verás:
    dist/RegistroDocPro.exe    ← Este es el archivo para distribuir

  El .exe incluye TODO adentro:
    • El programa completo
    • El módulo de seguridad
    • El Excel de registro
    • Python (no necesita instalación separada)

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
PASO 8 — PREPARAR EL PAQUETE PARA DISTRIBUCIÓN
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

  Lo que le entregas al docente:
    ✓ RegistroDocPro.exe       ← El programa completo
    ✓ MANUAL_USUARIO.txt       ← Para que sepa usarlo
    ✓ PREGUNTAS_FRECUENTES.txt ← Resolución de dudas
    ✓ LICENCIA_Y_TERMINOS.txt  ← Acuerdo legal

  Lo que NUNCA debes distribuir:
    ✗ generador_codigos.py
    ✗ rdsecurity.py
    ✗ app.py
    ✗ setup.py
    ✗ Cualquier archivo .py

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
PASO 9 — FLUJO DE VENTA
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

  1. El docente te contacta y te da su número de CÉDULA
  2. Tú abres generador_codigos.py
  3. Ingresas la cédula y el nombre → obtienes el código RD-XXXXX-XXXXX-XXXXX-XXXXX
  4. Le envías al docente:
     • El archivo RegistroDocPro.exe
     • Su código de activación
     • El manual (opcional pero recomendado)
  5. El docente abre el .exe, ingresa su cédula + código → activado para siempre
  6. Funciona solo en esa computadora, no puede pasarlo a otra persona

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
SOLUCIÓN DE PROBLEMAS EN DESARROLLO
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

ERROR: ModuleNotFoundError: No module named 'cryptography'
SOLUCIÓN: pip install cryptography

ERROR: ModuleNotFoundError: No module named 'openpyxl'
SOLUCIÓN: pip install openpyxl

ERROR: PyInstaller no genera el .exe
SOLUCIÓN: Asegúrate de tener la última versión:
          pip install --upgrade pyinstaller

ERROR: "licencia.dat no encontrado" al probar
SOLUCIÓN: Normal, es la primera vez. Activa con el generador primero.

ERROR: El .exe se abre y cierra inmediatamente
SOLUCIÓN: Ejecuta desde terminal para ver el error:
          cd dist && RegistroDocPro.exe

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
SUBIR CAMBIOS A GITHUB
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

  Solo sube el código fuente, NUNCA el .exe ni las licencias generadas:

    git add src/app.py src/rdsecurity.py src/rdprint.py
    git add docs/ setup.py README_INSTALACION.txt
    git commit -m "Actualización v2.0"
    git push

  El .gitignore debe incluir:
    dist/
    build/
    *.spec
    licencia.dat
    rd_*.bin
    config.enc
    ventas_privado.enc
    Registro_2026_bak.xlsx

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

© 2026 RegistroDoc Pro — Elmer Tugri — Panamá
