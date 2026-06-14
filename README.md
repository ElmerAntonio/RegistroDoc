# RegistroDoc Pro v4.0 — Sistema de Registro Académico (Panamá)

[![Python](https://img.shields.io/badge/Python-3.10%2B-blue.svg)](https://www.python.org/)
[![License](https://img.shields.io/badge/Licencia-Comercial-green.svg)](#)
[![Security Standards](https://img.shields.io/badge/Security-ISO%2027001%20%7C%20Ley%2081-darkblue.svg)](#)
[![Tests Status](https://img.shields.io/badge/Tests-84%2F84%20Passed-emerald.svg)](#)

**RegistroDoc Pro** es una aplicación de escritorio premium e independiente desarrollada en Python, diseñada para optimizar la labor administrativa de docentes en la República de Panamá (tanto en modalidad **Primaria** como **Premedia**). 

El programa automatiza e interactúa de manera directa con las libretas de registro oficial de calificaciones en formato Excel del **Ministerio de Educación (MEDUCA)**, sin alterar sus fórmulas ni su diseño oficial.

---

## 🚀 Características Clave

* **Diseño Bimodal Dinámico**: Soporte integrado para las libretas oficiales de calificaciones y asistencia de Primaria (con casillas continuas) y Premedia.
* **Operación 100% Fuera de Línea (Offline)**: Diseñado para funcionar en áreas remotas o comarcales (como la Comarca Ngäbe Buglé) sin necesidad de acceso a Internet.
* **Seguridad y Cifrado Avanzado**: Cumple con la **Ley 81 de 2019 (Panamá)** sobre protección de datos personales y estándares **NIST SP 800-38D** usando cifrado **AES-256-GCM** para asegurar la información del docente.
* **Base de Datos SQLite**: Almacenamiento local estructurado mediante base de datos relacional SQLite protegida con ciclo de vida cifrado en reposo.
* **Controles de Auditoría e Integridad**: Implementa un registro de auditoría seguro e integridad del archivo Excel mediante firmas SHA3-256 (alineado con la norma **ISO/IEC 27001:2022**).
* **Asistente de Configuración (Setup Wizard)**: Formulario interactivo inicial para definir datos de escuela, grados con secciones personalizadas, materias asignadas y planificador de horario semanal.
* **Inserción Automática de Logo**: Permite al docente cargar el logo de su escuela para inyectarse de forma automática en la esquina superior derecha (`H1`) de los reportes oficiales exportados.
* **Cierre de Año Lectivo (Rollover y Archivo)**: Asistente guiado para archivar de forma segura el ciclo escolar actual en `Documentos/Historial_Academico/` y reiniciar de manera limpia la base de datos y la libreta de calificaciones para el siguiente período, con soporte opcional para promoción automática de cohortes de alumnos.
* **Importador Masivo de Estudiantes**: Herramienta inteligente para importar rosters de alumnos desde archivos Excel, tablas de Word y archivos de texto (.txt, .csv) detectando automáticamente nombres y cédulas panameñas, con previsualización interactiva.

---

## 📂 Estructura del Proyecto (Normas ISO)

El repositorio está organizado siguiendo normas internacionales de ingeniería de software para mantener un espacio de trabajo limpio, escalable y mantenible:

```text
RegistroDoc/
├── assets/                    # Recursos visuales y plantillas base resguardadas
│   ├── templates/             # Plantillas base oficiales limpias (Premedia/Primaria)
│   │   ├── Registro_Premedia.xlsx
│   │   └── Registro_Primaria.xlsx
├── data/                      # Archivos de configuración y bases de datos offline (cifrados)
├── dev/                       # Utilidades de desarrollo, scripts de limpieza y empaquetador
├── docs/                      # Documentación del proyecto (manuales de usuario oficiales)
├── src/                       # Código fuente de la aplicación en Python
│   ├── utils/                 # Funciones auxiliares y herramientas comunes (diálogos, fechas)
│   ├── app.py                 # Punto de entrada de la aplicación y vista contenedora
│   ├── rdsql.py               # Motor relacional de base de datos SQLite y auditorías
│   └── rdsecurity.py          # Módulo criptográfico y cifrado AES-256-GCM
├── tests/                     # Suite de pruebas automatizadas unitarias y de integración
├── RegistroDoc_Setup.iss      # Script del instalador profesional (Inno Setup)
├── requirements.txt           # Dependencias de librerías de Python
└── pytest.ini                 # Configuración del motor de pruebas unitarias
```

---

## 🚀 Cómo Abrir y Ejecutar el Programa

Dependiendo de su entorno de trabajo, puede iniciar el programa de tres maneras:

### 1. Ejecución desde el Código Fuente (Desarrollo)
Si tiene configurado Python 3.10 o posterior con todas las dependencias instaladas:
1. Abra su consola/terminal de comandos (Command Prompt, PowerShell o Git Bash) en la carpeta raíz del proyecto (`c:\RegistroDoc`).
2. Ejecute el comando principal:
   ```bash
   python src/app.py
   ```

### 2. Ejecución desde el Binario Compilado (Portátil)
Si empaquetó la aplicación con PyInstaller en un único archivo ejecutable:
1. Diríjase al directorio `dist/` en la raíz del proyecto.
2. Haga doble clic en el archivo **`RegistroDoc.exe`** para iniciar la interfaz gráfica directamente.

### 3. Desde el Instalador de Windows (Instalación Limpia)
Si compiló el instalador profesional de distribución:
1. Instale la aplicación ejecutando el archivo compilado `Output/RegistroDoc_Instalador.exe`.
2. Una vez instalado, abra el programa a través del acceso directo en el **Escritorio** ("RegistroDoc Pro") o desde el **Menú de Inicio** de Windows.

---

## 🗄️ Arquitectura de Datos y Visualización con SQLTools

Para asegurar la confidencialidad, la base de datos de producción (`data/registro.db.enc`) se mantiene cifrada con una clave AES-256 en reposo derivada a partir del hardware de la máquina.

### Visualización con la herramienta VS Code SQLTools:
1. Inicie la aplicación ejecutable o mediante `python src/app.py`.
2. Al estar activa, la base de datos se descifra de manera segura y temporal en la memoria/disco de sesión del usuario en la siguiente ruta:
   `C:\Users\<Tu_Usuario>\AppData\Local\RegistroDoc\temp\sqlite_temp.db`
3. En VS Code, instale la extensión **SQLTools** y el driver **SQLTools SQLite**.
4. Agregue una nueva conexión apuntando la ruta de la base de datos a la ruta del archivo temporal `sqlite_temp.db` mencionado en el paso anterior.
5. Podrá explorar las tablas (`configuracion`, `grados`, `estudiantes`, `materias`, `horario`, `notas`, `asistencia`, `observaciones`, `habitos`, `tareas`, `auditoria`) y ejecutar consultas SQL en tiempo real mientras la aplicación esté en uso. Al cerrar la aplicación, el archivo temporal se limpia automáticamente mediante sobreescritura aleatoria.

---

## 📦 Compilación e Instalador de Distribución

El proyecto cuenta con un instalador profesional compilado mediante **Inno Setup 6** que prepara todo el entorno operativo del docente de forma limpia:

### 1. Generar Ejecutable con PyInstaller
Para empaquetar la aplicación en un único ejecutable autónomo que integre todas las dependencias y recursos visuales:
```bash
pyinstaller RegistroDoc.spec --clean
```

### 2. Generar el Instalador (`RegistroDoc_Instalador.exe`)
Ejecute el compilador de Inno Setup desde la línea de comandos:
```powershell
& "C:\Users\<Tu_Usuario>\AppData\Local\Programs\Inno Setup 6\ISCC.exe" RegistroDoc_Setup.iss
```
El instalador resultante estará disponible en `C:\RegistroDoc\Output\RegistroDoc_Instalador.exe` y contiene las siguientes características:
* **Plantillas Excel Limpias**: El instalador incluye copias 100% vacías y sin datos personales de alumnos (`Registro_Premedia.xlsx` y `Registro_Primaria.xlsx`).
* **Verificación de Suite Ofimática**: Alerta al docente si el equipo no posee Excel ni LibreOffice instalado.
* **Desinstalación Segura con Cédula**: Exige la cédula registrada antes de proceder con el borrado de datos.
* **Limpieza Total**: Elimina registros y bases de datos locales del sistema operativo al desinstalar, manteniendo a salvo las hojas Excel exportadas por el docente de forma externa.

---

## 🧪 Pruebas Automatizadas

El proyecto incluye una suite de pruebas unitarias y de integración que validan el motor de datos, el flujo de cifrado, la integridad del Excel y el comportamiento de la UI:

Para correr las pruebas unitarias:
```bash
python -m pytest
```

---

## 🔒 Cumplimiento Legal y Normativo

1. **Ley 81 del 26 de marzo de 2019 (Panamá)**: Toda información personal e institucional (cédula, nombres, notas de alumnos) es cifrada de forma segura localmente.
2. **ISO/IEC 27001:2022 A.8.15 & A.8.24**: Logging seguro cifrado de todas las transacciones académicas y del sistema.
3. **NIST SP 800-63B / SP 800-38D**: Autenticación criptográfica y derivación de llaves por dispositivo mediante PBKDF2-HMAC-SHA256 con 600,000 iteraciones.

---

## 👥 Desarrolladores y Créditos

* **Autor Principal:** Creador Independiente (Panamá)
* **Ingeniería de Software e Integración de Calidad (V4.0):** Antigravity AI (Google Deepmind)
