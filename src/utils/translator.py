import os
import json
from rdsecurity import cargar_config_segura

_current_lang = None

def get_current_lang():
    global _current_lang
    if _current_lang is None:
        try:
            cfg = cargar_config_segura({})
            _current_lang = cfg.get("idioma", "es").lower()
        except Exception:
            _current_lang = "es"
    return _current_lang

def set_current_lang(lang):
    global _current_lang
    _current_lang = lang.lower()

# Dictionary of translations for all GUI elements
TRANSLATIONS = {
    "Español": "Spanish",
    "English": "English",
    # Menu lateral (app.py)
    "Dashboard": "Dashboard",
    "Inicio": "Home",
    "Estudiantes": "Students",
    "Asistencia": "Attendance",
    "Calificaciones": "Grades",
    "Reportes": "Reports",
    "Hábitos": "Habits",
    "Observaciones": "Observations",
    "Reuniones": "Meetings",
    "Configuraciones": "Settings",
    "Ayuda": "Help",
    "Salir": "Exit",
    "Menu": "Menu",
    "Notas y Asistencia": "Grades & Attendance",
    "Registro Completo": "Consolidated Record",
    "Reportes y Gráficos": "Reports & Charts",
    "Impresión": "Print",
    "Configuración": "Configuration",
    "Ayuda y Guía": "Help & Guide",
    "Minutas y Citaciones": "Minutes & Citations",
    "Inicio › Estudiantes": "Home › Students",
    "Inicio › Notas y Asistencia": "Home › Grades & Attendance",
    "Inicio › Observaciones": "Home › Observations",
    "Inicio › Reportes y Gráficos": "Home › Reports & Charts",
    "Inicio › Impresión": "Home › Print",
    "Inicio › Hábitos": "Home › Habits",
    "Inicio › Hábitos y Actitudes": "Home › Habits",
    "Inicio › Ayuda y Guía": "Home › Help & Guide",
    "Inicio › Ayuda": "Home › Help",
    "Inicio › Tareas Programadas": "Home › Tasks",
    "Inicio › Reuniones": "Home › Meetings",
    "Inicio › Minutas y Citaciones": "Home › Minutes",
    "Inicio › Configuración": "Home › Configuration",
    "Inicio › Registro Completo": "Home › Consolidated Record",
    
    # Panel de Control Maestro (sapp.py)
    "Panel de Control Maestro": "Master Control Panel",
    "Datos de Portada y Carátula": "Cover & Sheet Data",
    "Horarios": "Schedules",
    "Gestión de Grados": "Grades Management",
    "Gestión de Materias": "Subjects Management",
    "Modalidad Activa:": "Active Mode:",
    "Cambiar Modalidad": "Change Mode",
    "Tema Visual (Contraste):": "Visual Theme (Contrast):",
    "Oscuro": "Dark",
    "Claro (Alto Contraste)": "Light (High Contrast)",
    "Informacion Academica (Portada y Caratula)": "Academic Information (Cover & Sheet)",
    "Nombre Docente:": "Teacher Name:",
    "Cedula:": "ID (Cédula):",
    "No. Seguro Social:": "Social Security No.:",
    "No. Posicion:": "Position No.:",
    "Escuela:": "School:",
    "Region Escolar:": "School Region:",
    "Distrito:": "District:",
    "Corregimiento:": "Township (Corregimiento):",
    "Zona Escolar:": "School Zone:",
    "Ano Lectivo:": "School Year:",
    "Jornada:": "Shift / Session:",
    "Fecha Trimestre 1:": "Trimester 1 Date:",
    "Fecha Trimestre 2:": "Trimester 2 Date:",
    "Fecha Trimestre 3:": "Trimester 3 Date:",
    "Director(a):": "Principal:",
    "Subdirector(a):": "Vice Principal:",
    "Coordinador:": "Coordinator:",
    "Telefono:": "Phone:",
    "Correo:": "Email:",
    "Condicion/Titulo (Portada y Caratula):": "Condition/Title (Cover & Sheet):",
    "Grupos (Caratula):": "Groups (Cover):",
    "Logo de Escuela (Word):": "School Logo (Word):",
    "Seleccionar Logo": "Select Logo",
    "Ubicación de Exportación:": "Export Location:",
    "Seleccionar Carpeta": "Select Folder",
    "Al sincronizar, el programa inyectara estos datos en Portadas, Horarios y Asistencias.": "Upon syncing, the app will inject this data into Covers, Schedules, and Attendance.",
    "SINCRONIZAR Y SOBREESCRIBIR EXCEL": "SYNC AND OVERWRITE EXCEL",
    "🎓 CIERRE DE AÑO LECTIVO (ROLLOVER / ARCHIVO)": "🎓 SCHOOL YEAR ROLLOVER (ARCHIVE)",
    "Ubicación de Exportación (Opcional):": "Export Location (Optional):",
    "Idioma:": "Language:",
    
    # Asistente de Configuración (setup.py)
    "Perfil": "Profile",
    "Grados": "Grades",
    "Materias": "Subjects",
    "Horario": "Schedule",
    "Finalizar": "Finish",
    "Paso 1: Perfil Profesional": "Step 1: Professional Profile",
    "Ingresa los datos oficiales de tu libreta docente.": "Enter the official data for your digital teacher book.",
    "Nombre del Docente *:": "Teacher Name *:",
    "Cédula del Docente *:": "Teacher ID (Cédula) *:",
    "Nombre de la Escuela *:": "School Name *:",
    "Región Educativa (Panamá):": "Educational Region (Panama):",
    "Año Lectivo *:": "School Year *:",
    "Director(a) de la Escuela:": "School Principal:",
    "Subdirector(a) de la Escuela:": "School Vice Principal:",
    "Jornada Escolar:": "School Shift:",
    "Idioma de la Aplicación:": "Application Language:",
    "Ubicación de Exportación de Archivos *:": "File Export Location *:",
    "Explorar 📂": "Browse 📂",
    "Logo de la Escuela (Opcional)": "School Logo (Optional)",
    "🚫 Sin Escudo\nCargado": "🚫 No Logo\nLoaded",
    "📁 Cargar Logo": "📁 Load Logo",
    "Recomendado: Imagen PNG\ncon fondo transparente.": "Recommended: PNG image\nwith transparent background.",
    "⚠️ Error al\ncargar logo": "⚠️ Error loading\nlogo",
    "Paso 2: Grados y Secciones": "Step 2: Grades & Sections",
    "Selecciona la modalidad y los grupos que vas a atender.": "Select the modality and groups you will manage.",
    "Modalidad de Enseñanza:": "Teaching Modality:",
    "Premedia": "Pre-media",
    "Primaria": "Primary",
    "Habilitar": "Enable",
    "Grado / Nivel": "Grade / Level",
    "Sección / Letra (ej: A, B, C)": "Section / Letter (e.g. A, B, C)",
    "Paso 3: Materias": "Step 3: Subjects",
    "Agrega tus materias y asígnalas a los grados correspondientes.": "Add your subjects and assign them to the corresponding grades.",
    "Escribe materia (ej: Matemáticas)": "Type subject (e.g. Mathematics)",
    "➕ Agregar": "➕ Add",
    "Selecciona una materia a la izquierda": "Select a subject on the left",
    "Sin materias agregadas.": "No subjects added.",
    "Esta materia ya está en la lista.": "This subject is already in the list.",
    "Paso 4: Horario de Clases": "Step 4: Class Schedule",
    "Asigna tus materias a los bloques correspondientes de la semana (puedes dejar celdas vacías).": "Assign your subjects to the corresponding weekly blocks (you can leave cells empty).",
    "Hora / Bloque": "Hour / Block",
    "Libre": "Free / Open",
    "Debes agregar al menos una materia.": "You must add at least one subject.",
    "La materia '{mat}' no tiene ningún grado seleccionado.": "The subject '{mat}' has no grade selected.",
    "Paso 5: Resumen y Confirmación": "Step 5: Summary & Confirmation",
    "Verifica que los datos sean correctos antes de construir tu libreta digital.": "Verify the data is correct before building your digital book.",
    "👤 Docente: ": "👤 Teacher: ",
    "🪪 Cédula: ": "🪪 ID: ",
    "🏫 Escuela: ": "🏫 School: ",
    "🗓️ Año Lectivo: ": "🗓️ School Year: ",
    "📈 Nivel: ": "📈 Level: ",
    "🖼️ Logo Escolar: ": "🖼️ School Logo: ",
    "Cargado": "Loaded",
    "Por defecto": "Default",
    "👥 Grupos Registrados:\n": "👥 Registered Groups:\n",
    "  - Grado ": "  - Grade ",
    " Sección ": " Section ",
    "\n📚 Asignaturas:\n": "\n📚 Subjects:\n",
    " dictada en: ": " taught in: ",
    "\n📅 Bloques de Horario Asignados: ": "\n📅 Assigned Schedule Blocks: ",
    " períodos escolares.": " school periods.",
    "✅ FINALIZAR CONFIGURACIÓN": "✅ FINALIZE CONFIGURATION",
    "⏳ Generando Libreta...": "⏳ Generating Book...",
    "Siguiente ➡️": "Next ➡️",
    "⬅️ Atrás": "⬅️ Back",
    "Omitir y Usar Demo ❌": "Skip & Use Demo ❌",
    
    # Horarios (sapp.py)
    "⏱️ Calculadora de Tiempos": "⏱️ Time Calculator",
    "Hora Entrada:": "Start Time:",
    "Hora Salida:": "End Time:",
    "Mins Receso:": "Break Mins:",
    "Después del per.:": "After period:",
    "⚡ Generar Horas": "⚡ Generate Hours",
    "Per.": "Per.",
    "Horas": "Hours",
    "Lunes": "Monday",
    "Martes": "Tuesday",
    "Miércoles": "Wednesday",
    "Jueves": "Thursday",
    "Viernes": "Friday",
    "💾 GUARDAR HORARIO EN EXCEL": "💾 SAVE SCHEDULE TO EXCEL",
    "RECESO ESCOLAR": "SCHOOL BREAK",
    
    # Asistencia (fapp.py)
    "Asistencia de Estudiantes": "Student Attendance",
    "Justificaciones": "Excuses",
    "Justificar Inasistencia": "Justify Absence",
    
    # Calificaciones (happ.py / grapp.py)
    "Calificaciones": "Grades",
    "Graficos": "Charts",
    "Generar Reporte": "Generate Report",
    "Descargar PDF": "Download PDF",
    "Imprimir Reporte": "Print Report",
    
    # Observaciones (obsapp.py)
    "Observaciones": "Observations",
    "Expediente de Estudiante": "Student Record File",
    "Guardar Observación": "Save Observation",
    "Abrir Expediente": "Open Record File",

    # General GUI Terms
    "Grado:": "Grade:",
    "Trimestre:": "Trimester:",
    "Materia:": "Subject:",
    "Todos": "All",
    "Trimestre 1": "Trimester 1",
    "Trimestre 2": "Trimester 2",
    "Trimestre 3": "Trimester 3",
    "Todos los Trimestres": "All Trimesters",
    "Estudiante": "Student",
    "Buscar por nombre...": "Search by name...",
    "Buscar alumno para ver su rendimiento...": "Search student to view performance...",
    "Guardar Cambios": "Save Changes",
    "Cancelar": "Cancel",
    "Confirmar": "Confirm",
    "Error": "Error",
    "Información": "Information",
    "Atención": "Attention",
    "Aviso": "Warning",
    "Éxito": "Success",

    # Dashboard view (dashapp.py)
    "Panel de Control Principal": "Main Control Panel",
    "Estatus de Riesgo": "Risk Status",
    "Estable": "Stable",
    "En Riesgo": "At Risk",
    "Sin notas": "No grades",
    "No registrado": "Not registered",
    "Cuadro de Honor": "Honor Roll",
    "No aplica": "N/A",
    "Asistencia Indiv.": "Indiv. Attendance",
    "Total Alumnos": "Total Students",
    "Matriculados": "Enrolled",
    "Alumnos en Riesgo": "Students at Risk",
    "Trimestre actual": "Current trimester",
    "Asistencia": "Attendance",
    "Clase Actual:": "Current Class:",
    "Próxima clase:": "Next class:",
    "No hay más clases hoy": "No more classes today",
    "Fin de semana — ¡descanse!": "Weekend — rest!",
    "Tareas Pendientes": "Pending Tasks",
    "Distribución de Estudiantes": "Student Distribution",
    "Activos": "Active",
    "Retirados": "Withdrawn",
    "Músculo / Masculino": "Male",
    "Femenino": "Female",
    "Promedio de Asistencia Grupal": "Group Attendance Average",
    "Rendimiento Académico por Grado / Sección": "Academic Performance by Grade / Section",
    "Acciones de Impresión": "Print Actions",
    "Boletín": "Report Card",
    "Lista de Clase": "Class List",
    "Reporte Docente": "Teacher Report",
    "Reporte Dirección": "Principal Report",

    # Student list view (dapp.py)
    "Listado Oficial de Estudiantes": "Official Student Directory",
    "Listado Oficial de Estudiantes - RegistroDoc Pro": "Official Student Directory - RegistroDoc Pro",
    "Añadir Estudiante": "Add Student",
    "Registrar Retiro": "Register Withdrawal",
    "Reactivar": "Reactivate",
    "Carga Masiva (Roster)": "Mass Import (Roster)",
    "Filtros de Búsqueda": "Search Filters",
    "Todos los alumnos": "All students",
    "Solo activos": "Only active",
    "Solo retirados": "Only withdrawn",
    "Exportar a Excel": "Export to Excel",
    "Importar Lista": "Import List",
    "Guardar Nuevo Estudiante": "Save New Student",
    "Nombre Completo:": "Full Name:",
    "Cédula:": "ID (Cédula):",
    "Sexo:": "Gender:",
    "Masculino": "Male",
    "Femenino": "Female",
    "Cédula (Opcional):": "ID (Cédula - Optional):",
    "Cédula del Docente *:": "Teacher ID (Cédula) *:",
    "Motivo del Retiro:": "Reason for Withdrawal:",
    "Confirmar Retiro": "Confirm Withdrawal",

    # Attendance view (fapp.py)
    "Asistencia de Estudiantes": "Student Attendance",
    "Registro de Asistencia": "Attendance Log",
    "Fecha:": "Date:",
    "Pasar Lista": "Take Attendance",
    "Guardar Asistencia": "Save Attendance",
    "Todos Presentes": "All Present",
    "Todos Ausentes": "All Absent",
    "Justificar Inasistencia": "Justify Absence",
    "Justificaciones": "Excuses",
    "Justificación de Inasistencia": "Absence Excuse",
    "Tipo de Registro:": "Type of Record:",
    "Ausencia": "Absence",
    "Tardanza": "Tardy",
    "Excusa Justificada": "Excused Absence",
    "Fecha de la Inasistencia:": "Date of Absence:",
    "Guardar Justificación": "Save Excuse",

    # Grades view (eapp.py)
    "Registro de Calificaciones": "Grades Log",
    "Asignatura / Materia:": "Subject / Course:",
    "Tipo de Evaluación:": "Evaluation Type:",
    "Diaria": "Daily",
    "Apreciación": "Appreciation",
    "Examen": "Exam",
    "Guardar Calificaciones": "Save Grades",
    "Título/Tema de la Tarea:": "Task Title/Topic:",
    "Puntos Máximos:": "Max Points:",
    "Nueva Tarea": "New Task",
    "Eliminar Tarea": "Delete Task",
    "Calificaciones": "Grades",
    "Nota MEDUCA": "MEDUCA Grade",
    "Puntos Obtenidos": "Points Obtained",

    # Consolidation view (registrocompletoapp.py)
    "Registro General Consolidado": "Consolidated General Record",
    "Filtrar por Materia:": "Filter by Subject:",
    "Filtrar por Tipo de Nota:": "Filter by Grade Type:",
    "Exportar a Excel": "Export to Excel",
    "Actualizar Vista": "Refresh View",

    # Habits and attitudes view (habapp.py)
    "Indicador General de Hábitos y Actitudes": "General Habits & Attitudes Indicator",
    "Frecuencia:": "Frequency:",
    "Trimestral": "Quarterly",
    "Diario": "Daily",
    "Semanal": "Weekly",
    "Mensual": "Monthly",
    "Guardar Evaluaciones": "Save Evaluations",
    "Sugerir con IA": "Suggest with AI",
    "Autorellenar con IA": "Autofill with AI",
    "Satisfactorios (S)": "Satisfactory (S)",
    "Regulares (R)": "Regular (R)",
    "Insatisfactorios (X)": "Unsatisfactory (X)",

    # Reports and charts view (happ.py / grapp.py)
    "Reportes y Estadísticas": "Reports & Statistics",
    "Gráficos de Rendimiento": "Performance Charts",
    "Generar Reporte": "Generate Report",
    "Descargar PDF": "Download PDF",
    "Descargar Word": "Download Word",
    "Imprimir Reporte": "Print Report",
    "Rendimiento por Grupo": "Performance by Group",
    "Aprobados / Reprobados": "Passed / Failed",
    "Cuadro de Honor (Top 4)": "Honor Roll (Top 4)",
    "Asistencia Promedio": "Average Attendance",

    # Meetings and citations view (reunionesapp.py)
    "Generador de Actas de Reuniones": "Meeting Minutes Generator",
    "Reunión de Docentes": "Teachers Meeting",
    "Reunión con Padres": "Parent Meeting",
    "Datos de la Reunión": "Meeting Details",
    "Fecha:": "Date:",
    "Hora Inicio:": "Start Time:",
    "Hora Fin:": "End Time:",
    "Lugar:": "Location:",
    "Acta N°:": "Minutes No.:",
    "Asistentes (uno por línea):": "Attendees (one per line):",
    "Temas Tratados (uno por línea):": "Topics Discussed (one per line):",
    "Acuerdos y Compromisos (uno por línea):": "Agreements & Commitments (one per line):",
    "Observaciones:": "Observations:",
    "Observaciones adicionales...": "Additional observations...",
    "GENERAR ACTA DE REUNIÓN": "GENERATE MEETING MINUTES",
    "Datos de la Citación": "Citation Details",
    "Alumno:": "Student:",
    "Acudiente:": "Guardian:",
    "Parentesco:": "Relationship:",
    "Padre": "Father",
    "Madre": "Mother",
    "Abuelo/a": "Grandparent",
    "Tío/a": "Uncle/Aunt",
    "Tutor Legal": "Legal Guardian",
    "Otro": "Other",
    "Teléfono:": "Phone:",
    "Motivo de la reunión:": "Reason for Meeting:",
    "Descripción de la situación:": "Situation Description:",
    "Acuerdos con el padre/madre:": "Agreements with Parent:",
    "Compromisos del alumno:": "Student Commitments:",
    "GENERAR ACTA DE CITACIÓN": "GENERATE CITATION RECORD",

    # Print center (impp.py)
    "Centro de Impresión y Exportación": "Print & Export Center",
    "Seleccione qué reporte desea visualizar, imprimir o abrir directamente en Microsoft Excel.": "Select which report you want to view, print, or open directly in Microsoft Excel.",
    "Parámetros del Reporte": "Report Parameters",
    "Seleccione el Tipo de Reporte:": "Select Report Type:",
    "Seleccione el Grado:": "Select Grade:",
    "Seleccione la Asignatura:": "Select Subject:",
    "Seleccione el Periodo:": "Select Period:",
    "Abrir en Microsoft Excel": "Open in Microsoft Excel",
    "Enviar a Impresora": "Send to Printer",
    "Solo Gráficos": "Charts Only",

    # Help center (helpapp.py)
    "Centro de Ayuda RegistroDoc Pro": "Help Center - RegistroDoc Pro",
    "Guías Paso a Paso": "Step-by-Step Guides",
    "Preguntas Frecuentes": "FAQ",
    "Manual de Usuario": "User Manual",
    "Buscar en toda la ayuda...": "Search help...",

    # Tasks / Schedule (tareasapp.py)
    "Tareas y Evaluaciones Programadas": "Scheduled Tasks & Evaluations",
    "Programar Nueva Tarea": "Schedule New Task",
    "Título de la Tarea:": "Task Title:",
    "Fecha Límite (DD-MM-YYYY):": "Due Date (DD-MM-YYYY):",
    "PROGRAMAR TAREA": "SCHEDULE TASK",
    "Tareas Programadas": "Scheduled Tasks",
    " ({detalle_val})": " ({detalle_val})",
    "1. Seleccione el Grado Destino:": "1. Select Target Grade:",
    "1. Seleccione el Tipo de Reporte:": "1. Select Report Type:",
    "2. Cargue el archivo con el listado:": "2. Load file list:",
    "2. Seleccione el Grado:": "2. Select Grade:",
    "3. Seleccione la Asignatura:": "3. Select Subject:",
    "4. Seleccione el Periodo:": "4. Select Period:",
    "ACCIÓN": "ACTION",
    "APELLIDO, NOMBRE": "LAST NAME, FIRST NAME",
    "APELLIDO, Nombre (Requerido)": "LAST NAME, First Name (Required)",
    "Académico": "Academic",
    "Actualizando en fondo...": "Updating in background...",
    "Actualizando...": "Updating...",
    "Actualizar Consejero": "Update Advisor Teacher",
    "Agendar Tarea": "Schedule Task",
    "Alumno: {nombre_alumno}": "Student: {nombre_alumno}",
    "Anual": "Annual",
    "Archivo: {os.path.basename(ruta)}": "File: {os.path.basename(ruta)}",
    "Asistencia %": "Attendance %",
    "Asunto/Motivo:": "Subject/Reason:",
    "Aviso PDF": "PDF Notice",
    "Bajo rendimiento académico": "Low academic performance",
    "Borrando...": "Deleting...",
    "Buscando...": "Searching...",
    "Calificaciones exportadas y abiertas en Notepad:\n{res_path}": "Grades exported and opened in Notepad:\n{res_path}",
    "Cambio de Modalidad": "Change Modality",
    "Cargando...": "Loading...",
    "Categoría de la plantilla:": "Template category:",
    "Categoría:": "Category:",
    "Citación a Acudiente": "Parent/Guardian Citation",
    "Clase Actual: {materia}": "Current Class: {materia}",
    "Clonando...": "Cloning...",
    "Clonar formato de:": "Clone format from:",
    "Conducta": "Behavior",
    "Conducta inapropiada": "Inappropriate behavior",
    "Consejero actualizado.": "Advisor updated.",
    "Conversión de Puntos a Nota": "Conversion of Points to Grade",
    "Creando...": "Creating...",
    "Crear Expediente": "Create Record File",
    "Crear Grado": "Create Grade",
    "Crear Materia": "Create Subject",
    "CÉDULA": "ID / CÉDULA",
    "Cédula": "ID / Cédula",
    "Cédula (Editable)": "ID / Cédula (Editable)",
    "Cédula (Opcional)": "ID / Cédula (Optional)",
    "Cédula incorrecta. Intente de nuevo.": "Incorrect ID. Try again.",
    "Datos exportados a:\n{res_path}": "Data exported to:\n{res_path}",
    "Debe colocar una fecha.": "You must enter a date.",
    "Debes activar al menos un grado para continuar.": "You must activate at least one grade to continue.",
    "Descarga DOCX": "Download DOCX",
    "Descarga PDF": "Download PDF",
    "Descripción (Ej. Charla): (0/25)": "Description (e.g. Chat): (0/25)",
    "Descripción (Ej. Charla): ({len(texto)}/25)": "Description (e.g. Chat): ({len(texto)}/25)",
    "Descripción (opcional):": "Description (optional):",
    "Detalle de la nota:": "Grade details:",
    "Detalles adicionales...": "Additional details...",
    "Diaria / Parcial": "Daily / Quiz",
    "Documento generado en DOCX:\n{tmp_docx}\n\nNo se pudo convertir automáticamente a PDF (requiere MS Word en Windows).": "Document generated in DOCX:\n{tmp_docx}\n\nCould not automatically convert to PDF (requires MS Word on Windows).",
    "Documento generado en:\n{tmp_docx}\n\nNo se pudo convertir a PDF automáticamente (requiere MS Word en Windows). Puedes abrirlo y guardarlo como PDF desde Word.": "Document generated at:\n{tmp_docx}\n\nCould not automatically convert to PDF (requires MS Word on Windows). You can open and save it as PDF from Word.",
    "Documento generado en:\n{tmp_docx}\n\nPuedes abrirlo y guardarlo como PDF desde Word.": "Document generated at:\n{tmp_docx}\n\nYou can open and save it as PDF from Word.",
    "ESTADO": "STATUS",
    "ESTUDIANTE": "STUDENT",
    "Ej. Charla, Taller, etc.": "e.g. Talk, Workshop, etc.",
    "Ej: 15-06-2026": "e.g. 15-06-2026",
    "Ej: 4-123-456": "e.g. 4-123-456",
    "Ej: 4-750-1234": "e.g. 4-750-1234",
    "Ej: C.E.B.G. Quebrada de Guabo": "e.g. C.E.B.G. Quebrada de Guabo",
    "Ej: Conducta: Falta grave": "e.g. Behavior: Serious misconduct",
    "Ej: Examen de Matemáticas": "e.g. Math Exam",
    "Ej: María González": "e.g. Maria Gonzalez",
    "Ej: Prof. Antonio Pérez": "e.g. Prof. Antonio Perez",
    "Ej: Prof. Juan Gómez": "e.g. Prof. Juan Gomez",
    "Ej: Prof. María Martínez": "e.g. Prof. Maria Martinez",
    "Ej: Traslado a otra escuela": "e.g. Transfer to another school",
    "El Apellido y Nombre son obligatorios.": "Last Name and First Name are required.",
    "El Horario se actualizó.": "The Schedule was updated.",
    "El Registro Oficial de Excel ha sido exportado con éxito a:\n{res_path}": "The Official Excel Registry has been successfully exported to:\n{res_path}",
    "El expediente de {nombre_est} aún no se ha creado.\n\n¿Desea crearlo ahora?": "The record file for {nombre_est} has not been created yet.\n\nWould you like to create it now?",
    "Eliminar Grado": "Delete Grade",
    "Eliminar Materia": "Delete Subject",
    "English": "English",
    "Enviado a la cola de impresión.": "Sent to the print queue.",
    "Enviado a la impresora.": "Sent to printer.",
    "Error al cargar": "Error loading",
    "Error al intentar imprimir: {e}": "Error attempting to print: {e}",
    "Error de Excel": "Excel Error",
    "Error de guardado": "Save Error",
    "Error de impresión": "Print Error",
    "Error en el análisis de IA: {e}": "AI analysis error: {e}",
    "Error en la nota para {id_est}: {msg}": "Error in grade for {id_est}: {msg}",
    "Escala 1–5 por trimestre": "Scale 1–5 per trimester",
    "Escala Puntos": "Points Scale",
    "Escriba el detalle de la nota.": "Write the details of the note.",
    "Escriba el nombre de la nueva materia.": "Type the name of the new subject.",
    "Escriba el nombre del new consejero.": "Type the advisor's name.",
    "Escriba la justificación...": "Write the excuse...",
    "Escriba un nombre para la plantilla.": "Type a name for the template.",
    "Escriba un título para la tarea.": "Type a title for the task.",
    "Escriba una fecha límite.": "Type a due date.",
    "Escribe el nombre del grado.": "Type the grade name.",
    "Escribe una descripción.": "Write a description.",
    "Español": "Spanish",
    "Estado": "Status",
    "Evaluaciones registradas y guardadas correctamente.": "Evaluations recorded and saved successfully.",
    "Exig:": "Exigency:",
    "Expediente Actualizado": "Record Updated",
    "Expediente y Observaciones": "Record & Observations",
    "Exportar Calificaciones": "Export Grades",
    "Exportar Excel": "Export Excel",
    "Exportar PDF": "Export PDF",
    "Exportar TXT": "Export TXT",
    "Exportar Word": "Export Word",
    "FECHA": "DATE",
    "Falta el archivo: {archivo_nuevo}": "File missing: {archivo_nuevo}",
    "Falta librería python-docx o utils de gráficos.": "Missing python-docx or graphics utils library.",
    "Falta librería python-docx.": "Missing python-docx library.",
    "Faltan datos": "Missing data",
    "Fecha (MM-DD):": "Date (MM-DD):",
    "Fecha de guardado (MM-DD):": "Save Date (MM-DD):",
    "Fecha de guardado:": "Save Date:",
    "Fecha del reporte (DD-MM-AAAA):": "Report Date (DD-MM-YYYY):",
    "Fecha límite (DD-MM-YYYY):": "Due Date (DD-MM-YYYY):",
    "GRADO": "GRADE",
    "General": "General",
    "Generar Word": "Generate Word",
    "Generar e Imprimir": "Generate & Print",
    "Gestión de Estudiantes": "Student Management",
    "Gestión de Notas": "Grades Management",
    "Grado creado.": "Grade created.",
    "Grado eliminado.": "Grade deleted.",
    "Grados Activos": "Active Grades",
    "Grados Disponibles:": "Available Grades:",
    "Grados/Grupos:": "Grades/Groups:",
    "Grupo Consejero 8° B": "Advisor Group 8th B",
    "Guardando...": "Saving...",
    "Guardar Nuevo": "Save New",
    "Guardar Plantilla": "Save Template",
    "Hubo un problema al guardar los cambios.": "There was a problem saving changes.",
    "Importar": "Import",
    "Importe su lista de alumnos directamente desde un archivo de Excel (.xlsx), Word (.docx) o Texto (.txt). El sistema detectará automáticamente los nombres y cédulas panameñas.": "Import your student list directly from an Excel (.xlsx), Word (.docx), or Text (.txt) file. The system will automatically detect names and Panamanian IDs (cédulas).",
    "Imprimir": "Print",
    "Imprimir Gráficos": "Print Charts",
    "Informe de calificaciones generado exitosamente:\n{ruta}": "Grades report generated successfully:\n{ruta}",
    "Informe de calificaciones individual generado:\n{ruta}": "Individual grades report generated successfully:\n{ruta}",
    "Ingrese su Cédula": "Enter your ID (Cédula)",
    "JUSTIFICACIÓN DE AUSENCIA/TARDANZA": "ABSENCE/TARDINESS EXCUSE",
    "Justificación Faltante": "Missing Excuse",
    "La fecha debe estar en formato DD-MM-YYYY (Ej: 15-06-2026).": "Date must be in DD-MM-YYYY format (e.g. 15-06-2026).",
    "La lista de {grado} está vacía. ¡Agrega un alumno arriba ☝️!": "The list for {grado} is empty. Add a student above ☝️!",
    "La nota debe estar entre 1.0 y 5.0.": "Grade must be between 1.0 and 5.0.",
    "Leyenda: P(.) A(-) T(T)": "Legend: P(.) A(-) T(T)",
    "Libreta actualizada con éxito.": "Digital book updated successfully.",
    "Los puntos deben estar entre 0 y {pts_max}.": "Points must be between 0 and {pts_max}.",
    "MOTIVO": "REASON",
    "Manual no encontrado.": "User manual not found.",
    "Materia a calificar:": "Subject to grade:",
    "Materia eliminada.": "Subject deleted.",
    "Matutina": "Morning",
    "Mención Honorífica": "Honorary Mention",
    "Motivo del retiro *": "Reason for withdrawal *",
    "Máx:": "Max:",
    "N/A": "N/A",
    "NOMBRE": "NAME",
    "Ninguna": "None",
    "Ningún archivo seleccionado (Rechazado).": "No file selected (Rejected).",
    "Ningún archivo seleccionado.": "No file selected.",
    "No cargada": "Not loaded",
    "No ha seleccionado ningún estudiante para importar.": "No student selected for import.",
    "No hay calificaciones registradas para esta materia y trimestre.": "No grades recorded for this subject and trimester.",
    "No hay coincidencias.": "No matches.",
    "No hay datos de evaluación para guardar.": "No evaluation data to save.",
    "No hay datos de resumen para mostrar.": "No summary data to display.",
    "No hay estudiantes en este grado.": "No students in this grade.",
    "No hay estudiantes inscritos en este grado.": "No students enrolled in this grade.",
    "No hay estudiantes registrados en este grado.": "No students registered in this grade.",
    "No hay estudiantes retirados en este grado.": "No withdrawn students in this grade.",
    "No hay grados": "No grades",
    "No hay grados activos": "No active grades",
    "No hay materias": "No subjects",
    "No hay notas para guardar.": "No grades to save.",
    "No hay registros de asistencia para este grado y trimestre.": "No attendance records for this grade and trimester.",
    "No hay registros de hábitos para este grado y trimestre.": "No habits records for this grade and trimester.",
    "No hay tareas programadas\npara esta materia.": "No tasks scheduled\nfor this subject.",
    "No hay un grado válido seleccionado.": "No valid grade selected.",
    "No se encontraron datos para exportar.": "No data found to export.",
    "No se encontraron estudiantes para el grado seleccionado.": "No students found for the selected grade.",
    "No se encontró el archivo de datos original para respaldar.": "Original data file not found to backup.",
    "No se encontró la hoja Horario.": "Schedule sheet not found.",
    "No se encontró la hoja correspondiente en el archivo Excel.": "Corresponding sheet not found in Excel file.",
    "No se encontró la plantilla de Excel limpia en:\n{plantilla}": "Clean Excel template not found at:\n{plantilla}",
    "No se pudo abrir: {e}": "Could not open: {e}",
    "No se pudo actualizar.": "Could not update.",
    "No se pudo copiar la imagen: {e}": "Could not copy image: {e}",
    "No se pudo crear el expediente: {e}": "Could not create record file: {e}",
    "No se pudo crear el respaldo: {e}": "Could not create backup: {e}",
    "No se pudo crear o verificar la ruta de exportación: {e}": "Could not create or verify export path: {e}",
    "No se pudo determinar la hoja de resumen para {grado}.": "Could not determine summary sheet for {grado}.",
    "No se pudo generar el acta. Verifique python-docx.": "Could not generate minutes. Verify python-docx.",
    "No se pudo generar. Verifique python-docx.": "Could not generate. Verify python-docx.",
    "No se pudo guardar el expediente: {extra}": "Could not save record file: {extra}",
    "No se pudo guardar la evaluación: {error_msg}": "Could not save evaluation: {error_msg}",
    "No se pudo guardar la plantilla: {e}": "Could not save template: {e}",
    "No se pudo guardar o abrir el archivo: {e}": "Could not save or open file: {e}",
    "No se pudo imprimir: {e}": "Could not print: {e}",
    "No se pudo leer el archivo: {e}": "Could not read file: {e}",
    "No se pudo reactivar.\n{resultado}": "Could not reactivate.\n{resultado}",
    "No se pudo retirar al estudiante.\n{resultado}": "Could not withdraw student.\n{resultado}",
    "Nocturna": "Night / Evening",
    "Nombre": "Name",
    "Nombre de Nueva Materia": "New Subject Name",
    "Nombre del Estudiante": "Student Name",
    "Nombre del acudiente (opcional)": "Parent/Guardian name (optional)",
    "Nombre del grado": "Grade name",
    "Nombre del padre/madre": "Father/Mother name",
    "Nombre/Título de la plantilla:": "Template Name/Title:",
    "Nota": "Grade",
    "Nota Final": "Final Grade",
    "Nuevo Nombre del Consejero": "New Advisor Name",
    "N°": "No.",
    "Observación guardada.\n\n¿Desea abrir el expediente de {nombre_est} ahora mismo?": "Observation saved.\n\nWould you like to open the record file of {nombre_est} right now?",
    "Ocurrió un error al exportar la información al archivo de Excel.": "An error occurred while exporting information to Excel file.",
    "Otro tip ↻": "Another tip ↻",
    "P  A  N  A  M  Á": "P  A  N  A  M  A",
    "PDF generado en:\n{pdf_path}": "PDF generated at:\n{pdf_path}",
    "Para reanudar su sesión de forma segura, ingrese su Cédula:": "To safely resume your session, please enter your ID (Cédula):",
    "Parcial": "Quiz / Partial",
    "Peligro": "Danger",
    "Período/Fecha:": "Period/Date:",
    "Por favor completa todos los campos requeridos (*).": "Please complete all required fields (*).",
    "Procesando en fondo...": "Processing in background...",
    "Prof. Consejero": "Advisor Teacher",
    "Programe trabajos con anticipación y reciba recordatorios": "Schedule assignments in advance and receive reminders",
    "Promedio": "Average",
    "Promedio final por sección": "Final average per section",
    "Próxima clase: {prox_mat}": "Next class: {prox_mat}",
    "Pts": "Pts",
    "Puntos": "Points",
    "Redacte la observación:": "Write the observation:",
    "Redacte una observación antes de guardarla como plantilla.": "Write an observation before saving it as a template.",
    "Registro Diario": "Daily Register",
    "RegistroDoc Pro": "RegistroDoc Pro",
    "RegistroDoc Pro Setup": "RegistroDoc Pro Setup",
    "RegistroDoc Pro — Control Académico": "RegistroDoc Pro — Academic Control",
    "Rellenar": "Fill",
    "Relleno rápido:": "Quick Fill:",
    "Reportes generados en Word:\n{tmp_docx}": "Word reports generated at:\n{tmp_docx}",
    "Respaldo creado con éxito en:\n{ruta_destino}": "Backup successfully created at:\n{ruta_destino}",
    "SEXO": "GENDER",
    "Se enviará a imprimir la hoja correspondiente al reporte seleccionado.": "The sheet corresponding to the selected report will be sent to print.",
    "Se generó la plantilla en blanco con éxito para el grado {grado}.\n\nSe abrirá en Excel para que la pueda imprimir.": "Blank template successfully generated for grade {grado}.\n\nIt will open in Excel so you can print it.",
    "Selecciona una materia\npara configurarle sus grados.": "Select a subject\nto configure its grades.",
    "Seleccionar Todos": "Select All",
    "Seleccionar plantilla...": "Select template...",
    "Seleccionar todos": "Select all",
    "Seleccione Grado:": "Select Grade:",
    "Seleccione Trimestre:": "Select Trimester:",
    "Seleccione al menos un Grado/Grupo.": "Select at least one Grade/Group.",
    "Seleccione arriba": "Select above",
    "Seleccione el Estudiante:": "Select the Student:",
    "Seleccione fecha existente:": "Select existing date:",
    "Seleccione grado": "Select grade",
    "Seleccione la nota a editar:": "Select the grade to edit:",
    "Seleccione parámetros arriba...": "Select parameters above...",
    "Seleccione un estudiante en la lista 👈": "Select a student from the list 👈",
    "Seleccione un grado destino válido.": "Select a valid target grade.",
    "Sin grados": "No grades",
    "Sin registros (4.5 - 5.0)": "No records (4.5 - 5.0)",
    "Solo si falta, tarde o excusa": "Only if absent, tardy, or excused",
    "Tendencia de Rendimiento Académico  (X/Y)": "Academic Performance Trend (X/Y)",
    "Tipo de evaluación:": "Evaluation Type:",
    "Tipo:": "Type:",
    "Todo lo que necesita saber, sin internet, en un solo lugar.": "Everything you need to know, offline, in one place.",
    "Todos los trimestres": "All trimesters",
    "Trimestre Completo": "Full Trimester",
    "Trimestre a corregir:": "Trimester to correct:",
    "Título de la tarea:": "Task Title:",
    "Use formato HH:MM (Ej: 07:00) para las horas.": "Use HH:MM format (e.g. 07:00) for hours.",
    "Vespertina": "Afternoon",
    "fecha_retiro": "withdrawal_date",
    "name": "name",
    "v.Prov.22:6": "v.Prov.22:6",
    "x": "x",
    "{count} alumno{'s' if count != 1 else ''} ({pct*100:.0f}%)": "{count} student{'s' if count != 1 else ''} ({pct*100:.0f}%)",
    "{count} evaluaciones ({pct:.1f}%)": "{count} evaluations ({pct:.1f}%)",
    "{dia_nombre} — {prox_rango} — Ahora: {hora_str}": "{dia_nombre} — {prox_rango} — Now: {hora_str}",
    "{i+1}.": "{i+1}.",
    "{len(guia[\"pasos\"])} pasos": "{len(guia[\"pasos\"])} steps",
    "{p} pts": "{p} pts",
    "{t['tipo']} | Lim: {t['fecha_limite']}": "{t['tipo']} | Due: {t['fecha_limite']}",
    "{t['tipo']} | Lim: {t['fecha_limite']}{dias_txt}": "{t['tipo']} | Due: {t['fecha_limite']}{dias_txt}",
    "¡Configuración inicial completada con éxito!\nLa libreta digital escolar ha sido generada.": "Initial configuration completed successfully!\nThe digital teacher book has been generated.",
    "«{texto}»": "«{texto}»",
    "¿Desea eliminar esta tarea programada?": "Do you want to delete this scheduled task?",
    "¿Eliminar DEFINITIVAMENTE {g}?": "Permanently delete {g}?",
    "¿Eliminar esta tarea?": "Delete this task?",
    "¿Eliminar {m} de {g}?": "Delete {m} from {g}?",
    "¿Para qué grados es {mat}?": "Which grades is {mat} for?",
    "— {referencia}": "— {referencia}",
    "→ {respuesta}": "→ {respuesta}",
    "⌨️ Atajos de Teclado — Trabaje más rápido": "⌨️ Keyboard Shortcuts — Work faster",
    "⏮️ Trimestres Anteriores:": "⏮️ Previous Trimesters:",
    "⏳ ACTUALIZANDO...": "⏳ UPDATING...",
    "⏳ ANALIZANDO...": "⏳ ANALYZING...",
    "⏳ GUARDANDO...": "⏳ SAVING...",
    "⏳ PENDIENTES": "⏳ PENDING",
    "⏳ Pendientes": "⏳ Pending",
    "⏳ Sincronizando...": "⏳ Syncing...",
    "⚙️ OPTIMIZANDO DISEÑO...": "⚙️ OPTIMIZING DESIGN...",
    "⚙️ Parámetros del Reporte": "⚙️ Report Parameters",
    "⚠️ No se detectaron estudiantes en el archivo. Verifique el formato e inténtelo de nuevo.": "⚠️ No students detected in the file. Verify format and try again.",
    "⚠️ Ocurrió un error o no hay evaluaciones de hábitos registradas.": "⚠️ An error occurred or no habit evaluations are registered.",
    "⚠️ Ocurrió un error o no hay hoja de asistencia válida en el Excel para este grupo.": "⚠️ An error occurred or no valid attendance sheet exists in Excel for this group.",
    "⚠️ Ocurrió un error o no hay hoja de notas válida en el Excel para este grupo.": "⚠️ An error occurred or no valid grades sheet exists in Excel for this group.",
    "⚡ Acciones Rápidas (Un Clic)": "⚡ Quick Actions (One Click)",
    "✅ Acuerdos y Compromisos (uno por línea):": "✅ Agreements & Commitments (one per line):",
    "✅ COMPLETADAS": "✅ COMPLETED",
    "✅ Completadas": "✅ Completed",
    "✅ Confirmar Retiro": "✅ Confirm Withdrawal",
    "✅ REINTENTAR": "✅ RETRY",
    "✅ Reactivar": "✅ Reactivate",
    "✅ Todo en orden. Ningún estudiante se encuentra en riesgo académico actualmente.": "✅ All in order. No students are currently at academic risk.",
    "✅ Todos Presentes": "✅ All Present",
    "✉️ Nota para Acudiente": "✉️ Note for Parent/Guardian",
    "✏️ Editar Tarea": "✏️ Edit Task",
    "✓ Cierre Exitoso": "✓ Successful Closure",
    "✓ Enviado a Impresora": "✓ Sent to Printer",
    "✓ Informe Generado": "✓ Report Generated",
    "✓ Lista Generada": "✓ List Generated",
    "✓ Éxito": "✓ Success",
    "✨ SINCRONIZAR Y SOBREESCRIBIR EXCEL": "✨ SYNC AND OVERWRITE EXCEL",
    "❌ CANCELAR": "❌ CANCEL",
    "❌ Estudiante no registrado": "❌ Student not registered",
    "❌ Todos Ausentes": "❌ All Absent",
    "❓ {pregunta}": "❓ {pregunta}",
    "➕ Agregar Alumno:": "➕ Add Student:",
    "➕ Agregar Materia a un Grado": "➕ Add Subject to a Grade",
    "➕ Agregar Nuevo Grado o Grupo (Ej. 10° o 8°B)": "➕ Add New Grade or Group (e.g. 10th or 8thB)",
    "➕ Nueva Tarea": "➕ New Task",
    "➕ Programar Tarea": "➕ Schedule Task",
    "⬇️ Exportar a PDF": "⬇️ Export to PDF",
    "⬇️ Exportar a Word": "⬇️ Export to Word",
    "⬇️ Gráficos a Word": "⬇️ Charts to Word",
    "👥 Asistentes (uno por línea):": "👥 Attendees (one per line):",
    "👥 Distribución de Estudiantes": "👥 Student Distribution",
    "👨‍🎓 Estudiante: {estudiante['nombre']}": "👨‍🎓 Student: {estudiante['nombre']}",
    "👪 Datos de la Citación": "👪 Citation Details",
    "💡 Marque todos presentes y corrija solo las excepciones": "💡 Mark all present and correct only exceptions",
    "💾 GUARDAR ASISTENCIA": "💾 SAVE ATTENDANCE",
    "💾 GUARDAR CAMBIOS": "💾 SAVE CHANGES",
    "💾 GUARDAR EVALUACIONES": "💾 SAVE EVALUATIONS",
    "💾 GUARDAR MODIFICACIONES DE LA LISTA": "💾 SAVE LIST MODIFICATIONS",
    "💾 GUARDAR NUEVA": "💾 SAVE NEW",
    "💾 Guardar plantilla": "💾 Save template",
    "💾 PROGRAMAR TAREA": "💾 SCHEDULE TASK",
    "💾 Respaldo": "💾 Backup",
    "📁 Seleccionar Archivo (Excel / Word / TXT)": "📁 Select File (Excel / Word / TXT)",
    "📂 Abrir Expediente": "📂 Open Record File",
    "📂 Abrir en Microsoft Excel": "📂 Open in Microsoft Excel",
    "📄 Abrir / Imprimir": "📄 Open / Print",
    "📄 Boletín": "📄 Report Card",
    "📄 GENERAR ACTA DE CITACIÓN": "📄 GENERATE CITATION RECORD",
    "📄 GENERAR ACTA DE REUNIÓN": "📄 GENERATE MEETING MINUTES",
    "📄 Informe del Estudiante": "📄 Student Report",
    "📄 PDF": "📄 PDF",
    "📅 Asistencia": "📅 Attendance",
    "📅 PROGRAMAR COMO TAREA": "📅 SCHEDULE AS TASK",
    "📅 Tareas Programadas": "📅 Scheduled Tasks",
    "📊 Excel": "📊 Excel",
    "📋 Generador de Actas de Reuniones": "📋 Meeting Minutes Generator",
    "📋 Registro General Consolidado": "📋 Consolidated General Record",
    "📋 Tabla": "📋 Table",
    "📋 Tareas Pendientes ({len(pendientes)})": "📋 Pending Tasks ({len(pendientes)})",
    "📋 Tareas y Evaluaciones Programadas": "📋 Scheduled Tasks & Evaluations",
    "📌 Temas Tratados (uno por línea):": "📌 Topics Discussed (one per line):",
    "📑 Ir a sección:": "📑 Go to section:",
    "📖 Manual de Usuario Oficial — v5.0 (SQL-First)": "📖 Official User Manual — v5.0 (SQL-First)",
    "📘 Términos y Conceptos del Sistema MEDUCA": "📘 Terms & Concepts of MEDUCA System",
    "📚 Centro de Ayuda RegistroDoc Pro": "📚 Help Center - RegistroDoc Pro",
    "📝 Calificar": "📝 Grade / Evaluate",
    "📝 Datos de la Reunión": "📝 Meeting Details",
    "📝 GUARDAR EN EXPEDIENTE OFICIAL": "📝 SAVE TO OFFICIAL RECORD FILE",
    "📝 Notas": "📝 Grades",
    "📝 Observaciones:": "📝 Observations:",
    "📝 TXT": "📝 TXT",
    "📝 Word": "📝 Word",
    "📥 Carga Masiva de Roster de Estudiantes": "📥 Mass Import of Student Roster",
    "📥 Confirmar Importación": "📥 Confirm Import",
    "📭 No hay tareas programadas.\nUse el formulario para crear una.": "📭 No tasks scheduled.\nUse the form to create one.",
    "🔄 ACTUALIZAR EXCEL": "🔄 UPDATE EXCEL",
    "🔄 Actualizar Profesor Consejero": "🔄 Update Advisor Teacher",
    "🔄 Cargando asistencia desde SQL...": "🔄 Loading attendance from SQL...",
    "🔄 Cargando hábitos desde SQL...": "🔄 Loading habits from SQL...",
    "🔄 Cargando registros de notas desde SQL...": "🔄 Loading grade records from SQL...",
    "🔄 Refrescar": "🔄 Refresh",
    "🔍 Buscar en toda la ayuda...": "🔍 Search help...",
    "🔍 CARGAR A LA LISTA": "🔍 LOAD TO LIST",
    "🔍 Observar": "🔍 Observe",
    "🔒 APLICACIÓN BLOQUEADA POR INACTIVIDAD": "🔒 APPLICATION LOCKED DUE TO INACTIVITY",
    "🔓 Desbloquear": "🔓 Unlock",
    "🔔 Estado y Alertas de Alumnos": "🔔 Student Status & Alerts",
    "🕐 Cargando...": "🕐 Loading...",
    "🕐 Próx: {prox_mat} ({prox_rango})": "🕐 Next: {prox_mat} ({prox_rango})",
    "🕐 {msg}": "🕐 {msg}",
    "🖨️ Centro de Impresión y Exportación": "🖨️ Print & Export Center",
    "🖨️ Enviar a Impresora": "🖨️ Send to Printer",
    "🖨️ Imprimir Reporte": "🖨️ Print Report",
    "🖨️ Lista de Clase": "🖨️ Class List",
    "🖨️ Solo Gráficos": "🖨️ Charts Only",
    "🗑️ Eliminar Grado Completo": "🗑️ Delete Complete Grade",
    "🗑️ Eliminar Materia": "🗑️ Delete Subject",
    "🚀 Acciones de Impresión": "🚀 Print Actions",
    "🚪 Registrar Retiro": "🚪 Register Withdrawal",
    "🚪 Retirar": "🚪 Withdraw",
    "🟢 Clase: {materia} ({rango})": "🟢 Class: {materia} ({rango})",
    "🧠 Autorellenar con IA": "🧠 Autofill with AI",
    "🧠 Hábitos y Actitudes Registrados — {nombre_est} ({grado_est})": "🧠 Registered Habits & Attitudes — {nombre_est} ({grado_est})",
    "🧠 Indicador General de Hábitos y Actitudes (Grupo Consejero 8° B)": "🧠 General Habits & Attitudes Indicator (Advisor Group 8th B)",
    "🧠 Sugerir con IA": "🧠 Suggest with AI",
}

# Reverse lookup dictionary to return Spanish keys from English selections
REVERSE_TRANSLATIONS = {v: k for k, v in TRANSLATIONS.items()}

def tr(text):
    if not isinstance(text, str):
        return text
    lang = get_current_lang()
    if lang != "en":
        return text
    
    # Try direct match
    if text in TRANSLATIONS:
        return TRANSLATIONS[text]
    
    # Try dynamic prefix / suffix patterns
    if text.startswith("Clase Actual: "):
        return "Current Class: " + text[len("Clase Actual: "):]
    if text.startswith("Próxima clase: "):
        return "Next class: " + text[len("Próxima clase: "):]
    if text.startswith("✓ Tema cambiado a "):
        return "✓ Theme changed to " + text[len("✓ Tema cambiado a "):]
    if text.startswith("✓ Idioma cambiado a "):
        return "✓ Language changed to " + text[len("✓ Idioma cambiado a "):]
    if text.startswith("✓ Observación agregada al expediente de "):
        return "✓ Observation added to student record of " + text[len("✓ Observación agregada al expediente de "):]
    if text.startswith("Promedio: "):
        return "Average: " + text[len("Promedio: "):]
    if text.startswith("Tareas Pendientes (") and text.endswith(")"):
        num = text[len("Tareas Pendientes ("):-1]
        return f"Pending Tasks ({num})"
    if text.startswith("🔔 En ") and text.endswith(" días"):
        days = text[len("🔔 En "):-len(" días")]
        return f"🔔 In {days} days"
    if text.startswith("📅 En ") and text.endswith(" días"):
        days = text[len("📅 En "):-len(" días")]
        return f"📅 In {days} days"
    if "Asistencia perfecta" in text:
        return text.replace("Asistencia perfecta", "Perfect Attendance")
    if "Falta python-docx" in text:
        return text.replace("Falta python-docx", "Missing python-docx")
    if "Expedientes actualizados" in text:
        return text.replace("Expedientes actualizados", "Student records updated")
    
    return text

def untr(text):
    if not isinstance(text, str):
        return text
    lang = get_current_lang()
    if lang != "en":
        return text
    
    # Try direct reverse translation
    if text in REVERSE_TRANSLATIONS:
        return REVERSE_TRANSLATIONS[text]
        
    # Revert dynamic prefixes
    if text.startswith("Current Class: "):
        return "Clase Actual: " + text[len("Current Class: "):]
    if text.startswith("Next class: "):
        return "Próxima clase: " + text[len("Next class: "):]
    if text.startswith("Average: "):
        return "Promedio: " + text[len("Average: "):]
        
    return text


# --- Automatic GUI Patching (Monkeypatching) ---
try:
    import customtkinter as ctk
    import tkinter as tk
    import tkinter.messagebox as msgbox

    def tr_value(val):
        if isinstance(val, str):
            return tr(val)
        elif isinstance(val, list):
            return [tr(x) if isinstance(x, str) else x for x in val]
        return val

    # Widgets to patch
    widgets_to_patch = [
        (ctk.CTkLabel, ["text"]),
        (ctk.CTkButton, ["text"]),
        (ctk.CTkCheckBox, ["text"]),
        (ctk.CTkRadioButton, ["text"]),
        (ctk.CTkEntry, ["placeholder_text"]),
        (ctk.CTkOptionMenu, ["values"]),
        (ctk.CTkComboBox, ["values"]),
        (ctk.CTkSegmentedButton, ["values"]),
    ]

    for cls, keys in widgets_to_patch:
        # Patch __init__
        orig_init = cls.__init__
        def make_init(orig=orig_init, ks=keys):
            def new_init(self, *args, **kwargs):
                for k in ks:
                    if k in kwargs:
                        kwargs[k] = tr_value(kwargs[k])
                orig(self, *args, **kwargs)
            return new_init
        cls.__init__ = make_init()

        # Patch configure
        orig_config = cls.configure
        def make_config(orig=orig_config, ks=keys):
            def new_config(self, *args, **kwargs):
                for k in ks:
                    if k in kwargs:
                        kwargs[k] = tr_value(kwargs[k])
                orig(self, *args, **kwargs)
            return new_config
        cls.configure = make_config()

    # Patch CTkTabview.add
    orig_tab_add = ctk.CTkTabview.add
    def new_tab_add(self, name, *args, **kwargs):
        translated_name = tr(name)
        return orig_tab_add(self, translated_name, *args, **kwargs)
    ctk.CTkTabview.add = new_tab_add

    # Patch CTkOptionMenu get / set
    orig_option_get = ctk.CTkOptionMenu.get
    def new_option_get(self):
        val = orig_option_get(self)
        return untr(val)
    ctk.CTkOptionMenu.get = new_option_get

    orig_option_set = ctk.CTkOptionMenu.set
    def new_option_set(self, value):
        orig_option_set(self, tr(value))
    ctk.CTkOptionMenu.set = new_option_set

    # Patch CTkComboBox get / set
    orig_combo_get = ctk.CTkComboBox.get
    def new_combo_get(self):
        val = orig_combo_get(self)
        return untr(val)
    ctk.CTkComboBox.get = new_combo_get

    orig_combo_set = ctk.CTkComboBox.set
    def new_combo_set(self, value):
        orig_combo_set(self, tr(value))
    ctk.CTkComboBox.set = new_combo_set

    # Patch tkinter messagebox functions
    for name in ["showinfo", "showwarning", "showerror", "askyesno", "askokcancel", "askquestion", "askretrycancel"]:
        orig_func = getattr(msgbox, name, None)
        if orig_func:
            def make_new_func(orig=orig_func):
                def new_func(title, message, **kwargs):
                    return orig(tr(title), tr(message), **kwargs)
                return new_func
            setattr(msgbox, name, make_new_func())

except Exception as e:
    # Fail-safe print if run without graphical libraries
    import sys
    print(f"[*] Translation GUI patching skipped: {e}", file=sys.stderr)
