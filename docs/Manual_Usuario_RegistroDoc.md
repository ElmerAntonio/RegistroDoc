## ══════════════════════════════════════════════════════════════════
##       MANUAL DE USUARIO — RegistroDoc Pro v5.0 (2026)
##       Guia Completa para el Docente Panameno
##       Ministerio de Educacion (MEDUCA) — Panama
## ══════════════════════════════════════════════════════════════════

Bienvenido a RegistroDoc Pro v5.0. Este programa fue disenado para
docentes panamenos que trabajan en areas sin internet. Funciona 100%
sin conexion y le ayudara a llevar notas, asistencia, conducta y
habitos de forma rapida, segura y profesional.

NOVEDADES v5.0:
  - Arquitectura SQL-First: todos los datos se leen desde SQLite
  - Numeracion visual 1,2,3... en todas las listas de estudiantes
  - Navegacion instantanea sin recarga entre modulos
  - 90 pruebas automatizadas garantizan integridad del sistema

INDICE:
  1.  Como Iniciar el Programa
  2.  Pantalla de Inicio (Dashboard)
  3.  Modulo de Estudiantes
  4.  Modulo de Calificaciones (Notas)
  5.  Modulo de Asistencia
  6.  Modulo de Observaciones (Expediente)
  7.  Modulo de Habitos y Aptitudes
  8.  Modulo de Tareas Programadas
  9.  Modulo de Reportes y Graficos
  10. Modulo de Registro Completo (Consolidado)
  11. Modulo de Impresion
  12. Modulo de Reuniones y Actas
  13. Configuracion del Sistema
  14. Atajos de Teclado
  15. Copias de Seguridad (Respaldos)
  16. Exportar a Excel
  17. Solucion de Problemas Comunes
  18. Glosario de Terminos MEDUCA
  19. Estructura de Base de Datos (SQLTools)
  20. Cumplimiento de Seguridad y Cifrado


═══════════════════════════════════════════════════════════════
  1. COMO INICIAR EL PROGRAMA
═══════════════════════════════════════════════════════════════

  Paso 1: Busque el icono de RegistroDoc en su escritorio.
  Paso 2: Haga DOBLE CLIC con el boton izquierdo del raton.
  Paso 3: Espere unos segundos mientras se carga la base de datos.
  Paso 4: Vera la pantalla principal con el menu a la izquierda.

  IMPORTANTE: Si es la primera vez, el programa le pedira
  configurar sus datos (nombre, escuela, grados, materias).

  NOTA TECNICA: El programa descifra automaticamente su base de
  datos local (SQLite) al iniciar. Todos los datos se leen desde
  ahi — no desde el archivo Excel.


═══════════════════════════════════════════════════════════════
  2. PANTALLA DE INICIO (DASHBOARD)
═══════════════════════════════════════════════════════════════

  Es lo primero que vera al abrir el programa.

  Que muestra el Dashboard:
  • Tarjetas con total de alumnos, alumnos en riesgo y
    el alumno con mejor promedio del salon.
  • Al buscar un alumno (escribiendo su nombre) las tarjetas
    y graficas muestran informacion individual.
  • Frase motivacional del dia.
  • Tareas pendientes con colores de urgencia.
  • Graficos de rendimiento academico del trimestre.
  • Panel de exportacion a Excel (pie de pagina).
  • Boton de Respaldo rapido.

  CONSEJO: Presione Escape en cualquier pantalla para
  volver aqui rapidamente. Use Ctrl+1 desde cualquier modulo.


═══════════════════════════════════════════════════════════════
  3. MODULO DE ESTUDIANTES
═══════════════════════════════════════════════════════════════

  Atajo de teclado: Ctrl+2

  La lista muestra numeracion correlativa (1, 2, 3...).
  Los IDs internos de la base de datos son invisibles al docente.

  AGREGAR un estudiante nuevo:
    1. Presione "Estudiantes" en el menu izquierdo.
    2. Escriba: APELLIDO, Nombre (con mayusculas en apellido).
    3. Escriba la cedula (opcional, formato 4-123-4567).
    4. Seleccione el sexo y presione "Guardar Nuevo".
    5. El alumno aparecera numerado al final de la lista.

  CORREGIR un nombre o cedula:
    1. Haga clic en la casilla del nombre o cedula.
    2. Borre y escriba el correcto.
    3. Presione "GUARDAR MODIFICACIONES DE LA LISTA".

  RETIRAR un estudiante:
    1. Presione el boton rojo (papelera) junto al nombre.
    2. Confirme presionando "Si".

  IMPRIMIR LISTA DE CLASE:
    1. Presione el boton "Lista de Clase" (esquina superior derecha).
    2. Se generara un documento Word con el listado oficial.

  NOTA: Maximo 36 alumnos por grado (premedia) o 34 (primaria).


═══════════════════════════════════════════════════════════════
  4. MODULO DE CALIFICACIONES (NOTAS)
═══════════════════════════════════════════════════════════════

  Atajo de teclado: Ctrl+3

  REGISTRAR una nota nueva:
    1. Seleccione Grado, Materia y Trimestre.
    2. Elija el Tipo de nota:
       • Parcial / Diaria  = tareas y talleres (1.0 a 5.0)
       • Apreciacion       = cuadernos, participacion (1.0 a 5.0)
       • Examen            = prueba trimestral oficial
    3. Escriba la Fecha (DD-MM) y Descripcion del trabajo.
    4. Coloque la nota de cada alumno en su casilla.
       (Use las teclas Enter / Abajo / Arriba para navegar)
    5. Presione el boton verde "GUARDAR NUEVA".
       El guardado se realiza al instante en SQLite.

  CORREGIR una nota ya guardada:
    1. Cambie a la pestana "Modificar" dentro de Notas.
    2. Seleccione la descripcion del trabajo en el desplegable.
    3. La fecha es NO MODIFICABLE (protege el historial oficial).
    4. Realice las correcciones y presione "ACTUALIZAR NOTAS".

  MODO PUNTOS:
    - Active "Usar Puntos" para ingresar puntaje obtenido / maximo.
    - El programa calcula la nota automaticamente.


═══════════════════════════════════════════════════════════════
  5. MODULO DE ASISTENCIA
═══════════════════════════════════════════════════════════════

  Atajo de teclado: Ctrl+4

  PASAR LISTA:
    1. Seleccione Grado y Trimestre.
    2. Escriba la fecha de hoy (DD-MM).
    3. Presione "Todos Presentes" (un solo clic para marcar el grupo).
    4. Corrija SOLO las excepciones con la simbologia oficial MEDUCA:
       • P = Presente   (se guarda como '.' en SQLite)
       • A = Ausente    (se guarda como '-' en SQLite)
       • T = Tardanza   (se guarda como 'T' en SQLite)
       • E = Excusa     (falta justificada, NO resta puntos de asistencia)
    5. Escriba el motivo de ausencia / tardanza si corresponde.
    6. Presione "GUARDAR ASISTENCIA".

  MODIFICAR asistencia ya guardada:
    1. Pestana "Modificar" → seleccione el Trimestre y la Fecha.
    2. Presione "Cargar a la Lista".
    3. Edite los estados y presione "ACTUALIZAR".

  NOTA: El sistema genera automaticamente un expediente Word
  para las ausencias y tardanzas con su justificacion.


═══════════════════════════════════════════════════════════════
  6. MODULO DE OBSERVACIONES (EXPEDIENTE)
═══════════════════════════════════════════════════════════════

  Atajo de teclado: Ctrl+5

  REGISTRAR una observacion:
    1. Seleccione el Grado.
    2. Haga clic en el nombre del alumno.
    3. Elija la categoria:
       • Conducta         (comportamiento en clase)
       • Academico        (rendimiento escolar)
       • Citacion         (llamado al acudiente)
       • Merito           (reconocimiento positivo)
    4. Escriba la observacion o elija una plantilla predefinida.
    5. Presione "GUARDAR EN EXPEDIENTE OFICIAL".

  El programa guarda la observacion en SQLite Y crea / actualiza
  automaticamente un documento Word en "Expedientes_Estudiantes"
  con el historial completo del alumno. Listo para imprimir.

  CORRECTOR ORTOGRAFICO:
    - Las palabras mal escritas se subrayan en rojo.
    - Haga clic derecho en una palabra para ver sugerencias.


═══════════════════════════════════════════════════════════════
  7. MODULO DE HABITOS Y APTITUDES
═══════════════════════════════════════════════════════════════

  Atajo de teclado: Ctrl+6

  CALIFICAR habitos:
    1. Seleccione Grado, Trimestre y Frecuencia:
       • Diario  = califica un dia especifico
       • Semanal = califica una semana completa
       • Mensual = califica un mes
    2. Haga clic en un alumno de la lista.
    3. Califique cada criterio:
       • S (Satisfactorio)  → Verde
       • R (Regular)        → Amarillo
       • X (No Satisface)   → Rojo
    4. Use el boton "Sugerencia de IA" para analisis automatico.
    5. Presione "GUARDAR EVALUACIONES".

  ADVERTENCIA: La frecuencia se BLOQUEA despues del primer
  guardado del trimestre para mantener los datos ordenados.


═══════════════════════════════════════════════════════════════
  8. MODULO DE TAREAS PROGRAMADAS
═══════════════════════════════════════════════════════════════

  Acceso: Boton "Tareas" en el menu izquierdo.

  CREAR una tarea:
    1. Escriba el titulo (ej: "Examen de Espanol T1").
    2. Seleccione Grado, Materia y Tipo.
    3. Escriba la fecha limite (DD-MM-YYYY).
    4. Presione "PROGRAMAR TAREA".

  COLORES de urgencia en el Dashboard:
    Rojo    = Tarea vencida (ya paso la fecha)
    Amarillo = Vence HOY
    Naranja = Vence en 1-2 dias
    Azul    = Fecha normal / futura

  COMPLETAR: Presione el boton junto a la tarea.
  Al presionar una tarea en el Dashboard, va directo a Notas.


═══════════════════════════════════════════════════════════════
  9. MODULO DE REPORTES Y GRAFICOS
═══════════════════════════════════════════════════════════════

  Atajos: Ctrl+7 (Reportes) / Ctrl+8 (Graficos)

  REPORTES: Tablas con promedios, asistencia y rendimiento
  general por grado, materia y trimestre. Datos 100% desde SQL.

  GRAFICOS: Barras y lineas para presentar en reuniones de
  padres o a la direccion de la escuela.

  CUADRO DE HONOR: Lista automatica de los mejores alumnos
  por trimestre basada en promedios calculados desde SQLite.


═══════════════════════════════════════════════════════════════
  10. MODULO DE REGISTRO COMPLETO (CONSOLIDADO)
═══════════════════════════════════════════════════════════════

  VISTA DE CALIFICACIONES:
    1. Presione "Registro Completo" en el menu izquierdo.
    2. Seleccione Grado, Trimestre y Materia.
    3. Tarjetas resumen: actividades tomadas, promedio grupal,
       porcentaje de aprobacion y nota maxima del salon.
    4. Cuadricula: un alumno por fila, una columna por actividad.
       La fecha aparece debajo del titulo de cada tarea.
       Verde = nota alta (>=4.5), Rojo = reprobatoria (<3.0).

  VISTA DE ASISTENCIA:
    1. Cambie a la pestana "Asistencia".
    2. Tarjetas: dias evaluados, % asistencia promedio,
       alumnos con 100% y alumnos en alerta (<90%).
    3. Matriz: historial completo por fecha con colores indicativos
       y porcentaje de asistencia acumulado por alumno.

  NOTA: Cambiar trimestre, grado o materia actualiza los datos
  instantaneamente sin recargar toda la pantalla.


═══════════════════════════════════════════════════════════════
  11. MODULO DE IMPRESION
═══════════════════════════════════════════════════════════════

  Atajo de teclado: Ctrl+9

  Tipos de documentos:
  • Portada oficial del grado
  • Planilla de calificaciones por materia
  • Lista de asistencia del trimestre
  • Resumen de notas (Reporte Docente)
  • Reporte de Direccion / Aprobados
  • Plantilla auxiliar en blanco (para impresion manual)

  COMO IMPRIMIR:
    1. Seleccione el tipo de documento.
    2. Seleccione Grado y Materia (si aplica).
    3. Presione "Abrir en Excel" para revisar antes.
    4. O presione "Enviar a Impresora" para imprimir directo.

  NOTA: Los datos se obtienen de SQLite y se vuelcan al Excel
  para impresion. El archivo Excel de plantilla no se altera.


═══════════════════════════════════════════════════════════════
  12. MODULO DE REUNIONES Y ACTAS
═══════════════════════════════════════════════════════════════

  REUNION DE DOCENTES:
    Acta con: fecha, hora, lugar, asistentes, agenda,
    acuerdos y compromisos, firmas.

  REUNION CON PADRES DE FAMILIA:
    Formato de citacion con: datos del alumno, motivo,
    descripcion, acuerdos, compromisos, firmas.

  Acceso: Menu izquierdo → "Reuniones"


═══════════════════════════════════════════════════════════════
  13. CONFIGURACION DEL SISTEMA
═══════════════════════════════════════════════════════════════

  Configure sus datos institucionales:
  • Nombre del docente, cedula, telefono, correo
  • Nombre de la escuela, region, distrito
  • Director, subdirector, coordinador
  • Horario de clases (se muestra en el Dashboard)
  • Gestion de materias y grados

  SINCRONIZAR: El boton "Sincronizar y Sobreescribir" copia
  todos estos datos al archivo Excel oficial automaticamente.


═══════════════════════════════════════════════════════════════
  14. ATAJOS DE TECLADO
═══════════════════════════════════════════════════════════════

  Ctrl+1  → Dashboard (Inicio)
  Ctrl+2  → Estudiantes
  Ctrl+3  → Notas / Calificaciones
  Ctrl+4  → Asistencia
  Ctrl+5  → Observaciones
  Ctrl+6  → Habitos
  Ctrl+7  → Reportes
  Ctrl+8  → Graficos
  Ctrl+9  → Impresion
  Ctrl+S  → Guardar en la pantalla actual
  F1      → Abrir esta Ayuda
  Escape  → Volver al Dashboard

  CONSEJO: Use Ctrl+1 a Ctrl+9 para navegar sin el raton.


═══════════════════════════════════════════════════════════════
  15. COPIAS DE SEGURIDAD (RESPALDOS)
═══════════════════════════════════════════════════════════════

  Su informacion es valiosa. Haga respaldos frecuentes.

  RESPALDO AUTOMATICO:
    El programa hace una copia silenciosa cada 30 minutos
    en la carpeta "Respaldos_Auto". Maximo 10 copias.
    Formato: registro_db_backup_YYYY-MM-DD_HH-MM.db.enc

  RESPALDO MANUAL:
    1. En el Dashboard, presione "Respaldo".
    2. Se crea una copia en "Respaldos_Locales" con timestamp.

  RESPALDO EN USB:
    Copie la carpeta "data/" y "Expedientes_Estudiantes/"
    a su memoria USB. Haga esto cada viernes.

  NUNCA borre el archivo registro.db.enc sin tener respaldo.


═══════════════════════════════════════════════════════════════
  16. EXPORTAR A EXCEL
═══════════════════════════════════════════════════════════════

  RegistroDoc Pro trabaja internamente con SQLite (rapido).
  Cuando necesite el archivo Excel oficial MEDUCA:

  1. Vaya al Dashboard (Ctrl+1).
  2. Localice el panel "Exportar Calificaciones" en el pie de pagina.
  3. Seleccione "Todos los trimestres" o un trimestre especifico.
  4. Presione el boton Excel.
  5. El programa vuelca TODOS los datos de SQLite al archivo Excel.

  IMPORTANTE: El Excel resultante es para impresion y entrega
  a la direccion. No lo use para editar datos manualmente;
  use el programa para eso.


═══════════════════════════════════════════════════════════════
  17. SOLUCION DE PROBLEMAS COMUNES
═══════════════════════════════════════════════════════════════

  PROBLEMA: "Las notas no aparecen en mi archivo Excel fisico"
  SOLUCION: Use la funcion Exportar a Excel del Dashboard (seccion 16).
            El programa trabaja en SQLite por velocidad. El Excel
            se actualiza solo cuando usted lo exporta.

  PROBLEMA: "El programa muestra letras y numeros raros"
  SOLUCION: Los datos estan cifrados. Si ve texto como
            "ed6f444c..." es un problema de descifrado.
            Contacte soporte. Sus datos estan seguros.

  PROBLEMA: "No puedo cambiar la frecuencia de Habitos"
  SOLUCION: La frecuencia se bloquea tras el primer guardado
            del trimestre. Espere al siguiente trimestre.

  PROBLEMA: "Error al guardar / Acceso denegado"
  SOLUCION: No tenga el Excel abierto en Excel/LibreOffice
            mientras el programa exporta.

  PROBLEMA: "El programa no abre"
  SOLUCION: Haga clic derecho → "Ejecutar como administrador".
            Verifique que Python 3.10+ esta instalado.

  PROBLEMA: "Perdi mis datos"
  SOLUCION: Busque en "Respaldos_Auto" o "Respaldos_Locales".
            El archivo mas reciente tiene sus datos guardados.
            Renombrelo a "registro.db.enc" en la carpeta data/.


═══════════════════════════════════════════════════════════════
  18. GLOSARIO DE TERMINOS MEDUCA
═══════════════════════════════════════════════════════════════

  CALIFICACIONES:
  Parcial / Diaria  = Nota de tareas y talleres (1.0 a 5.0)
  Apreciacion       = Nota de cuadernos y participacion (1.0 a 5.0)
  Examen            = Prueba trimestral oficial
  Trimestre         = Periodo academico (3 por ano lectivo)

  ASISTENCIA (simbologia oficial MEDUCA):
  . (Presente)      = Asistio a clase
  - (Ausente)       = Falto sin justificacion
  T (Tardanza)      = Llego tarde
  E (Excusa)        = Falta justificada, NO resta puntos

  HABITOS Y APTITUDES:
  S (Satisfactorio) = Cumple bien el criterio     → Verde
  R (Regular)       = Cumple parcialmente          → Amarillo
  X (No Satisface)  = Requiere atencion urgente    → Rojo

  OTROS:
  Expediente        = Documento Word con historial del alumno
  Sincronizar       = Copiar configuracion al Excel oficial
  SQLite            = Base de datos local de alta velocidad
  AES-256-GCM       = Cifrado militar para proteger sus datos
  Exportar          = Volcar datos de SQLite al Excel MEDUCA


═══════════════════════════════════════════════════════════════
  19. ESTRUCTURA DE BASE DE DATOS Y TABLAS (SQLTools)
═══════════════════════════════════════════════════════════════

  RegistroDoc Pro almacena todos sus datos en SQLite cifrado.
  Para explorar en tiempo real con VS Code SQLTools:

  Paso 1: Inicie el programa. La BD temporal se descifra en:
          C:\Users\<Usuario>\AppData\Local\RegistroDoc\temp\sqlite_temp.db

  Paso 2: En VS Code, instale extension "SQLTools" y driver
          "SQLTools SQLite" (por Matheus Teixeira).

  Paso 3: Cree una conexion apuntando al sqlite_temp.db.

  TABLAS DISPONIBLES:
  1. configuracion  — opciones generales (clave/valor)
  2. grados         — grupos con modalidad (primaria/premedia)
  3. estudiantes    — datos cifrados (nombre, cedula, sexo, grado_id)
  4. materias       — asignaturas por grado
  5. horario        — planificador semanal
  6. notas          — calificaciones (tipo, descripcion, valor, fecha, trimestre)
  7. asistencia     — registros diarios (estado, motivo, fecha)
  8. observaciones  — historial de conducta del expediente
  9. habitos        — evaluaciones S/R/X por criterio y trimestre
  10. tareas        — recordatorios programados con fecha limite
  11. reuniones     — actas y citaciones
  12. auditoria     — log inmutable de todas las transacciones

  NOTA: Al cerrar el programa, sqlite_temp.db es eliminado
  y sobreescrito de forma segura (wiping con datos aleatorios).


═══════════════════════════════════════════════════════════════
  20. CUMPLIMIENTO DE SEGURIDAD Y CIFRADO
═══════════════════════════════════════════════════════════════

  ISO/IEC 27001:2022 — Control A.8.24 (Cifrado):
    - AES-256-GCM en reposo para la base de datos de produccion.
    - La llave se deriva del numero de serie del hardware local
      con PBKDF2-HMAC-SHA256 (600,000 iteraciones). No se guarda
      ninguna llave en texto plano en el disco.
    - Nonce aleatorio unico por cada escritura (sin reutilizacion).
    - Columnas sensibles (nombres, cedulas) cifradas individualmente
      con llave derivada adicional por columna.

  Ley 81 del 26 de marzo de 2019 (Panama):
    - Toda la informacion personal (cedulas, nombres, promedios)
      es cifrada y permanece 100% local en la maquina del docente.
    - No hay envio de datos por internet en ninguna operacion.
    - Los entornos de prueba usan datos ficticios generados.

  NIST SP 800-38D:
    - Modo Galois/Counter Mode (GCM) que verifica autenticidad e
      integridad del archivo cifrado antes de abrirlo.

  Robustez contra corrupcion:
    - Escrituras atomicas: se usa archivo .tmp antes de renombrar.
    - Respaldos automaticos cada 30 minutos.
    - Log de auditoria inmutable para no repudio.
    - Desinstalacion segura que exige cedula del docente.


═══════════════════════════════════════════════════════════════
  (c) 2026 RegistroDoc Pro — MEDUCA Panama
  "Instruye al nino en su camino, y aun cuando
   fuere viejo no se apartara de el." — Proverbios 22:6
═══════════════════════════════════════════════════════════════
