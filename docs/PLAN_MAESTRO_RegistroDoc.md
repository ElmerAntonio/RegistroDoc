
# RegistroDoc Pro — Plan Maestro de Mejora
## Sistema de Registro Académico | Panamá 2026

---

## Estado actual del sistema

| Archivo | Líneas | Tamaño | Estado |
|---|---|---|---|
| app.py (principal) | 1,122 | 56.6 KB | ✅ Funcional |
| rdsecurity.py | 622 | 25.1 KB | ✅ AES-256 |
| rdprint.py | 210 | 8.3 KB | ⚠️ Básico |
| generador_codigos.py | 303 | 14.3 KB | ✅ Funcional |
| Documentación | 676 | 36.1 KB | ✅ Completa |
| **TOTAL** | **3,065** | **145.5 KB** | |

---

## Equipo de Agentes Especializados

Cada agente tiene un rol fijo, responsabilidades claras y entregables medibles.

---

### 🔵 AGENTE 1 — Arquitecto de Seguridad
**Nombre en equipo:** SecureArch  
**Especialidad:** Criptografía, protección de datos, cumplimiento legal  
**Responsabilidades:**
- Auditar y fortalecer el módulo `rdsecurity.py`
- Implementar renovación automática de licencias por año
- Agregar firma digital HMAC a cada nota guardada
- Implementar modo de emergencia para bloqueo remoto
- Verificar cumplimiento ISO 27001, Ley 81/2019 Panamá

**Entregables:**
- `rdsecurity_v3.py` con firma digital por registro
- `rdlicense.py` — gestor avanzado de licencias
- Informe de cumplimiento normativo

---

### 🟢 AGENTE 2 — Desarrollador de Interfaz
**Nombre en equipo:** UIBuilder  
**Especialidad:** Tkinter, experiencia de usuario, diseño visual  
**Responsabilidades:**
- Rediseñar la interfaz con modo oscuro/claro
- Crear pantalla de inicio con logo y progreso de carga
- Mejorar el panel de notas con tabla visual tipo grid
- Agregar indicadores de progreso al guardar
- Crear modo de vista previa antes de imprimir

**Entregables:**
- `rdui.py` — biblioteca de componentes visuales reutilizables
- `rdthemes.py` — sistema de temas visuales
- `app_v3.py` — interfaz rediseñada

---

### 🟡 AGENTE 3 — Especialista en Datos
**Nombre en equipo:** DataEngine  
**Especialidad:** Excel, openpyxl, estructuras de datos, fórmulas MEDUCA  
**Responsabilidades:**
- Mapear y documentar TODAS las celdas del Excel (todas las materias)
- Crear motor de lectura/escritura robusto con manejo de errores
- Implementar transacciones (rollback si falla a mitad)
- Optimizar velocidad de lectura del Excel
- Agregar soporte para nuevas materias dinámicas

**Entregables:**
- `rddata.py` — motor de datos refactorizado
- `rdmapping.json` — mapa completo de celdas documentado
- Tests unitarios para todas las operaciones Excel

---

### 🟠 AGENTE 4 — Especialista en Reportes e Impresión
**Nombre en equipo:** ReportPro  
**Especialidad:** Impresión, PDF, reportes visuales  
**Responsabilidades:**
- Mejorar módulo de impresión con vista previa integrada
- Generar reportes PDF desde el programa (sin abrir Excel)
- Crear reporte de boletín individual por estudiante
- Exportar resumen de calificaciones a PDF
- Generar lista de asistencia imprimible

**Entregables:**
- `rdprint_v2.py` — impresión avanzada con vista previa
- `rdreports.py` — generador de reportes PDF
- Plantillas de boletín y resumen

---

### 🔴 AGENTE 5 — Especialista en Marketing y Comercio
**Nombre en equipo:** MarketDev  
**Especialidad:** Branding, ventas, documentación comercial  
**Responsabilidades:**
- Crear página web de producto (HTML estático)
- Diseñar material de venta (folleto digital)
- Crear sistema de precios y planes
- Redactar guion de demostración del producto
- Crear estrategia de distribución entre docentes panameños

**Entregables:**
- `landing_page.html` — página web del producto
- `folleto_ventas.pdf` — brochure digital
- Plan de precios y estrategia de venta
- Guion de demo de 5 minutos

---

### 🟣 AGENTE 6 — Especialista en Calidad y Testing
**Nombre en equipo:** QAGuard  
**Especialidad:** Pruebas, validación, documentación técnica  
**Responsabilidades:**
- Crear suite de pruebas automatizadas
- Probar todos los casos límite de notas MEDUCA
- Validar que el Excel nunca se corrompe
- Crear checklist de entrega para cada versión
- Documentar todos los bugs encontrados y resueltos

**Entregables:**
- `test_suite.py` — pruebas automatizadas completas
- `CHECKLIST_QA.md` — lista de verificación de calidad
- Informe de bugs y resoluciones

---

## Cronograma por Fases

### FASE 1 — Estabilización (Semanas 1–2)
*Asegurar que lo que existe funciona perfectamente*

| Tarea | Agente | Prioridad | Estado |
|---|---|---|---|
| Auditar rdsecurity.py completo | SecureArch | 🔴 Alta | Pendiente |
| Mapear TODAS las celdas del Excel | DataEngine | 🔴 Alta | Pendiente |
| Crear tests básicos de funciones críticas | QAGuard | 🔴 Alta | Pendiente |
| Verificar que el .exe genera correctamente | UIBuilder | 🔴 Alta | Pendiente |
| Probar activación en Windows real | QAGuard | 🔴 Alta | Pendiente |

**Objetivo de la fase:** El programa funciona sin errores en una PC real con Windows.

---

### FASE 2 — Mejoras Funcionales (Semanas 3–5)
*Agregar lo que hace falta para un producto completo*

| Tarea | Agente | Prioridad | Semana |
|---|---|---|---|
| Mapeo completo de celdas de todas las materias | DataEngine | 🔴 Alta | 3 |
| Vista previa antes de imprimir | ReportPro | 🔴 Alta | 3 |
| Generación de boletín PDF por estudiante | ReportPro | 🔴 Alta | 4 |
| Rediseño del panel de notas (tabla visual) | UIBuilder | 🟡 Media | 4 |
| Firma digital HMAC por nota guardada | SecureArch | 🟡 Media | 5 |
| Sistema de respaldo automático programado | DataEngine | 🟡 Media | 5 |

**Objetivo de la fase:** El docente puede hacer todo desde el programa sin tocar el Excel.

---

### FASE 3 — Pulido y Comercialización (Semanas 6–8)
*Preparar el producto para venta masiva*

| Tarea | Agente | Prioridad | Semana |
|---|---|---|---|
| Pantalla de carga con logo | UIBuilder | 🟡 Media | 6 |
| Modo oscuro / modo claro | UIBuilder | 🟢 Baja | 6 |
| Landing page HTML del producto | MarketDev | 🔴 Alta | 6 |
| Folleto digital de ventas | MarketDev | 🔴 Alta | 7 |
| Sistema de actualizaciones por versión | SecureArch | 🟡 Media | 7 |
| Suite de pruebas automatizadas completa | QAGuard | 🟡 Media | 7 |
| Video demo de 5 minutos (guion) | MarketDev | 🟢 Baja | 8 |
| Documentación técnica final | QAGuard | 🟢 Baja | 8 |

**Objetivo de la fase:** Producto listo para vender a cualquier docente de Panamá.

---

### FASE 4 — Expansión (Mes 3 en adelante)
*Crecer el producto hacia otros mercados*

| Iniciativa | Agente | Impacto |
|---|---|---|
| Soporte para otras modalidades educativas | DataEngine | 🔴 Alto |
| Versión para Primaria | DataEngine + UIBuilder | 🔴 Alto |
| Panel web del vendedor (ver licencias online) | SecureArch + MarketDev | 🟡 Medio |
| App de respaldo en Android (solo lectura) | UIBuilder | 🟢 Bajo |
| Integración con Google Drive para respaldos | DataEngine | 🟡 Medio |

---

## Plan de Mejoras por Área

### 🔐 Seguridad
| Mejora | Impacto | Complejidad |
|---|---|---|
| Firma HMAC por cada nota (no solo el Excel) | Alto | Media |
| Renovación anual de licencia automática | Alto | Media |
| Bloqueo remoto de emergencia | Alto | Alta |
| Cifrado adicional del Excel con contraseña | Medio | Baja |
| Detección de VM (evitar análisis en máquinas virtuales) | Medio | Alta |

### 📊 Estructura de Datos
| Mejora | Impacto | Complejidad |
|---|---|---|
| Mapeo completo de las 38 hojas del Excel | Alto | Baja |
| Transacciones con rollback en escritura | Alto | Media |
| Caché en memoria para lecturas frecuentes | Medio | Baja |
| Validación de fórmulas al cargar el Excel | Alto | Media |
| Soporte para nuevas materias sin hardcodear | Alto | Media |

### 🖥️ Interfaz y Experiencia
| Mejora | Impacto | Complejidad |
|---|---|---|
| Tabla visual de notas (como hoja de cálculo) | Alto | Media |
| Vista previa de impresión integrada | Alto | Media |
| Búsqueda de estudiante por nombre | Medio | Baja |
| Historial de cambios visible por el docente | Medio | Baja |
| Modo pantalla completa | Bajo | Baja |

### 🖨️ Reportes e Impresión
| Mejora | Impacto | Complejidad |
|---|---|---|
| Boletín PDF individual por estudiante | Alto | Media |
| Reporte de asistencia mensual | Alto | Baja |
| Exportar lista de estudiantes a PDF | Medio | Baja |
| Reporte de rendimiento del grupo | Medio | Media |
| Certificado de notas por trimestre | Medio | Media |

### 📣 Marketing
| Mejora | Impacto | Complejidad |
|---|---|---|
| Landing page con demo en video | Alto | Baja |
| Sistema de referidos entre docentes | Alto | Media |
| Versión de prueba de 30 días | Alto | Baja |
| Grupo de WhatsApp de usuarios | Medio | Baja |
| Testimonios de docentes que lo usan | Alto | Baja |

---

## Métricas de Éxito

### Técnicas
- 0 errores críticos en 30 días de uso
- Tiempo de carga menor a 3 segundos
- Guardado de nota en menos de 1 segundo
- 100% de notas validadas correctamente en escala MEDUCA

### Comerciales
- Objetivo Mes 1: 5 docentes activados
- Objetivo Mes 3: 20 docentes
- Objetivo Mes 6: 50 docentes en toda la Comarca Ngäbe Buglé
- Precio sugerido: B/. 25.00 por licencia anual

### Calidad
- Manual leído y comprendido por el 100% de usuarios
- Tasa de activación exitosa: 95% sin ayuda del vendedor
- Satisfacción del docente: 4/5 estrellas mínimo

---

## Próximas acciones inmediatas (esta semana)

1. **Probar el .exe en Windows** — Confirmar que se instala y activa correctamente
2. **Mapear todas las materias del Excel** — Completar el mapa de celdas que falta
3. **Integrar rdprint.py en el .exe** — Verificar que la impresión funciona
4. **Crear una licencia de prueba** — Para demos a docentes interesados
5. **Subir todo a GitHub** — Con el .gitignore correcto

---

*Plan elaborado por: Isabel — Asistente de Planificación Educativa y Desarrollo*  
*Fecha: Marzo 2026 | Versión del documento: 1.0*


### Actualización de rddata.py (DataEngine)
- Implementa sistema de caché `_wb_cache`.
- Manejo exhaustivo de `try/except PermissionError`.
- Lógica bimodal integrada: 28 notas continuas para Primaria, secciones divididas para Premedia.
- Sincronización automática de carátulas (Q24, Q25) con `sapp.py`.
