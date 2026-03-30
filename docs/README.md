# RegistroDoc Pro v3.0

RegistroDoc Pro v3.0 es un software de escritorio independiente y comercial desarrollado en Python, diseñado para facilitar la gestión académica (calificaciones, asistencia, reportes y métricas) en instituciones educativas bajo la reglamentación del sistema de Panamá, sin embargo no está afiliado con MEDUCA u organizaciones gubernamentales.

El software integra una arquitectura fija y fluida con customtkinter y pandas/openpyxl, además de un estricto módulo criptográfico.

---

## Estructura del Proyecto

```
.
├── docs/                   # Documentación y diagramas
├── img/                    # Activos visuales (iconos, avatar, logos)
├── src/
│   ├── app.py              # Punto de entrada y contenedor principal de la Arquitectura (Header/Sidebar fijos)
│   ├── dashapp.py          # Panel de control de métricas y gráficas (DashboardFrame)
│   ├── dapp.py             # Módulo de gestión de Estudiantes
│   ├── eapp.py             # Módulo de gestión de Notas
│   ├── fapp.py             # Módulo de gestión de Asistencia
│   ├── obsapp.py           # Módulo de Observaciones / Comentarios
│   ├── sapp.py             # Módulo de Configuración de usuario
│   ├── rddata.py           # Motor de Datos: CRUD sobre Excel usando Pandas y Openpyxl
│   ├── rdsecurity.py       # Motor Criptográfico y Validación de Seguridad
│   └── setup.py            # Asistente de configuración inicial (si no existe perfil.json)
├── Registro_*.xlsx         # Plantillas Base Excel (Almacenamiento Local)
├── perfil.json             # Configuración actual del usuario y su espacio de trabajo
└── requirements.txt        # Dependencias Python
```

## Arquitectura de la Interfaz

La versión 3.0 introdujo una recodificación principal en `app.py`. En lugar de destruir la pantalla completa en cada navegación de módulo, se estableció `MainApplication` como un **contenedor estático** que retiene:

- Un **Sidebar** plegable a la izquierda con el Logo integrado.
- Un **Header** superior con búsqueda (ligada al dashboard), notificaciones y perfil de usuario.
- Un **main_content_frame** (área de trabajo) que se limpia de manera asíncrona permitiendo cambiar los frames (notas, asistencia, dashboard) manteniendo la navegación viva.

El rediseño evita parpadeos (flickering), y el código UI usa colores hexadecimales de 6 dígitos para asegurar total compatibilidad con la librería de CustomTkinter.

---

## Seguridad y Cumplimiento de Normativas

`rdsecurity.py` es el núcleo de protección de datos, logrando que el software cumpla con normativas rigurosas:

1. **Ley 81 de 2019 (Panamá)**: Protección de Datos Personales mediante el cifrado de la configuración.
2. **ISO/IEC 27001:2022 y 27002:2022**: Gestión de Seguridad, controles (A.8.2, A.8.15, A.8.24) e inventario criptográfico y log de auditorías (`rd_audit.bin`).
3. **NIST SP 800-63B / SP 800-38D**: Autenticación digital y cifrado usando el estándar **AES-256-GCM**.

El módulo genera una **huella de hardware** irrompible (`_hw_fingerprint`) basada en UUID de la placa, volumen C: y nodo, atando la licencia al dispositivo. También cifra todos los archivos `.bin` con PBKDF2-HMAC-SHA256 (600,000 iteraciones).

---

## Instrucciones de Ejecución

### Prerrequisitos (Entorno de Desarrollo VS Code)
Se recomienda abrir la carpeta raíz en VS Code. La configuración utilizada por los desarrolladores contempla Pyenv o un entorno virtual (venv) para la resolución de librerías.

### Instalación de Dependencias

Ejecute en la terminal:

```bash
pip install -r requirements.txt
```

### Iniciar la Aplicación

Para lanzar la aplicación con los permisos de lectura/escritura en el mismo directorio, ejecute:

```bash
python src/app.py
```

En la primera ejecución, si `perfil.json` no existe, se disparará automáticamente el SetupWizard (si está disponible) o se creará un perfil inicial básico.

---

## Créditos y Desarrolladores

El proyecto ha sido desarrollado independientemente por el Autor original junto con asistencias puntuales de IAs avanzadas en ingeniería de software y refactorización para lograr su modernización:
- **Autor Principal:** Creador Independiente (Panamá)
- **Codex / Claude:** Asistencia algorítmica y de lógica base
- **Jules:** Refactorización Arquitectónica, Estabilidad de Interfaz y Auditoría de Código (V3.0)
