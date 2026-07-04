# RegistroDoc Pro v5.0 — Sistema de Registro Académico (Panamá)

[![Python](https://img.shields.io/badge/Python-3.10%2B-blue.svg)](https://www.python.org/)
[![Tests Status](https://img.shields.io/badge/Tests-90%2F90%20Passed-brightgreen.svg)](#)
[![Architecture](https://img.shields.io/badge/Datos-SQL--First-orange.svg)](#)
[![Security](https://img.shields.io/badge/Security-AES--256--GCM-darkblue.svg)](#)

**RegistroDoc Pro** es una aplicación de escritorio premium para docentes panameños (Primaria y Premedia). Automatiza el registro académico bajo los lineamientos del **Ministerio de Educación (MEDUCA)** con arquitectura **SQL-First**: SQLite es la única fuente de datos en tiempo real; Excel y Word son destinos de exportación, nunca fuentes de lectura.

---

## Arquitectura SQL-First (v5.0)

| Capa | Responsabilidad |
|------|----------------|
| `rdsql.py` | Motor SQLite — cifrado AES-256-GCM, tablas, triggers, auditoria |
| `rddata.py` | DataEngine — toda la logica de negocio sobre SQL |
| `src/*.py` | Modulos de UI — leen exclusivamente desde DataEngine |
| Excel / Word | Solo exportacion e impresion bajo demanda |

---

## Caracteristicas Clave

- Base de datos SQLite cifrada AES-256-GCM — respuesta en milisegundos
- Exportacion a Excel bajo demanda (plantilla oficial MEDUCA)
- Soporte bimodal: Primaria y Premedia
- 100% offline — zonas rurales y Comarca Ngabe Bugle
- Clave criptografica derivada del hardware del equipo (PBKDF2-HMAC-SHA256, 600 000 iteraciones)
- Numeracion visual 1,2,3... — IDs internos ocultos al docente
- Navegacion sin recarga de pagina — cache de frames en memoria
- 90/90 tests automatizados

---

## Estructura del Proyecto

```
RegistroDoc/
├── assets/templates/     # Plantillas MEDUCA vacias (Premedia/Primaria)
├── data/                 # Base de datos cifrada (.db.enc)
├── docs/                 # Manual de usuario y documentacion tecnica
├── src/
│   ├── app.py            # Punto de entrada y navegador (cache de frames)
│   ├── rdsql.py          # Motor SQLite y cifrado
│   ├── rddata.py         # DataEngine — logica SQL
│   ├── rdsecurity.py     # Modulo criptografico AES-256-GCM
│   ├── dashapp.py        # Dashboard
│   ├── eapp.py           # Calificaciones
│   ├── fapp.py           # Asistencia
│   ├── dapp.py           # Gestion de estudiantes
│   ├── habapp.py         # Habitos y aptitudes
│   ├── obsapp.py         # Observaciones / expedientes
│   ├── registrocompletoapp.py  # Registro consolidado
│   ├── tareasapp.py      # Tareas programadas
│   ├── reunionesapp.py   # Reuniones y actas
│   ├── happ.py           # Reportes y graficos
│   ├── rdprint.py        # Panel de impresion
│   ├── impp.py           # Importador de estudiantes
│   └── helpapp.py        # Centro de ayuda offline
├── tests/                # Suite de 90 pruebas automatizadas
├── Respaldos_Auto/       # Respaldos automaticos (cada 30 min, max 10)
├── Respaldos_Locales/    # Respaldos manuales
└── Expedientes_Estudiantes/  # Documentos Word por alumno
```

---

## Ejecutar el Programa

```bash
# Desarrollo
python src/app.py

# Con logs de debug
set REGISTRODOC_DEV_MODE=1
python src/app.py
```

---

## Explorar la BD en Tiempo Real

Mientras la app esta abierta la BD temporal descifrada se encuentra en:
```
C:\Users\<Usuario>\AppData\Local\RegistroDoc\temp\sqlite_temp.db
```

Tablas disponibles: `configuracion`, `grados`, `estudiantes`, `materias`, `horario`, `notas`, `asistencia`, `observaciones`, `habitos`, `tareas`, `reuniones`, `auditoria`.

Al cerrar la app el archivo temporal es eliminado con sobreescritura segura.

---

## Tests Automatizados

```bash
python -m pytest        # 90 tests
python -m pytest -q     # modo silencioso
```

Estado: **90/90 passed** (cifrado, logica SQL, UI, exportacion).

---

## Compilar e Instalar

```bash
# Ejecutable
pyinstaller RegistroDoc.spec --clean

# Instalador Windows (Inno Setup)
& "C:\...\Inno Setup 6\ISCC.exe" RegistroDoc_Setup.iss
```

---

## Cumplimiento Legal

| Norma | Aplicacion |
|-------|-----------|
| Ley 81 / 2019 (Panama) | Datos cifrados localmente, sin envio a internet |
| ISO/IEC 27001:2022 A.8.24 | AES-256-GCM en reposo, log de auditoria inmutable |
| NIST SP 800-38D | GCM con nonce aleatorio por escritura |
| NIST SP 800-63B | PBKDF2-HMAC-SHA256, 600 000 iteraciones |

---

## Creditos

- **Autor Principal:** Docente Independiente (Panama)
- **Arquitectura SQL-First v5.0 & QA:** Antigravity AI (Google DeepMind)
