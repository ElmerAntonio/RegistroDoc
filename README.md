# RegistroDoc Pro v3.0 — Sistema de Registro Académico (Panamá)

[![Python](https://img.shields.io/badge/Python-3.10%2B-blue.svg)](https://www.python.org/)
[![License](https://img.shields.io/badge/Licencia-Comercial-green.svg)](#)
[![Security Standards](https://img.shields.io/badge/Security-ISO%2027001%20%7C%20Ley%2081-darkblue.svg)](#)
[![Tests Status](https://img.shields.io/badge/Tests-73%2F73%20Passed-emerald.svg)](#)

**RegistroDoc Pro** es una aplicación de escritorio premium e independiente desarrollada en Python, diseñada específicamente para optimizar la labor administrativa de docentes en la República de Panamá (tanto en modalidad **Primaria** como **Premedia**). 

El programa automatiza e interactúa de manera directa con las libretas de registro oficial de calificaciones en formato Excel del **Ministerio de Educación (MEDUCA)**, sin alterar sus fórmulas ni su diseño oficial.

---

## 🚀 Características Clave

* **Diseño Bimodal Dinámico**: Soporte integrado para las libretas oficiales de calificaciones y asistencia de Primaria (con casillas continuas) y Premedia.
* **Operación 100% Fuera de Línea (Offline)**: Diseñado para funcionar en áreas remotas o comarcales (como la Comarca Ngäbe Buglé) sin necesidad de acceso a Internet.
* **Seguridad y Cifrado Avanzado**: Cumple con la **Ley 81 de 2019 (Panamá)** sobre protección de datos personales y estándares **NIST SP 800-38D** usando cifrado **AES-256-GCM** para asegurar la información del docente.
* **Controles de Auditoría e Integridad**: Implementa un registro de auditoría seguro cifrado (`data/rd_audit.bin`) e integridad del archivo Excel mediante firmas SHA3-256 (alineado con la norma **ISO/IEC 27001:2022**).
* **Política de Backups Inteligente**: Sistema de copias de seguridad automáticas y periódicas con un límite de retención estricto de **máximo 3 archivos** para evitar el desgaste del almacenamiento.
* **Arranque Rápido de Prueba (Instant Boot)**: El sistema inicia directamente en un perfil de demostración autogenerado, permitiendo evaluar la funcionalidad del panel de control inmediatamente y sin obligar a registrar credenciales al abrir la aplicación por primera vez.

---

## 📂 Estructura del Proyecto (Normas ISO)

El repositorio está organizado siguiendo normas internacionales de ingeniería de software para mantener un espacio de trabajo limpio, escalable y mantenible:

```text
RegistroDoc/
├── assets/                    # Recursos visuales (logotipos, iconos, splash de inicio)
├── data/                      # Archivos de configuración y bases de datos offline (cifrados)
├── dev/                       # Utilidades de desarrollo, scripts de limpieza y empaquetador
├── docs/                      # Documentación del proyecto (manuales de usuario oficiales)
├── src/                       # Código fuente de la aplicación en Python
│   ├── utils/                 # Funciones auxiliares y herramientas comunes (diálogos, fechas)
│   ├── app.py                 # Punto de entrada de la aplicación y vista contenedora
│   └── rdsecurity.py          # Módulo criptográfico y validación de licencias
├── tests/                     # Suite de pruebas automatizadas unitarias y de integración
├── Iniciar_RegistroDoc.bat    # Asistente de arranque rápido para Windows
├── Registro Primaria.xlsx     # Plantilla base oficial para nivel Primaria
├── Registro_2026.xlsx         # Plantilla base oficial para nivel Premedia
├── requirements.txt           # Dependencias de librerías de Python
└── pytest.ini                 # Configuración del motor de pruebas unitarias
```

---

## 🛠️ Instalación y Configuración

### Prerrequisitos
* **Python 3.10** o superior instalado.
* Sistema Operativo **Windows 10** o **Windows 11**.

### Instalación de Dependencias
Abre una terminal en la raíz del proyecto y ejecuta:
```bash
pip install -r requirements.txt
```

### Ejecutar la Aplicación
Puedes arrancar la aplicación de dos formas:

1. **Mediante el script automatizado (Recomendado)**:
   Haz doble clic en `Iniciar_RegistroDoc.bat`. Este script verificará automáticamente tu instalación de Python, instalará las dependencias necesarias y abrirá la aplicación.

2. **Mediante la terminal**:
   ```bash
   python src/app.py
   ```

---

## 🧪 Pruebas Automatizadas

El proyecto incluye una amplia suite de pruebas unitarias y de integración que validan el motor de datos, el flujo de cifrado, la integridad del Excel y el comportamiento de la UI (incluyendo soporte headless/sin pantalla):

Para correr las pruebas unitarias:
```bash
python -m pytest
```

---

## 📦 Empaquetado de Distribución limpia

El directorio `dev/` incluye un script automatizado para empaquetar una distribución limpia en formato `.zip` que excluye archivos de entorno locales, cachés de compilación o respaldos temporales:

Para generar el paquete de distribución limpia:
```bash
python dev/recreate_zip.py
```
El archivo resultante se guardará en la raíz del proyecto como `RegistroDoc_Limpio.zip`.

---

## 🔒 Cumplimiento Legal y Normativo

1. **Ley 81 del 26 de marzo de 2019 (Panamá)**: Toda información personal e institucional (cédula, nombres, notas de alumnos) es cifrada de forma segura localmente.
2. **ISO/IEC 27001:2022 A.8.15 & A.8.24**: Logging seguro cifrado de todas las transacciones académicas y del sistema.
3. **NIST SP 800-63B / SP 800-38D**: Autenticación criptográfica y derivación de llaves por dispositivo mediante PBKDF2-HMAC-SHA256 con 600,000 iteraciones.

---

## 👥 Desarrolladores y Créditos

* **Autor Principal:** Creador Independiente (Panamá)
* **Ingeniería de Software e Integración de Calidad (V3.0):** Antigravity AI (Google Deepmind)
