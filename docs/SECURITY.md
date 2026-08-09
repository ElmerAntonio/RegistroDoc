# Política de Seguridad y Calidad (Security & Quality Policy)

¡Gracias por tomarte el tiempo de revisar la seguridad de **RegistroDoc / MARC**! La seguridad y la integridad de los datos escolares son nuestra máxima prioridad.

## Versiones Soportadas

Actualmente, solo proporcionamos actualizaciones de seguridad y parches de calidad para la versión principal más reciente del sistema en producción.

| Versión | Soporte de Seguridad |
| ------- | ------------------ |
| 5.0.x   | :white_check_mark: |
| < 5.0   | :x:                |

## Reportar una Vulnerabilidad

Dado que reviso y administro este repositorio de forma activa y diaria, gestiono todos los reportes de seguridad directamente a través de GitHub para mantener el flujo de trabajo centralizado. **No es necesario enviar correos electrónicos.**

Si descubres una vulnerabilidad de seguridad o un problema crítico de calidad en el código, por favor repórtalo siguiendo uno de estos pasos:

**Opción 1: Reporte Privado (Recomendado)**
1. Ve a la pestaña **Security** de este repositorio.
2. Selecciona **Advisories** en el menú lateral izquierdo.
3. Haz clic en el botón **Report a vulnerability**.
4. Describe el problema y los pasos para reproducirlo. Este reporte será privado y solo yo podré verlo.

**Opción 2: Mediante un Issue**
Si la Opción 1 no está disponible, por favor abre un nuevo **Issue** en el repositorio. 
* Usa el título: `[SECURITY]` seguido de una breve descripción.
* No compartas datos reales de escuelas o contraseñas en el Issue.

### ¿Qué puedes esperar?
* **Tiempo de respuesta:** Reviso GitHub constantemente, por lo que acusaré recibo de tu reporte en un plazo máximo de 24 a 48 horas.
* **Proceso:** Analizaré el problema, confirmaré si es reproducible y te mantendré informado sobre el progreso del parche directamente en el hilo de GitHub.
* Una vez que la vulnerabilidad sea parcheada, se lanzará una actualización en la rama principal.

---

## Sistema de Licencias (actualización 2026-08)

RegistroDoc Pro usa **firma digital asimétrica Ed25519** para las licencias:

- **Clave privada** (firma/genera códigos): vive **solo** en la herramienta del vendedor (`C:\RegistroDoc_Mis_Ventas\`), nunca se distribuye.
- **Clave pública** (verifica): embebida en `src/rdlicense.py`. Solo permite verificar; **no** puede generar códigos.
- Los códigos de activación se verifican **offline** (sin internet) y son **imposibles de falsificar** sin la clave privada.
- El programa incluye un **período de prueba de 30 días**; luego exige el código de activación.
- La verificación de licencia **falla-abierto** ante errores para no bloquear al docente por un fallo técnico.

Ver `docs/GUIA_VENTAS_Y_LICENCIAS.md` para la operación de ventas.
