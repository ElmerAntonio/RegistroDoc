# 💼 RegistroDoc Pro — Guía de Ventas y Licencias

Guía operativa para vender el programa y gestionar los códigos de activación.
**Precio sugerido: B/. 20.00 por licencia anual** (meta: 50 docentes en la Comarca Ngäbe Buglé).

---

## 1. Cómo funciona la seguridad de las licencias (resumen)

Se usa **firma digital asimétrica (Ed25519)**, el modelo correcto para vender software:

| Clave | Dónde vive | Qué permite |
|---|---|---|
| 🔐 **Privada** (`clave_privada_vendedor.key`) | **SOLO** en tu carpeta `C:\RegistroDoc_Mis_Ventas\` | **Generar** códigos |
| 🔓 **Pública** | Embebida dentro del programa que entregas | **Solo verificar** — NO puede generar |

➡️ Un cliente con el programa **puede validar su código pero jamás fabricar uno**, porque no tiene la clave privada. Así se protegen tus ventas.

> ⚠️ **REGLAS DE ORO**
> - **NUNCA** compartas ni subas `clave_privada_vendedor.key`. Si se pierde, no podrás generar códigos nuevos que validen. Si se filtra, cualquiera podría generar códigos gratis.
> - **NUNCA** distribuyas `generador_codigos.py`. Es tu herramienta privada.
> - Haz una **copia de respaldo** de `clave_privada_vendedor.key` en un lugar seguro (USB / nube privada).

---

## 2. Generar los códigos (por lote, para usar después)

No necesitas generar un código en cada venta. **Pre-generas un lote y lo guardas en un documento.**

1. Abre el generador:
   ```
   C:\RegistroDoc_Mis_Ventas\generador_codigos.py
   ```
2. Pestaña **🏭 Generar lote** → elige la cantidad (ej. 20) → **Generar lote**.
3. Los códigos únicos se guardan en:
   - `licencias_generadas.txt` → **documento** para copiar/pegar cuando vendas.
   - `registro_ventas.enc` → registro cifrado (privado de tu equipo).
4. Cada código es **único e irrepetible** y tiene el formato `RD-XXXXXX-XXXXXX-...`.

## 3. Vender una licencia

1. En el generador, pestaña **📋 Registro / Ventas**, elige un código **DISPONIBLE**.
2. Clic en **💵 Marcar como vendido** → anota nombre, cédula, precio y teléfono.
3. Copia ese código y envíaselo al docente (WhatsApp/correo). Ejemplo de mensaje:
   > *"Su código de activación de RegistroDoc Pro es: RD-... . Ábralo, use el programa; cuando termine la prueba, ingrese el código en la pantalla de activación."*

## 4. Cómo lo activa el docente (cliente)

- Al instalar, el programa funciona **30 días de prueba** con TODAS las funciones (aparece "Versión de prueba: N días restantes").
- Al terminar la prueba, aparece la **pantalla de activación**: el docente pega su código y hace clic en **Activar programa**.
- El código se verifica **sin internet** (offline). Si es válido, queda activado en esa computadora.

---

## 5. Los instaladores

| Instalador (`Output/`) | Para qué |
|---|---|
| `RegistroDoc_Instalador.exe` | Venta normal, instalación limpia (con licencia + prueba 30 días) |
| `RegistroDoc_Instalador_ConDatos.exe` | Cuando el docente ya tiene datos previos (los preserva, no los borra) |
| `RegistroDoc_Instalador_Prueba.exe` | Para un **tester/beta**: prueba todo 30 días y llena `PLANTILLA_EXPERIENCIA_PRUEBA.md` |

Para regenerar los instaladores tras cambios de código:
1. `python -m PyInstaller RegistroDoc.spec --noconfirm --clean` (genera `dist/RegistroDoc.exe`).
2. Compilar cada `.iss` con Inno Setup (ISCC).

---

## 6. Preguntas frecuentes de ventas

- **¿Y si el docente cambia de computadora?** La activación queda en esa PC. Si cambia de equipo, entrégale de nuevo su código (o uno nuevo) para activar en la nueva.
- **¿Se puede usar el mismo código en dos PC?** Técnicamente sí (es offline). Por eso llevas el **registro** de a quién le vendiste cada código; para control estricto se necesitaría activación en línea (futuro).
- **¿Qué pasa si olvido activar y vence la prueba?** El programa pide el código; al ingresarlo, sigue funcionando con todos los datos intactos.
- **¿La licencia es anual?** El modelo previsto es anual (B/.20). La renovación por año es una mejora futura del sistema.

---

## 7. Checklist para empezar a vender

- [ ] Respaldar `clave_privada_vendedor.key` en lugar seguro.
- [ ] Generar un lote de códigos (ej. 20).
- [ ] Probar en TU PC: instalar, dejar vencer/activar con un código, confirmar que activa.
- [ ] Entregar `RegistroDoc_Instalador.exe` a los docentes.
- [ ] (Opcional) Entregar `RegistroDoc_Instalador_Prueba.exe` a un tester y recoger su documento de experiencia.
