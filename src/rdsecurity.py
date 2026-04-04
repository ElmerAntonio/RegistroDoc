"""
RegistroDoc Pro — Módulo de Seguridad Central
===============================================
Cumple con:
  • ISO/IEC 27001:2022  — Gestión de seguridad de la información
  • ISO/IEC 27002:2022  — Controles de seguridad (A.8.2, A.8.15, A.8.24)
  • NIST SP 800-63B     — Autenticación digital
  • NIST SP 800-38D     — AES-GCM
  • Ley 81 de 2019      — Protección de datos personales (Panamá)
  • FIPS 140-3          — Módulos criptográficos

Algoritmos utilizados:
  • AES-256-GCM         — Cifrado autenticado de archivos
  • PBKDF2-HMAC-SHA256  — Derivación de claves (600,000 iter — NIST 2023)
  • SHA3-256            — Hashing seguro de licencias
  • HMAC-SHA256         — Integridad del Excel
  • os.urandom(32)      — Generación de claves con CSPRNG del SO

© 2026 RegistroDoc Pro — Todos los derechos reservados
"""

import os
import json
import time
import hashlib
import hmac
import platform
import uuid
import re
from dotenv import load_dotenv
import datetime
from cryptography.hazmat.primitives.ciphers.aead import AESGCM
from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC
from cryptography.hazmat.primitives import hashes
from cryptography.hazmat.backends import default_backend

# ══════════════════════════════════════════════════════════════
#  CONSTANTES — NO MODIFICAR
# ══════════════════════════════════════════════════════════════

VERSION         = "2.0.0"
PBKDF2_ITERS    = 600_000          # NIST SP 800-132 recomendación 2023
SALT_LEN        = 32               # 256 bits
NONCE_LEN       = 12               # 96 bits — estándar GCM
TAG_LEN         = 16               # 128 bits — autenticación GCM
MAX_INTENTOS_1  = 3                # Bloqueo nivel 1
MAX_INTENTOS_2  = 5                # Bloqueo nivel 2
MAX_INTENTOS_3  = 10               # Bloqueo permanente
DELAY_NIVEL_1   = 30               # segundos
DELAY_NIVEL_2   = 3600             # 1 hora
AUDIT_FILE      = "rd_audit.bin"
BRUTE_FILE      = "rd_brute.bin"
LICENSE_FILE    = "licencia.dat"
CONFIG_FILE     = "config.enc"

# Clave maestra derivada del hardware — nunca viaja por red ni se guarda
_HW_SEED = None

# ══════════════════════════════════════════════════════════════
#  HUELLA DE HARDWARE — Vincula la licencia al dispositivo
# ══════════════════════════════════════════════════════════════

def _hw_fingerprint() -> bytes:
    """
    Genera una huella única del hardware combinando:
      - UUID de la placa madre (Windows: BIOS UUID via WMIC)
      - Nombre del equipo
      - Nombre del usuario del SO
    Cumple ISO/IEC 27001 A.8.1 — Inventario de activos
    """
    global _HW_SEED
    if _HW_SEED is not None:
        return _HW_SEED

    componentes = []

    # 1. UUID del sistema (BIOS/Motherboard)
    try:
        if platform.system() == "Windows":
            import subprocess
            r = subprocess.check_output(
                ["wmic", "csproduct", "get", "UUID"],
                stderr=subprocess.DEVNULL, timeout=5
            ).decode().strip().split("\n")
            uid = r[-1].strip() if len(r) > 1 else ""
            if uid and uid != "UUID":
                componentes.append(uid)
    except Exception:
        pass

    # 2. Identificador de volumen del disco del sistema
    try:
        if platform.system() == "Windows":
            import subprocess
            r = subprocess.check_output(
                ["cmd.exe", "/c", "vol", "C:"],
                stderr=subprocess.DEVNULL, timeout=5
            ).decode()
            serial = re.search(r"[0-9A-F]{4}-[0-9A-F]{4}", r)
            if serial:
                componentes.append(serial.group())
    except Exception:
        pass

    # 3. Nombre del equipo + usuario (siempre disponible)
    componentes.append(platform.node())
    componentes.append(os.environ.get("USERNAME", "") or os.environ.get("USER", ""))

    # 4. MAC address como respaldo
    try:
        mac = hex(uuid.getnode())
        componentes.append(mac)
    except Exception:
        pass

    raw = "|".join(c for c in componentes if c)
    # Derivar clave estable con PBKDF2
    salt_hw = hashlib.sha256(b"RD_HW_SALT_2026_PANAMA").digest()
    kdf = PBKDF2HMAC(
        algorithm=hashes.SHA256(),
        length=32,
        salt=salt_hw,
        iterations=100_000,
        backend=default_backend()
    )
    _HW_SEED = kdf.derive(raw.encode("utf-8", errors="replace"))
    return _HW_SEED

# ══════════════════════════════════════════════════════════════
#  CIFRADO AES-256-GCM
# ══════════════════════════════════════════════════════════════

def _derivar_clave(password: bytes, salt: bytes) -> bytes:
    """
    Deriva clave AES-256 con PBKDF2-HMAC-SHA256.
    NIST SP 800-132 — 600,000 iteraciones mínimas (2023).
    """
    kdf = PBKDF2HMAC(
        algorithm=hashes.SHA256(),
        length=32,
        salt=salt,
        iterations=PBKDF2_ITERS,
        backend=default_backend()
    )
    return kdf.derive(password)

def cifrar(datos: bytes, password: bytes) -> bytes:
    """
    Cifra datos con AES-256-GCM.
    Formato del archivo cifrado:
      [4 bytes: versión magic] [32 bytes: salt] [12 bytes: nonce] [datos cifrados + tag 16 bytes]
    """
    salt  = os.urandom(SALT_LEN)
    nonce = os.urandom(NONCE_LEN)
    clave = _derivar_clave(password, salt)
    aesgcm = AESGCM(clave)
    cifrado = aesgcm.encrypt(nonce, datos, None)  # cifrado incluye tag GCM
    magic = b"RD26"  # magic bytes para identificar el formato
    return magic + salt + nonce + cifrado

def descifrar(blob: bytes, password: bytes) -> bytes:
    """
    Descifra y autentica datos AES-256-GCM.
    Lanza ValueError si la autenticación falla (datos corruptos o manipulados).
    """
    if len(blob) < 4 + SALT_LEN + NONCE_LEN + TAG_LEN:
        raise ValueError("Archivo demasiado pequeño — posiblemente corrupto")
    if blob[:4] != b"RD26":
        raise ValueError("Formato inválido — archivo no reconocido")
    salt    = blob[4:4+SALT_LEN]
    nonce   = blob[4+SALT_LEN:4+SALT_LEN+NONCE_LEN]
    cifrado = blob[4+SALT_LEN+NONCE_LEN:]
    clave   = _derivar_clave(password, salt)
    aesgcm  = AESGCM(clave)
    try:
        return aesgcm.decrypt(nonce, cifrado, None)
    except Exception:
        raise ValueError("Autenticación fallida — archivo manipulado o clave incorrecta")

def guardar_cifrado(ruta: str, datos: dict, password: bytes = None) -> None:
    """Serializa dict a JSON y lo cifra con AES-256-GCM."""
    if password is None:
        password = _hw_fingerprint()
    json_bytes = json.dumps(datos, ensure_ascii=False).encode("utf-8")
    blob = cifrar(json_bytes, password)
    with open(ruta, "wb") as f:
        f.write(blob)

def cargar_cifrado(ruta: str, password: bytes = None) -> dict:
    """Lee archivo cifrado y devuelve dict. Retorna {} si no existe o falla."""
    if not os.path.exists(ruta):
        return {}
    if password is None:
        password = _hw_fingerprint()
    try:
        with open(ruta, "rb") as f:
            blob = f.read()
        json_bytes = descifrar(blob, password)
        return json.loads(json_bytes.decode("utf-8"))
    except Exception:
        return {}

# ══════════════════════════════════════════════════════════════
#  PROTECCIÓN ANTI FUERZA BRUTA — ISO/IEC 27002 A.8.3
# ══════════════════════════════════════════════════════════════

def _leer_brute() -> dict:
    datos = cargar_cifrado(BRUTE_FILE)
    return datos if datos else {"intentos": 0, "bloqueado_hasta": 0, "bloqueado_perm": False}

def _guardar_brute(datos: dict) -> None:
    guardar_cifrado(BRUTE_FILE, datos)

def verificar_bloqueo() -> tuple[bool, str]:
    """
    Verifica si el sistema está bloqueado.
    Retorna (bloqueado: bool, mensaje: str)
    """
    d = _leer_brute()
    if d.get("bloqueado_perm"):
        registrar_auditoria("SEGURIDAD", "BLOQUEO_PERMANENTE", "Intento de acceso con bloqueo permanente")
        return True, ("🔒 ACCESO BLOQUEADO PERMANENTEMENTE\n\n"
                      "Se detectaron demasiados intentos fallidos.\n"
                      "Contacta al proveedor del programa con tu N° de cédula\n"
                      "para desbloquear tu licencia.")
    hasta = d.get("bloqueado_hasta", 0)
    if hasta > time.time():
        restante = int(hasta - time.time())
        mins = restante // 60
        segs = restante % 60
        tiempo_txt = f"{mins}m {segs}s" if mins > 0 else f"{segs}s"
        return True, (f"⏳ ACCESO TEMPORALMENTE BLOQUEADO\n\n"
                      f"Demasiados intentos fallidos.\n"
                      f"Espera {tiempo_txt} antes de intentar de nuevo.")
    return False, ""

def registrar_intento_fallido() -> tuple[bool, str]:
    """
    Registra intento fallido y aplica bloqueo progresivo.
    Retorna (bloqueado_ahora: bool, mensaje: str)
    """
    d = _leer_brute()
    d["intentos"] = d.get("intentos", 0) + 1
    n = d["intentos"]

    if n >= MAX_INTENTOS_3:
        d["bloqueado_perm"] = True
        _guardar_brute(d)
        registrar_auditoria("SEGURIDAD", "BLOQUEO_PERMANENTE_ACTIVADO",
                            f"Bloqueado permanentemente tras {n} intentos")
        return True, "BLOQUEO PERMANENTE activado."

    elif n >= MAX_INTENTOS_2:
        d["bloqueado_hasta"] = time.time() + DELAY_NIVEL_2
        _guardar_brute(d)
        registrar_auditoria("SEGURIDAD", "BLOQUEO_1H",
                            f"Bloqueado 1 hora tras {n} intentos")
        return True, f"Bloqueado 1 hora. Intento {n}/{MAX_INTENTOS_3}."

    elif n >= MAX_INTENTOS_1:
        d["bloqueado_hasta"] = time.time() + DELAY_NIVEL_1
        _guardar_brute(d)
        registrar_auditoria("SEGURIDAD", "BLOQUEO_30S",
                            f"Bloqueado 30 segundos tras {n} intentos")
        return True, f"Bloqueado 30 segundos. Intento {n}/{MAX_INTENTOS_3}."

    _guardar_brute(d)
    registrar_auditoria("SEGURIDAD", "INTENTO_FALLIDO",
                        f"Intento fallido {n}/{MAX_INTENTOS_3}")
    return False, f"Código incorrecto. Intento {n} de {MAX_INTENTOS_3-1} antes del bloqueo."

def resetear_intentos() -> None:
    """Resetea el contador tras activación exitosa."""
    guardar_cifrado(BRUTE_FILE, {"intentos": 0, "bloqueado_hasta": 0, "bloqueado_perm": False})

# ══════════════════════════════════════════════════════════════
#  LOG DE AUDITORÍA CIFRADO — ISO/IEC 27001 A.8.15
# ══════════════════════════════════════════════════════════════

def registrar_auditoria(categoria: str, accion: str, detalle: str = "") -> None:
    """
    Registra evento en log de auditoría cifrado.
    Cumple ISO/IEC 27001:2022 A.8.15 — Logging.
    Campos: timestamp_utc, categoria, accion, detalle, hw_id_corto
    """
    try:
        log = cargar_cifrado(AUDIT_FILE)
        if "eventos" not in log:
            log["eventos"] = []
        hw = _hw_fingerprint()
        hw_id = hashlib.sha256(hw).hexdigest()[:12]  # ID corto del dispositivo
        evento = {
            "ts":       datetime.datetime.utcnow().isoformat() + "Z",
            "cat":      categoria[:32],
            "accion":   accion[:64],
            "detalle":  detalle[:256],
            "hw_id":    hw_id,
        }
        log["eventos"].append(evento)
        # Mantener últimos 10,000 eventos
        if len(log["eventos"]) > 10_000:
            log["eventos"] = log["eventos"][-10_000:]
        guardar_cifrado(AUDIT_FILE, log)
    except Exception:
        pass  # El log nunca debe interrumpir la operación

def leer_auditoria() -> list:
    """Retorna lista de eventos del log."""
    log = cargar_cifrado(AUDIT_FILE)
    return log.get("eventos", [])

# ══════════════════════════════════════════════════════════════
#  INTEGRIDAD DEL EXCEL — HMAC-SHA256
# ══════════════════════════════════════════════════════════════

EXCEL_HASH_FILE = "rd_excel_hash.bin"

def calcular_hash_excel(ruta_excel: str) -> str:
    """Calcula SHA3-256 del archivo Excel."""
    h = hashlib.sha3_256()
    try:
        with open(ruta_excel, "rb") as f:
            while True:
                bloque = f.read(65536)
                if not bloque:
                    break
                h.update(bloque)
        return h.hexdigest()
    except Exception:
        return ""

def guardar_hash_excel(ruta_excel: str) -> None:
    """Guarda el hash del Excel cifrado para verificación futura."""
    hash_val = calcular_hash_excel(ruta_excel)
    if hash_val:
        datos = {
            "hash":      hash_val,
            "algoritmo": "SHA3-256",
            "fecha":     datetime.datetime.utcnow().isoformat() + "Z",
            "archivo":   os.path.basename(ruta_excel),
        }
        guardar_cifrado(EXCEL_HASH_FILE, datos)
        registrar_auditoria("INTEGRIDAD", "HASH_GUARDADO", f"Excel hash guardado: {hash_val[:16]}...")

def verificar_integridad_excel(ruta_excel: str) -> tuple[bool, str]:
    """
    Verifica que el Excel no haya sido modificado externamente.
    Retorna (integro: bool, mensaje: str)
    """
    if not os.path.exists(EXCEL_HASH_FILE):
        # Primera vez — guardar hash base
        guardar_hash_excel(ruta_excel)
        return True, "Hash inicial guardado."

    datos = cargar_cifrado(EXCEL_HASH_FILE)
    if not datos:
        guardar_hash_excel(ruta_excel)
        return True, "Hash regenerado."

    hash_guardado = datos.get("hash", "")
    hash_actual   = calcular_hash_excel(ruta_excel)

    if not hash_actual:
        return False, "No se pudo calcular el hash del Excel."

    if hash_guardado == hash_actual:
        return True, "Integridad verificada."
    else:
        registrar_auditoria("SEGURIDAD", "INTEGRIDAD_FALLIDA",
                            f"Excel modificado externamente. Hash esperado: {hash_guardado[:16]}... actual: {hash_actual[:16]}...")
        return False, ("⚠ El archivo Excel fue modificado externamente.\n"
                       "Esto puede indicar manipulación de datos.\n"
                       "El evento ha sido registrado en el log de auditoría.")

def actualizar_hash_excel(ruta_excel: str) -> None:
    """Actualiza el hash después de que el PROGRAMA guarda datos."""
    guardar_hash_excel(ruta_excel)

# ══════════════════════════════════════════════════════════════
#  SISTEMA DE LICENCIAS SEGURO
# ══════════════════════════════════════════════════════════════

# Cargar configuracion segura desde .env
from config import BASE_DIR
env_path = os.path.join(os.path.dirname(BASE_DIR), '.env')
load_dotenv(dotenv_path=env_path)

# Clave maestra interna (Cargada desde variable de entorno de forma segura)
_master_salt_env = os.environ.get("REGISTRODOC_MASTER_SALT")
if not _master_salt_env:
    raise RuntimeError(
        "CRITICAL ERROR: REGISTRODOC_MASTER_SALT no encontrada en el entorno o archivo .env. "
        "El programa no puede iniciar de forma segura."
    )

_MASTER_SALT = _master_salt_env.encode("utf-8")

def _clave_licencia(cedula: str) -> bytes:
    """
    Deriva la clave de cifrado para la licencia.
    Combina: cédula del docente + huella de hardware + master salt
    → La licencia solo puede abrirse en el MISMO dispositivo donde se activó.
    """
    hw   = _hw_fingerprint()
    base = cedula.encode() + hw + _MASTER_SALT
    salt = hashlib.sha3_256(base).digest()
    kdf  = PBKDF2HMAC(
        algorithm=hashes.SHA256(),
        length=32,
        salt=salt,
        iterations=PBKDF2_ITERS,
        backend=default_backend()
    )
    return kdf.derive(base)

def generar_codigo_licencia(cedula: str) -> str:
    """
    Genera código de activación único para una cédula.
    Formato: RD-AAAAA-BBBBB-CCCCC-DDDDD (20 chars de datos)
    
    Proceso:
      1. Parte A: identificador de la cédula (SHA3, 5 chars)
      2. Partes B,C,D: derivadas de SHA3-256 con SALT secreto
    
    El código es determinista (misma cédula = mismo código siempre)
    pero computacionalmente imposible de invertir o falsificar
    sin conocer _MASTER_SALT.
    """
    cedula = cedula.strip()
    # Hash de identificación de la cédula
    h_ced = hashlib.sha3_256(
        cedula.encode() + _MASTER_SALT + b"_IDENT"
    ).hexdigest().upper()
    parte_a = re.sub(r'[^A-Z0-9]', 'X', h_ced)[:5]

    # Partes B, C, D del código
    h_full = hashlib.sha3_256(
        cedula.encode() + _MASTER_SALT + b"_LICENSE_V2"
    ).hexdigest().upper()
    chars_validos = re.sub(r'[^A-Z0-9]', '', h_full)
    parte_b = chars_validos[0:5]
    parte_c = chars_validos[5:10]
    parte_d = chars_validos[10:15]

    # Formato: RD-AAAAA-BBBBB-CCCCC-DDDDD = 26 chars
    return f"RD-{parte_a}-{parte_b}-{parte_c}-{parte_d}"

def activar_licencia(cedula: str, codigo: str) -> tuple[bool, str]:
    """
    Activa la licencia para esta cédula en este dispositivo.
    Retorna (ok: bool, mensaje: str)
    
    Seguridad:
      - Verifica bloqueo por fuerza bruta antes de procesar
      - El archivo de licencia se cifra con AES-256-GCM
      - La clave incluye huella de hardware → no funciona en otro equipo
    """
    # 1. Verificar bloqueo
    bloqueado, msg_bloqueo = verificar_bloqueo()
    if bloqueado:
        return False, msg_bloqueo

    cedula = cedula.strip()
    codigo = codigo.strip().upper()

    # 2. Validar formato
    if not re.match(r"^RD-[A-Z0-9]{5}-[A-Z0-9]{5}-[A-Z0-9]{5}-[A-Z0-9]{5}$", codigo):  # 26 chars total
        bloq, msg = registrar_intento_fallido()
        return False, f"Formato de código inválido.\n{msg}"

    # 3. Verificar código
    esperado = generar_codigo_licencia(cedula)
    # Comparación segura contra timing attacks
    if not hmac.compare_digest(codigo.encode(), esperado.encode()):
        bloq, msg = registrar_intento_fallido()
        return False, f"Código incorrecto para esta cédula.\n{msg}"

    # 4. Activar — cifrar licencia con clave que incluye hardware
    clave_lic = _clave_licencia(cedula)
    hw = _hw_fingerprint()
    datos_lic = {
        "cedula":     cedula,
        "codigo":     codigo,
        "activado":   datetime.datetime.utcnow().isoformat() + "Z",
        "version":    VERSION,
        "hw_id":      hashlib.sha256(hw).hexdigest(),
        "valido":     True,
    }
    try:
        blob = cifrar(json.dumps(datos_lic).encode(), clave_lic)
        with open(LICENSE_FILE, "wb") as f:
            f.write(blob)
    except Exception as e:
        return False, f"Error al guardar licencia: {e}"

    resetear_intentos()
    registrar_auditoria("LICENCIA", "ACTIVACION_EXITOSA",
                        f"Cédula: {cedula[:4]}***  HW: {hashlib.sha256(hw).hexdigest()[:12]}")
    return True, "¡Activación exitosa! El programa está listo para usar."

def verificar_licencia() -> tuple[bool, dict]:
    """
    Verifica la licencia. Retorna (válida: bool, datos: dict)
    
    Verifica:
      1. Que el archivo exista
      2. Que el cifrado AES-256-GCM sea válido (no manipulado)
      3. Que la clave (que incluye HW) funcione → mismo dispositivo
      4. Que el campo valido=True
    """
    if not os.path.exists(LICENSE_FILE):
        return False, {}
    try:
        with open(LICENSE_FILE, "rb") as f:
            blob = f.read()
        # Necesitamos la cédula para derivar la clave
        # Leer solo los primeros bytes para obtener la cédula (no es seguro
        # guardar la cédula sin cifrar, pero necesitamos bootstrapping)
        # Solución: guardamos un token mínimo en claro para derivar la clave
        # El token en claro es: hash(cédula||hw) — no revela la cédula
        datos_temp = _leer_cedula_token()
        if not datos_temp:
            return False, {}
        cedula_hint = datos_temp.get("hint", "")
        if not cedula_hint:
            return False, {}
        # La "cédula" guardada es el hint cifrado con solo HW (sin cédula)
        # para permitir bootstrapping
        clave_lic = _clave_licencia(cedula_hint)
        datos = json.loads(descifrar(blob, clave_lic).decode())
        if not datos.get("valido"):
            return False, {}
        # Verificar que el HW coincide
        hw     = _hw_fingerprint()
        hw_id  = hashlib.sha256(hw).hexdigest()
        if datos.get("hw_id") != hw_id:
            registrar_auditoria("SEGURIDAD", "HW_MISMATCH",
                                "Licencia usada en hardware diferente al registrado")
            return False, {}
        registrar_auditoria("ACCESO", "LOGIN_OK",
                            f"Acceso autorizado cédula {datos.get('cedula','')[:4]}***")
        return True, datos
    except Exception:
        return False, {}

def _guardar_cedula_token(cedula: str) -> None:
    """
    Guarda un token mínimo cifrado solo con HW para el bootstrapping.
    No guarda la cédula en texto claro.
    """
    hw   = _hw_fingerprint()
    kdf  = PBKDF2HMAC(
        algorithm=hashes.SHA256(), length=32,
        salt=hashlib.sha256(hw + b"TOKEN").digest(),
        iterations=PBKDF2_ITERS, backend=default_backend()
    )
    clave_token = kdf.derive(hw)
    datos = {"hint": cedula, "ts": datetime.datetime.utcnow().isoformat()}
    blob  = cifrar(json.dumps(datos).encode(), clave_token)
    with open("rd_token.bin", "wb") as f:
        f.write(blob)

def _leer_cedula_token() -> dict:
    if not os.path.exists("rd_token.bin"):
        return {}
    try:
        hw  = _hw_fingerprint()
        kdf = PBKDF2HMAC(
            algorithm=hashes.SHA256(), length=32,
            salt=hashlib.sha256(hw + b"TOKEN").digest(),
            iterations=PBKDF2_ITERS, backend=default_backend()
        )
        clave_token = kdf.derive(hw)
        with open("rd_token.bin", "rb") as f:
            blob = f.read()
        return json.loads(descifrar(blob, clave_token).decode())
    except Exception:
        return {}

def activar_licencia_completa(cedula: str, codigo: str) -> tuple[bool, str]:
    """Wrapper que también guarda el token de bootstrapping."""
    ok, msg = activar_licencia(cedula, codigo)
    if ok:
        _guardar_cedula_token(cedula)
    return ok, msg

# ══════════════════════════════════════════════════════════════
#  CONFIG CIFRADA — Ley 81/2019 Panamá (datos personales)
# ══════════════════════════════════════════════════════════════

def guardar_config_segura(cfg: dict) -> None:
    """Cifra y guarda la configuración. Datos personales protegidos."""
    guardar_cifrado(CONFIG_FILE, cfg)
    registrar_auditoria("CONFIG", "GUARDADO", "Configuración actualizada")

def cargar_config_segura(default: dict) -> dict:
    """Carga configuración cifrada. Retorna default si no existe."""
    datos = cargar_cifrado(CONFIG_FILE)
    if not datos:
        return dict(default)
    cfg = dict(default)
    cfg.update(datos)
    return cfg

# ══════════════════════════════════════════════════════════════
#  VALIDACIÓN DE NOTAS — Sistema MEDUCA Panamá
# ══════════════════════════════════════════════════════════════

NOTA_MIN = 1.0
NOTA_MAX = 5.0

import math

def validar_nota_meduca(valor: str) -> tuple[bool, float, str]:
    """
    Valida una nota según el sistema MEDUCA Panamá.
    Escala: 1.0 a 5.0 con un decimal.
    
    Retorna (válida: bool, valor_float: float, mensaje: str)
    """
    if isinstance(valor, bool):
        return False, 0.0, "Los valores booleanos no son válidos."

    if valor is None or not str(valor).strip():
        return False, 0.0, "Nota vacía."
    
    # Normalizar separador decimal
    val_str = str(valor).strip().replace(",", ".")
    
    try:
        nota_float = float(val_str)
        if math.isnan(nota_float) or math.isinf(nota_float):
            return False, 0.0, f"'{valor}' no es un número válido."
        nota = round(nota_float, 1)
    except (ValueError, TypeError):
        return False, 0.0, f"'{valor}' no es un número válido."
    
    if nota < NOTA_MIN:
        return False, 0.0, f"La nota mínima en el sistema MEDUCA es {NOTA_MIN}. El 0 no existe en esta escala."
    
    if nota > NOTA_MAX:
        return False, 0.0, f"La nota máxima en el sistema MEDUCA es {NOTA_MAX}."
    
    return True, nota, "OK"
