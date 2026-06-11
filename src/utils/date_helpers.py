import datetime
from rdsecurity import cargar_config_segura

MESES = {
    "ENERO": 1, "FEBRERO": 2, "MARZO": 3, "ABRIL": 4, "MAYO": 5, "JUNIO": 6,
    "JULIO": 7, "AGOSTO": 8, "SEPTIEMBRE": 9, "OCTUBRE": 10, "NOVIEMBRE": 11, "DICIEMBRE": 12
}

def parse_rango_fecha(str_val, ano_lectivo):
    """
    Parsea una cadena como '06 DE MARZO AL 02 DE JUNIO' o '06/03/2026 al 02/06/2026'.
    Retorna tupla (start_date, end_date) o None.
    """
    if not str_val:
        return None
    val = str(str_val).upper().strip()
    parts = val.split(" AL ")
    if len(parts) != 2:
        parts = val.split(" - ")
    if len(parts) != 2:
        parts = val.split(" A ")
    if len(parts) != 2:
        return None
        
    try:
        ano = int(ano_lectivo)
    except Exception:
        ano = datetime.date.today().year

    def parse_single_date(part):
        part = part.strip()
        # Caso DD/MM/YYYY o DD/MM
        if "/" in part:
            sub = part.split("/")
            if len(sub) >= 2:
                try:
                    d = int(sub[0])
                    m = int(sub[1])
                    return datetime.date(ano, m, d)
                except Exception:
                    pass
        # Caso DD-MM-YYYY o DD-MM
        if "-" in part:
            sub = part.split("-")
            if len(sub) >= 2:
                try:
                    d = int(sub[0])
                    m = int(sub[1])
                    return datetime.date(ano, m, d)
                except Exception:
                    pass
        # Caso texto en español, ej: '06 DE MARZO'
        for mes_name, mes_num in MESES.items():
            if mes_name in part:
                words = part.replace("DE", " ").split()
                for w in words:
                    try:
                        d = int(w)
                        return datetime.date(ano, mes_num, d)
                    except ValueError:
                        continue
        return None

    d_start = parse_single_date(parts[0])
    d_end = parse_single_date(parts[1])
    if d_start and d_end:
        return d_start, d_end
    return None

def obtener_trimestre_actual(fecha_actual=None):
    """
    Calcula el trimestre actual (Trimestre 1, 2 o 3) según la fecha provista
    y los rangos de la configuración.
    """
    if fecha_actual is None:
        fecha_actual = datetime.date.today()
    elif isinstance(fecha_actual, str):
        # Intentar parsear 'DD-MM-YYYY'
        try:
            partes = fecha_actual.split("-")
            fecha_actual = datetime.date(int(partes[2]), int(partes[1]), int(partes[0]))
        except Exception:
            fecha_actual = datetime.date.today()

    cfg = cargar_config_segura({})
    ano_lectivo = cfg.get("ano_lectivo", str(fecha_actual.year))
    
    # Valores por defecto de Panamá (Escolar)
    try:
        ano = int(ano_lectivo)
    except Exception:
        ano = fecha_actual.year
        
    t1_start = datetime.date(ano, 3, 1)
    t1_end = datetime.date(ano, 6, 2)
    t2_start = datetime.date(ano, 6, 3)
    t2_end = datetime.date(ano, 9, 8)
    t3_start = datetime.date(ano, 9, 9)
    t3_end = datetime.date(ano, 12, 20)

    # Reemplazar por fechas en config si están disponibles y válidas
    p1 = parse_rango_fecha(cfg.get("fecha_t1", ""), ano_lectivo)
    if p1: t1_start, t1_end = p1
    
    p2 = parse_rango_fecha(cfg.get("fecha_t2", ""), ano_lectivo)
    if p2: t2_start, t2_end = p2
    
    p3 = parse_rango_fecha(cfg.get("fecha_t3", ""), ano_lectivo)
    if p3: t3_start, t3_end = p3

    # Comparar
    if t1_start <= fecha_actual <= t1_end:
        return "Trimestre 1"
    elif t2_start <= fecha_actual <= t2_end:
        return "Trimestre 2"
    elif t3_start <= fecha_actual <= t3_end:
        return "Trimestre 3"
        
    # Fallback aproximado por meses
    if 3 <= fecha_actual.month <= 5:
        return "Trimestre 1"
    elif 6 <= fecha_actual.month <= 8:
        return "Trimestre 2"
    else:
        return "Trimestre 3"

def obtener_rango_fechas_trimestre(trimestre_str, ano_lectivo=None):
    """
    Retorna la tupla (start_date, end_date) para un trimestre dado.
    """
    if ano_lectivo is None:
        cfg = cargar_config_segura({})
        ano_lectivo = cfg.get("ano_lectivo", str(obtener_hoy_panama().year))
    try:
        ano = int(ano_lectivo)
    except Exception:
        ano = obtener_hoy_panama().year
        
    t1_start = datetime.date(ano, 3, 1)
    t1_end = datetime.date(ano, 6, 2)
    t2_start = datetime.date(ano, 6, 3)
    t2_end = datetime.date(ano, 9, 8)
    t3_start = datetime.date(ano, 9, 9)
    t3_end = datetime.date(ano, 12, 20)

    cfg = cargar_config_segura({})
    p1 = parse_rango_fecha(cfg.get("fecha_t1", ""), ano_lectivo)
    if p1: t1_start, t1_end = p1
    p2 = parse_rango_fecha(cfg.get("fecha_t2", ""), ano_lectivo)
    if p2: t2_start, t2_end = p2
    p3 = parse_rango_fecha(cfg.get("fecha_t3", ""), ano_lectivo)
    if p3: t3_start, t3_end = p3

    if "1" in trimestre_str:
        return t1_start, t1_end
    elif "2" in trimestre_str:
        return t2_start, t2_end
    else:
        return t3_start, t3_end

# Variables globales para el desfase del reloj de Panamá
_PANAMA_TIME_OFFSET_SECONDS = 0.0
_TIME_SYNCHRONIZED = False

def iniciar_sincronizacion_hora_panama():
    """Inicia un hilo en segundo plano para sincronizar la hora de Panamá con internet."""
    import threading
    def sync_thread():
        global _PANAMA_TIME_OFFSET_SECONDS, _TIME_SYNCHRONIZED
        import urllib.request
        import json
        import email.utils
        
        # 1. Intentar con WorldTimeAPI (Panamá es GMT-5)
        try:
            req = urllib.request.Request(
                "http://worldtimeapi.org/api/timezone/America/Panama",
                headers={"User-Agent": "RegistroDoc/3.0"}
            )
            with urllib.request.urlopen(req, timeout=2.0) as response:
                data = json.loads(response.read().decode('utf-8'))
                datetime_str = data.get("datetime")
                if datetime_str:
                    dt_internet = datetime.datetime.fromisoformat(datetime_str.split(".")[0])
                    dt_sys = datetime.datetime.now()
                    _PANAMA_TIME_OFFSET_SECONDS = (dt_internet - dt_sys).total_seconds()
                    _TIME_SYNCHRONIZED = True
                    print(f"[Hora] Sincronizado vía WorldTimeAPI. Desfase: {_PANAMA_TIME_OFFSET_SECONDS} s")
                    return
        except Exception as e:
            print(f"[Hora] Falló WorldTimeAPI: {e}")

        # 2. Intentar con Google HTTP Date Header (Fallback)
        try:
            req = urllib.request.Request(
                "https://www.google.com",
                method="HEAD",
                headers={"User-Agent": "RegistroDoc/3.0"}
            )
            with urllib.request.urlopen(req, timeout=2.0) as response:
                date_header = response.headers.get("Date")
                if date_header:
                    parsed_time = email.utils.parsedate_to_datetime(date_header)
                    # Panamá está a GMT-5
                    panama_time = parsed_time + datetime.timedelta(hours=-5)
                    dt_internet = panama_time.replace(tzinfo=None)
                    dt_sys = datetime.datetime.now()
                    _PANAMA_TIME_OFFSET_SECONDS = (dt_internet - dt_sys).total_seconds()
                    _TIME_SYNCHRONIZED = True
                    print(f"[Hora] Sincronizado vía Google. Desfase: {_PANAMA_TIME_OFFSET_SECONDS} s")
                    return
        except Exception as e:
            print(f"[Hora] Falló Google Time Sync: {e}")

    threading.Thread(target=sync_thread, daemon=True).start()

def obtener_ahora_panama():
    """Retorna la fecha y hora actual ajustada a Panamá usando el desfase calculado."""
    ahora_sistema = datetime.datetime.now()
    if _TIME_SYNCHRONIZED:
        return ahora_sistema + datetime.timedelta(seconds=_PANAMA_TIME_OFFSET_SECONDS)
    return ahora_sistema

def obtener_hoy_panama():
    """Retorna la fecha actual ajustada a Panamá."""
    return obtener_ahora_panama().date()

def es_hora_sincronizada():
    """Retorna True si la hora fue sincronizada por internet."""
    return _TIME_SYNCHRONIZED
