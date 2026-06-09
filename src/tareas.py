"""
RegistroDoc Pro — Sistema de Tareas Programadas
Permite al docente crear tareas/evaluaciones con anticipación
y recibir notificaciones en el Dashboard.
100% offline — almacena en JSON local.
"""
import os
import json
import datetime
from config import BASE_DIR

TAREAS_FILE = os.path.join(BASE_DIR, "..", "tareas_pendientes.json")


def cargar_tareas():
    """Carga las tareas desde el archivo JSON."""
    if not os.path.exists(TAREAS_FILE):
        return []
    try:
        with open(TAREAS_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return []


def guardar_tareas(tareas):
    """Guarda las tareas en el archivo JSON."""
    try:
        with open(TAREAS_FILE, "w", encoding="utf-8") as f:
            json.dump(tareas, f, ensure_ascii=False, indent=2)
        return True
    except Exception:
        return False


def agregar_tarea(titulo, grado, materia, tipo, fecha_limite, descripcion=""):
    """
    Agrega una nueva tarea pendiente.
    tipo: 'Parcial', 'Examen', 'Apreciación', 'Asistencia', 'Otro'
    fecha_limite: 'DD-MM-YYYY'
    """
    tareas = cargar_tareas()
    nueva = {
        "id": len(tareas) + 1,
        "titulo": titulo,
        "grado": grado,
        "materia": materia,
        "tipo": tipo,
        "fecha_limite": fecha_limite,
        "descripcion": descripcion,
        "completada": False,
        "creada": datetime.datetime.now().strftime("%Y-%m-%d %H:%M")
    }
    tareas.append(nueva)
    guardar_tareas(tareas)
    return nueva


def marcar_completada(tarea_id):
    """Marca una tarea como completada."""
    tareas = cargar_tareas()
    for t in tareas:
        if t["id"] == tarea_id:
            t["completada"] = True
            break
    guardar_tareas(tareas)


def eliminar_tarea(tarea_id):
    """Elimina una tarea."""
    tareas = cargar_tareas()
    tareas = [t for t in tareas if t["id"] != tarea_id]
    guardar_tareas(tareas)


def obtener_pendientes():
    """Retorna tareas pendientes ordenadas por fecha límite."""
    tareas = cargar_tareas()
    pendientes = [t for t in tareas if not t.get("completada", False)]
    hoy = datetime.date.today()

    for t in pendientes:
        try:
            partes = t["fecha_limite"].split("-")
            fecha = datetime.date(int(partes[2]), int(partes[1]), int(partes[0]))
            dias = (fecha - hoy).days
            t["_dias_restantes"] = dias
            if dias < 0:
                t["_urgencia"] = "vencida"
            elif dias == 0:
                t["_urgencia"] = "hoy"
            elif dias <= 2:
                t["_urgencia"] = "urgente"
            else:
                t["_urgencia"] = "normal"
        except Exception:
            t["_dias_restantes"] = 999
            t["_urgencia"] = "normal"

    return sorted(pendientes, key=lambda x: x.get("_dias_restantes", 999))


def obtener_clase_actual(horario):
    """
    Dado el horario del docente, retorna la clase actual según la hora.
    Retorna (materia, hora_rango) o (None, None) si no hay clase.
    """
    ahora = datetime.datetime.now()
    dia_semana = ahora.weekday()  # 0=lunes, 4=viernes

    dias_map = {0: "lunes", 1: "martes", 2: "miercoles", 3: "jueves", 4: "viernes"}
    if dia_semana > 4:
        return None, None, "Fin de semana"

    dia_key = dias_map[dia_semana]
    hora_actual = ahora.hour * 60 + ahora.minute  # minutos desde medianoche

    for periodo in horario:
        rango = periodo.get("horas", "")
        if not rango or "-" not in rango:
            continue
        try:
            inicio_str, fin_str = rango.split("-")
            h_ini, m_ini = map(int, inicio_str.strip().split(":"))
            h_fin, m_fin = map(int, fin_str.strip().split(":"))
            inicio_min = h_ini * 60 + m_ini
            fin_min = h_fin * 60 + m_fin

            if inicio_min <= hora_actual <= fin_min:
                materia = periodo.get(dia_key, "")
                if materia and materia.strip():
                    return materia, rango, dia_key
        except Exception:
            continue

    return None, None, dia_key


def obtener_proxima_clase(horario):
    """Retorna la próxima clase del día."""
    ahora = datetime.datetime.now()
    dia_semana = ahora.weekday()

    dias_map = {0: "lunes", 1: "martes", 2: "miercoles", 3: "jueves", 4: "viernes"}
    if dia_semana > 4:
        return None, None

    dia_key = dias_map[dia_semana]
    hora_actual = ahora.hour * 60 + ahora.minute

    for periodo in horario:
        rango = periodo.get("horas", "")
        if not rango or "-" not in rango:
            continue
        try:
            inicio_str, _ = rango.split("-")
            h_ini, m_ini = map(int, inicio_str.strip().split(":"))
            inicio_min = h_ini * 60 + m_ini

            if inicio_min > hora_actual:
                materia = periodo.get(dia_key, "")
                if materia and materia.strip():
                    return materia, rango
        except Exception:
            continue

    return None, None
