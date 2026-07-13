import datetime
from utils.date_helpers import parse_rango_fecha, obtener_trimestre_actual, obtener_rango_fechas_trimestre

def test_parse_rango_fecha_spanish():
    # Rango en español
    res = parse_rango_fecha("06 DE MARZO AL 02 DE JUNIO", "2026")
    assert res is not None
    start, end = res
    assert start == datetime.date(2026, 3, 6)
    assert end == datetime.date(2026, 6, 2)

def test_parse_rango_fecha_hyphen():
    # Rango con guiones y slash
    res = parse_rango_fecha("06/03/2026 al 02/06/2026", "2026")
    assert res is not None
    start, end = res
    assert start == datetime.date(2026, 3, 6)
    assert end == datetime.date(2026, 6, 2)

def test_obtener_trimestre_actual():
    # Caso 1: 15 de abril (debería ser Trimestre 1)
    t1 = obtener_trimestre_actual(datetime.date(2026, 4, 15))
    assert t1 == "Trimestre 1"

    # Caso 2: 15 de julio (debería ser Trimestre 2)
    t2 = obtener_trimestre_actual(datetime.date(2026, 7, 15))
    assert t2 == "Trimestre 2"

    # Caso 3: 15 de octubre (debería ser Trimestre 3)
    t3 = obtener_trimestre_actual(datetime.date(2026, 10, 15))
    assert t3 == "Trimestre 3"

def test_obtener_rango_fechas_trimestre():
    start, end = obtener_rango_fechas_trimestre("Trimestre 2", "2026")
    # Validar que caiga en algún rango razonable (por defecto de Panamá o de config)
    assert start < end
    assert start.year == 2026
    assert end.year == 2026

def test_parse_fecha_segura():
    from utils.date_helpers import parse_fecha_segura
    # MM-DD
    assert parse_fecha_segura("05-12", ano_lectivo="2026") == datetime.date(2026, 5, 12)
    # DD-MM parsed securely
    assert parse_fecha_segura("25-10", ano_lectivo="2026") == datetime.date(2026, 10, 25)

def test_normalizar_fecha_a_mm_dd():
    from utils.date_helpers import normalizar_fecha_a_mm_dd
    # Standard format
    assert normalizar_fecha_a_mm_dd("05-12") == "05-12"
    # Ambiguous 12-05 - Trimestre 1 (March-June) maps to May 12 -> 05-12
    assert normalizar_fecha_a_mm_dd("12-05", trimestre=1) == "05-12"
    # Trimestre 3 (Sept-Dec) maps to Dec 5 -> 12-05
    assert normalizar_fecha_a_mm_dd("12-05", trimestre=3) == "12-05"
