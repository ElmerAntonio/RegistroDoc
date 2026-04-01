import random

# Frases y versículos para educación
FRASES_EDUCACION = [
    # Versículos bíblicos
    ("Proverbios 22:6", "Instruye al niño en su camino, y aun cuando fuere viejo no se apartará de él."),
    ("Salmos 32:8", "Te haré entender, y te enseñaré el camino en que debes andar; sobre ti fijaré mis ojos."),
    ("Deuteronomio 6:6-7", "Y estas palabras que yo te mando hoy, estarán sobre tu corazón; y las repetirás a tus hijos...") ,
    ("Proverbios 1:7", "El principio de la sabiduría es el temor de Jehová; los insensatos desprecian la sabiduría y la enseñanza."),
    ("Colosenses 3:16", "La palabra de Cristo habite en abundancia en vosotros, enseñándoos y exhortándoos unos a otros en toda sabiduría."),
    # Elena G. de White
    ("Elena G. de White, La Educación, p. 13", "La verdadera educación significa más que la persecución de un cierto curso de estudio. Significa más que una preparación para la vida presente. Abarca todo el ser, y todo el período de la existencia posible al hombre."),
    ("Elena G. de White, La Educación, p. 17", "El propósito de la educación es restaurar en el hombre la imagen de su Hacedor, devolverle la perfección con que fue creado."),
    ("Elena G. de White, La Educación, p. 29", "El mayor deseo de Dios es que cada estudiante llegue a ser semejante a Cristo, el Modelo perfecto."),
    ("Elena G. de White, La Educación, p. 30", "La educación verdadera es el desarrollo armonioso de las facultades físicas, mentales y espirituales."),
]

def frase_del_dia():
    """
    Devuelve una frase/versículo diferente cada día, variando si se llama varias veces el mismo día.
    """
    import datetime
    hoy = datetime.date.today().toordinal()
    # Para variedad, mezclamos con la hora y un random
    hora = datetime.datetime.now().hour
    idx = (hoy + hora + random.randint(0, 100)) % len(FRASES_EDUCACION)
    ref, texto = FRASES_EDUCACION[idx]
    return texto, ref
