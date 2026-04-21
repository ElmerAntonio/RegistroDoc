import os
import tempfile
from docx.shared import Inches

def save_figure_as_image(fig, name_prefix="grafico"): 
    """Guarda una figura de matplotlib como imagen temporal y retorna la ruta."""
    with tempfile.NamedTemporaryFile(delete=False, suffix=".png", prefix=name_prefix) as tmp:
        fig.savefig(tmp.name, bbox_inches='tight')
        return tmp.name

def add_image_to_doc(doc, image_path, ancho=5):
    doc.add_picture(image_path, width=Inches(ancho))
    doc.add_paragraph("")
    return doc
