import os

# Directorio base del código fuente (src)
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# Archivo de configuración principal (perfil.json)
CONFIG_FILE = os.path.join(BASE_DIR, "..", "perfil.json")

def establecer_icono_ventana(window):
    """Establece el icono de la aplicación en la ventana especificada (Toplevel o Tk)."""
    import os
    import sys
    from PIL import Image, ImageTk
    
    icon_path = os.path.join(BASE_DIR, "..", "assets", "icon.ico")
    png_path = os.path.join(BASE_DIR, "..", "assets", "icono.png")
    
    # En Windows, usar iconbitmap directamente con retardos para asegurar que no se sobrescriba
    if sys.platform == "win32" and os.path.exists(icon_path):
        try:
            window.iconbitmap(icon_path)
            if hasattr(window, "after"):
                window.after(100, lambda: window.iconbitmap(icon_path))
                window.after(300, lambda: window.iconbitmap(icon_path))
            return
        except Exception:
            pass
            
    # Fallback para macOS/Linux o si falla iconbitmap
    if os.path.exists(png_path) and hasattr(window, "iconphoto"):
        try:
            pil = Image.open(png_path)
            tk_img = ImageTk.PhotoImage(pil)
            window.iconphoto(False, tk_img)
            window._icono_app = tk_img
            if hasattr(window, "after"):
                window.after(200, lambda: window.iconphoto(False, tk_img))
        except Exception:
            pass

