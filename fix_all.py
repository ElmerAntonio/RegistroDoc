import re

# -------------
# 1. fix app.py
# -------------
with open("src/app.py", "r") as f:
    app_content = f.read()

if "import ctypes" not in app_content:
    app_content = re.sub(r'import sys', 'import sys\nimport ctypes', app_content)

if 'self.protocol("WM_DELETE_WINDOW", self.on_closing)' not in app_content:
    app_content = app_content.replace(
        'self.resizable(True, True) # Permite maximizar y achicar',
        'self.resizable(True, True) # Permite maximizar y achicar\n        self.protocol("WM_DELETE_WINDOW", self.on_closing)'
    )

if 'def on_closing(self):' not in app_content:
    on_closing_func = """
    def on_closing(self):
        \"\"\"Limpieza segura y salida completa del programa.\"\"\"
        self.destroy()
        sys.exit(0)
"""
    app_content = app_content.replace(
        'def destroy(self):\n        """Override destroy para marcar la aplicación como destruida y limpiar recursos."""',
        on_closing_func + '\n    def destroy(self):\n        """Override destroy para marcar la aplicación como destruida y limpiar recursos."""'
    )

if "ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID" not in app_content:
    app_content = app_content.replace(
        "def iniciar_programa_principal():\n    try:",
        'def iniciar_programa_principal():\n    if sys.platform == "win32":\n        try:\n            import ctypes\n            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("registrodoc.pro.v3")\n        except Exception:\n            pass\n    try:'
    )

old_icon_code = """        # Iconos de la ventana
        icon_path = os.path.join(BASE_DIR, "..", "img", "icono.png")
        if os.path.exists(icon_path):
            try:
                pil = Image.open(icon_path).resize((64, 64))
                self._icono_app = ImageTk.PhotoImage(pil)
                self.iconphoto(True, self._icono_app)
            except Exception:
                pass"""

new_icon_code = """        # Iconos de la ventana (resolución para Windows / Barra de tareas)
        icon_path = os.path.abspath(os.path.join(BASE_DIR, "..", "img", "icon.ico"))
        if os.path.exists(icon_path):
            try:
                self.iconbitmap(icon_path)
            except Exception:
                pass

        # Fallback multi-plataforma
        png_path = os.path.abspath(os.path.join(BASE_DIR, "..", "img", "icono.png"))
        if os.path.exists(png_path) and getattr(self, "iconphoto", None):
            try:
                pil = Image.open(png_path).resize((64, 64))
                self._icono_app = ImageTk.PhotoImage(pil)
                self.iconphoto(True, self._icono_app)
            except Exception:
                pass"""

app_content = app_content.replace(old_icon_code, new_icon_code)

with open("src/app.py", "w") as f:
    f.write(app_content)

# -------------
# 2. fix splash.py
# -------------
with open("src/splash.py", "r") as f:
    splash_content = f.read()

import_code = """import customtkinter as ctk
import tkinter as tk
from PIL import Image, ImageTk, ImageOps
import os
import time
from utils.frases_educacion import frase_del_dia"""

splash_content = re.sub(r'import customtkinter as ctk.*from utils.frases_educacion import frase_del_dia', import_code, splash_content, flags=re.DOTALL)

init_img = """        # Imagen de inicio
        img_path = os.path.join(os.path.dirname(__file__), "..", "img", "inicio.png")
        if os.path.exists(img_path):
            try:
                pil_img = Image.open(img_path)
                # Resize image keeping aspect ratio to fit inside 800x600
                pil_img = ImageOps.contain(pil_img, (800, 600), Image.Resampling.LANCZOS)

                # Create a background matching the app's background
                bg_img = Image.new("RGBA", (800, 600), "#0A1628")

                # Calculate position to center the image
                offset = ((800 - pil_img.width) // 2, (600 - pil_img.height) // 2)
                bg_img.paste(pil_img, offset)

                self.img = ImageTk.PhotoImage(bg_img)
                img_label = tk.Label(self, image=self.img, bg="#0A1628")
                # Fill entire window and make sure to show behind text
                img_label.place(x=0, y=0, relwidth=1, relheight=1)
            except Exception:
                self._fallback_sin_imagen()
        else:
            self._fallback_sin_imagen()"""

splash_content = re.sub(r'        # Imagen de inicio.*?        else:\n            self._fallback_sin_imagen\(\)', init_img, splash_content, flags=re.DOTALL)

if 'self.protocol("WM_DELETE_WINDOW"' not in splash_content:
    splash_content = splash_content.replace(
        'self.attributes("-topmost", True)  # Siempre encima',
        'self.attributes("-topmost", True)  # Siempre encima\n        self.protocol("WM_DELETE_WINDOW", self._on_closing)'
    )

if 'def _on_closing(self):' not in splash_content:
    on_closing = """
    def _on_closing(self):
        \"\"\"Limpieza al cerrar\"\"\"
        self._cerrar_splash()
"""
    splash_content = splash_content.replace(
        'def _cerrar_splash(self):',
        on_closing + '\n    def _cerrar_splash(self):'
    )

old_frase_label = """        # Frase educativa
        texto, ref = frase_del_dia()
        frase_label = ctk.CTkLabel(
            self,
            text=f"{texto}\\n\\n{ref}",
            font=("Segoe UI", 14, "italic"),
            text_color="#00DDEB",
            wraplength=700,
            justify="center"
        )
        frase_label.pack(side="bottom", pady=20)"""

new_frase_label = """        # Contenedor para Frase educativa con fondo oscuro semitransparente
        texto, ref = frase_del_dia()

        texto_frame = ctk.CTkFrame(self, fg_color="#0A1628", corner_radius=10, bg_color="transparent")
        texto_frame.pack(side="bottom", pady=20, padx=40)

        frase_label = ctk.CTkLabel(
            texto_frame,
            text=f"{texto}\\n\\n{ref}",
            font=("Segoe UI", 16, "italic"),
            text_color="#00DDEB",
            wraplength=700,
            justify="center"
        )
        frase_label.pack(padx=20, pady=10)"""

if "frase_label = ctk.CTkLabel(" in splash_content and "texto_frame = ctk.CTkFrame" not in splash_content:
    splash_content = splash_content.replace(old_frase_label, new_frase_label)

with open("src/splash.py", "w") as f:
    f.write(splash_content)


# -------------
# 3. fix grapp.py
# -------------
with open("src/grapp.py", "r") as f:
    grapp_content = f.read()

if "from matplotlib.figure import Figure" not in grapp_content:
    grapp_content = grapp_content.replace("import matplotlib.pyplot as plt", "import matplotlib.pyplot as plt\nfrom matplotlib.figure import Figure")

def repl_subplots(match):
    args = match.group(1)
    return f"fig = Figure({args})\n        ax = fig.add_subplot(111)"

grapp_content = re.sub(r'fig, ax = plt\.subplots\((.*?)\)', repl_subplots, grapp_content)

grapp_content = grapp_content.replace("for fig in self.fig_objs:\n            plt.close(fig)", "self.fig_objs.clear()")

grapp_content = grapp_content.replace("fig = Figure(figsize=(10, 4)\n        ax = fig.add_subplot(111), dpi=100)", "fig = Figure(figsize=(10, 4), dpi=100)\n        ax = fig.add_subplot(111)")
grapp_content = grapp_content.replace("fig = Figure(figsize=(5, 4)\n        ax = fig.add_subplot(111), dpi=100)", "fig = Figure(figsize=(5, 4), dpi=100)\n        ax = fig.add_subplot(111)")
grapp_content = grapp_content.replace("fig = Figure(figsize=(6, 4)\n        ax = fig.add_subplot(111), dpi=100)", "fig = Figure(figsize=(6, 4), dpi=100)\n        ax = fig.add_subplot(111)")
grapp_content = grapp_content.replace("fig = Figure(figsize=(10, 3.5)\n        ax = fig.add_subplot(111), dpi=100)", "fig = Figure(figsize=(10, 3.5), dpi=100)\n        ax = fig.add_subplot(111)")

grapp_content = grapp_content.replace(
"""        if not historial or len(historial) < 2:
            return""",
"""        if not historial or len(historial) < 2:
            ax.text(0.5, 0.5, "Sin Datos", color=self.C["texto_sec"], fontsize=14, ha='center', va='center')
            ax.axis('off')
            canvas = FigureCanvasTkAgg(fig, master=f_grafico)
            canvas.draw()
            canvas.get_tk_widget().pack(fill="both", expand=True, padx=10, pady=10)
            return"""
)

grapp_content = grapp_content.replace(
"""        total = aprobados + riesgo + reprobados
        if total == 0:
            return""",
"""        total = aprobados + riesgo + reprobados
        if total == 0:
            ax.text(0.5, 0.5, "Sin Datos", color=self.C["texto_sec"], fontsize=14, ha='center', va='center')
            ax.axis('off')
            canvas = FigureCanvasTkAgg(fig, master=f_grafico)
            canvas.draw()
            canvas.get_tk_widget().pack(fill="both", expand=True, padx=10, pady=10)
            return"""
)

grapp_content = grapp_content.replace(
"""        if not promedios_dict:
            return""",
"""        if not promedios_dict:
            ax.text(0.5, 0.5, "Sin Datos", color=self.C["texto_sec"], fontsize=14, ha='center', va='center')
            ax.axis('off')
            canvas = FigureCanvasTkAgg(fig, master=f_grafico)
            canvas.draw()
            canvas.get_tk_widget().pack(fill="both", expand=True, padx=10, pady=10)
            return"""
)

grapp_content = grapp_content.replace(
"""        if not notas:
            return""",
"""        if not notas:
            ax.text(0.5, 0.5, "Sin Datos", color=self.C["texto_sec"], fontsize=14, ha='center', va='center')
            ax.axis('off')
            canvas = FigureCanvasTkAgg(fig, master=f_grafico)
            canvas.draw()
            canvas.get_tk_widget().pack(fill="both", expand=True, padx=10, pady=10)
            return"""
)


with open("src/grapp.py", "w") as f:
    f.write(grapp_content)

# -------------
# 4. fix eapp.py
# -------------
with open("src/eapp.py", "r") as f:
    eapp_content = f.read()

eapp_content = eapp_content.replace('fg_color="red"', 'fg_color="transparent"')

with open("src/eapp.py", "w") as f:
    f.write(eapp_content)

# -------------
# 5. fix sapp.py
# -------------
with open("src/sapp.py", "r") as f:
    sapp_content = f.read()

sapp_content = re.sub(r'pady=\(50,\s*20\)', 'pady=(20, 10)', sapp_content)
sapp_content = re.sub(r'pady=\(50,\s*0\)', 'pady=(20, 0)', sapp_content)
sapp_content = re.sub(r'pady=40', 'pady=10', sapp_content)
sapp_content = re.sub(r'pady=\(100,\s*20\)', 'pady=(20, 10)', sapp_content)

with open("src/sapp.py", "w") as f:
    f.write(sapp_content)

# -------------
# 6. fix dashapp.py
# -------------
with open("src/dashapp.py", "r") as f:
    dashapp_content = f.read()

safe_toast = """
        def clear_toast():
            if self.winfo_exists() and hasattr(self, "_toast_lbl") and self._toast_lbl.winfo_exists():
                self._toast_lbl.configure(text="")
        self.after(3000, clear_toast)
"""

dashapp_content = re.sub(r'self\.after\(3000, lambda: self\._toast_lbl\.configure\(text=""\)\)', safe_toast.strip(), dashapp_content)

with open("src/dashapp.py", "w") as f:
    f.write(dashapp_content)
