import customtkinter as ctk
import tkinter as tk
from PIL import Image, ImageTk, ImageOps
import os
from utils.frases_educacion import frase_del_dia

class SplashScreen(ctk.CTkToplevel):
    def __init__(self, parent=None):
        super().__init__(parent)
        self._destroyed = False  # Bandera para prevenir operaciones después de destrucción
        
        self.title("")
        self.geometry("800x600")
        self.resizable(False, False)
        self.overrideredirect(True)  # Sin bordes de ventana
        self.attributes("-topmost", True)  # Siempre encima
        self.protocol("WM_DELETE_WINDOW", self._on_closing)

        # Centrar en pantalla
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        x = (screen_width - 800) // 2
        y = (screen_height - 600) // 2
        self.geometry(f"800x600+{x}+{y}")

        # Fondo negro
        self.configure(fg_color="#0A1628")

        # Imagen de inicio
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
            self._fallback_sin_imagen()

        # Contenedor para Frase educativa con fondo oscuro semitransparente
        texto, ref = frase_del_dia()

        texto_frame = ctk.CTkFrame(self, fg_color="#0A1628", corner_radius=10, bg_color="transparent")
        texto_frame.pack(side="bottom", pady=20, padx=40)

        frase_label = ctk.CTkLabel(
            texto_frame,
            text=f"{texto}\n\n{ref}",
            font=("Segoe UI", 16, "italic"),
            text_color="#00DDEB",
            wraplength=700,
            justify="center"
        )
        frase_label.pack(padx=20, pady=10)

        # Barra de progreso
        self.progress = ctk.CTkProgressBar(self, width=400, height=10)
        self.progress.pack(side="bottom", pady=(0, 40))
        self.progress.set(0)

        # Animar progreso
        self.after(100, self._animar_carga)

    def _fallback_sin_imagen(self):
        ctk.CTkLabel(
            self,
            text="RegistroDoc Pro",
            font=("Segoe UI", 48, "bold"),
            text_color="#00DDEB"
        ).pack(expand=True)

    def _animar_carga(self):
        # Verificar si la ventana aún existe y no fue destruida antes de continuar
        if not self.winfo_exists() or self._destroyed:
            return
            
        current = self.progress.get()
        if current < 1.0:
            self.progress.set(current + 0.1)
            self.after(200, self._animar_carga)
        else:
            self.after(500, self._cerrar_splash)


    def _on_closing(self):
        """Limpieza al cerrar"""
        self._cerrar_splash()

    def _cerrar_splash(self):
        """Método seguro para cerrar el splash screen."""
        if self.winfo_exists() and not self._destroyed:
            self.destroy()

    def destroy(self):
        """Override destroy para marcar como destruido."""
        self._destroyed = True
        super().destroy()

    def mostrar(self):
        self.deiconify()
        self.wait_window()  # Esperar a que se cierre