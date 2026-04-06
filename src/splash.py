import customtkinter as ctk
import tkinter as tk
from PIL import Image, ImageTk, ImageOps
import os


class SplashScreen(ctk.CTkToplevel):
    def __init__(self, parent=None, duration_ms=1800):
        super().__init__(parent)
        self._destroyed = False
        self._close_job = None

        self.title("")
        self.geometry("800x600")
        self.resizable(False, False)
        self.overrideredirect(True)
        self.attributes("-topmost", True)
        self.protocol("WM_DELETE_WINDOW", self._on_closing)

        # Center window
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        x = (screen_width - 800) // 2
        y = (screen_height - 600) // 2
        self.geometry(f"800x600+{x}+{y}")

        self.configure(fg_color="#0A1628")

        img_path = os.path.join(os.path.dirname(__file__), "..", "img", "inicio.png")
        if os.path.exists(img_path):
            try:
                pil_img = Image.open(img_path).convert("RGB")
                pil_img = ImageOps.fit(
                    pil_img,
                    (800, 600),
                    method=Image.Resampling.LANCZOS,
                    centering=(0.5, 0.5),
                )

                self.img = ImageTk.PhotoImage(pil_img)
                self._img_label = tk.Label(
                    self,
                    image=self.img,
                    bg="#0A1628",
                    bd=0,
                    highlightthickness=0,
                )
                self._img_label.place(x=0, y=0, relwidth=1, relheight=1)
            except Exception:
                self._fallback_sin_imagen()
        else:
            self._fallback_sin_imagen()

        # Auto close after fixed time
        self._close_job = self.after(max(300, int(duration_ms)), self._cerrar_splash)

    def _fallback_sin_imagen(self):
        ctk.CTkLabel(
            self,
            text="RegistroDoc Pro",
            font=("Segoe UI", 48, "bold"),
            text_color="#00DDEB",
        ).pack(expand=True)

    def _on_closing(self):
        self._cerrar_splash()

    def _cerrar_splash(self):
        if self.winfo_exists() and not self._destroyed:
            self.destroy()

    def destroy(self):
        self._destroyed = True
        if self._close_job is not None:
            try:
                self.after_cancel(self._close_job)
            except Exception:
                pass
            self._close_job = None
        super().destroy()

    def mostrar(self):
        if self._destroyed:
            return
        self.deiconify()
        try:
            self.wait_window()
        except tk.TclError:
            pass

