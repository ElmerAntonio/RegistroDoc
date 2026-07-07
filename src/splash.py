import customtkinter as ctk
import tkinter as tk
from PIL import Image, ImageTk
import os
import sys
import math
import random
from config import ASSETS_DIR

CITAS_ADVENTISTAS = [
    ("Proverbios 22:6", "Instruye al niño en su camino, y aun cuando fuere viejo no se apartará de él."),
    ("Isaías 54:13", "Y todos tus hijos serán enseñados por Jehová; y se multiplicará la paz de tus hijos."),
    ("Elena G. de White", "La verdadera educación es el desarrollo armonioso de las facultades físicas, mentales y espirituales."),
    ("Proverbios 9:10", "El temor de Jehová es el principio de la sabiduría, y el conocimiento del Santísimo es la inteligencia."),
    ("Elena G. de White", "El objeto de la educación es restaurar la imagen de su Creador en el alma."),
    ("Salmo 119:105", "Lámpara es a mis pies tu palabra, y lumbrera a mi camino."),
    ("Proverbios 2:6", "Porque Jehová da la sabiduría, y de su boca viene el conocimiento y la inteligencia."),
    ("Elena G. de White", "La educación que dura por la eternidad es la que se obtiene al conocer a Dios.")
]

class CircularProgress(tk.Canvas):
    def __init__(self, parent, size=74, thickness=3, fg_color="#06B6D4", bg_color="#1E293B", bg="#091524", **kwargs):
        super().__init__(parent, width=size, height=size, bg=bg, bd=0, highlightthickness=0, **kwargs)
        self.size = size
        self.thickness = thickness
        self.fg_color = fg_color
        self.bg_color = bg_color
        self.porcentaje = 0.0
        self.angulo_rotacion = 0
        self.dibujar()
        self.animar()

    def animar(self):
        if not self.winfo_exists():
            return
        self.angulo_rotacion = (self.angulo_rotacion + 12) % 360
        self.dibujar()
        self.after(60, self.animar)

    def dibujar(self):
        self.delete("all")
        
        # Draw background ring
        pad = self.thickness + 2
        self.create_oval(pad, pad, self.size - pad, self.size - pad, outline=self.bg_color, width=self.thickness)
        
        # Segmented glow effect (Neon Chaser)
        num_ticks = 16
        centro = self.size / 2
        radio = (self.size - 2 * pad) / 2
        
        for i in range(num_ticks):
            angulo_grad = (i * (360 / num_ticks) + self.angulo_rotacion) % 360
            angulo_rad = math.radians(angulo_grad)
            
            x1 = centro + radio * math.cos(angulo_rad)
            y1 = centro + radio * math.sin(angulo_rad)
            
            # Length of each tick segment
            x2 = centro + (radio - 4) * math.cos(angulo_rad)
            y2 = centro + (radio - 4) * math.sin(angulo_rad)
            
            # Chaser fade logic
            dist = i
            if dist < 4:
                # Fade chaser colors
                alpha_vals = ["#22D3EE", "#06B6D4", "#0891B2", "#0E7490"]
                color = alpha_vals[dist]
                width_val = self.thickness + (1.5 if dist == 0 else 0)
            elif dist == num_ticks - 1:
                color = "#0891B2" # Glow tail
                width_val = self.thickness
            else:
                # Base loaded/unloaded color
                is_active = (i / num_ticks) * 100.0 < self.porcentaje
                color = "#0E7490" if is_active else self.bg_color
                width_val = self.thickness
                
            self.create_line(x1, y1, x2, y2, fill=color, width=width_val, capstyle="round")

    def actualizar_progreso(self, porcentaje):
        self.porcentaje = max(0.0, min(100.0, porcentaje))
        self.dibujar()

class SplashScreen(ctk.CTkToplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.title("Cargando RegistroDoc...")
        self.overrideredirect(True)
        
        # Compact and elegant dimensions
        width = 330
        height = 420
        
        # Settle Tkinter state and center window natively without scale factor multiplication
        self.update_idletasks()
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        
        x = (screen_width - width) // 2
        y = (screen_height - height) // 2
        self.geometry(f"{width}x{height}+{x}+{y}")
        
        # Transparent window key
        self.trans_color = "#000001"
        self.configure(fg_color=self.trans_color)
        
        if sys.platform == "win32":
            try:
                self.wm_attributes("-transparentcolor", self.trans_color)
            except Exception:
                pass

        # Configure AppUserModelID for taskbar grouping
        if sys.platform == "win32":
            try:
                import ctypes
                ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("registrodoc.pro.v3")
            except Exception:
                pass

        # Load taskbar and window icons
        icon_path = os.path.join(ASSETS_DIR, "icon_fixed.ico")
        png_path = os.path.join(ASSETS_DIR, "icono.png")
        if os.path.exists(icon_path):
            try:
                self.iconbitmap(icon_path)
            except Exception:
                pass
        if os.path.exists(png_path):
            try:
                pil_icon = Image.open(png_path).resize((64, 64))
                self._icono_app = ImageTk.PhotoImage(pil_icon)
                self.iconphoto(True, self._icono_app)
            except Exception:
                pass
                
        # Start hidden/transparent for fade-in
        self.attributes("-alpha", 0.0)
        self.attributes("-topmost", True)
        
        # Obsidian Floating Card Container
        self.card = ctk.CTkFrame(
            self,
            fg_color="#091524",
            border_width=2,
            border_color="#1E2E42",
            corner_radius=28
        )
        self.card.pack(fill="both", expand=True, padx=10, pady=10)

        # Header Badge
        badge_frame = ctk.CTkFrame(self.card, fg_color="transparent", border_width=1, border_color="#06B6D4", corner_radius=12)
        badge_frame.pack(pady=(22, 5))
        
        lbl_badge = ctk.CTkLabel(badge_frame, text="  SISTEMA ACTIVO  ", font=("Outfit", 8, "bold"), text_color="#06B6D4", height=18)
        lbl_badge.pack()
        
        # Centered logo
        self.logo_label = tk.Label(self.card, bg="#091524")
        self.logo_label.pack(pady=(5, 5))
        
        logo_path = os.path.join(ASSETS_DIR, "logo.png")
        if not os.path.exists(logo_path):
            logo_path = os.path.join(ASSETS_DIR, "icono.png")
            
        if os.path.exists(logo_path):
            try:
                pil_img = Image.open(logo_path).resize((130, 130), Image.Resampling.LANCZOS)
                self.logo_img = ImageTk.PhotoImage(pil_img)
                self.logo_label.configure(image=self.logo_img)
                self.logo_label.image = self.logo_img
            except Exception:
                self.logo_label.configure(text="📘", font=("Outfit", 54), fg="#06B6D4")
        else:
            self.logo_label.configure(text="📘", font=("Outfit", 54), fg="#06B6D4")

        # Compact segmented circular progress canvas
        self.progress = CircularProgress(
            self.card,
            size=74,
            thickness=3,
            fg_color="#06B6D4",
            bg_color="#1E293B",
            bg="#091524"
        )
        self.progress.pack(pady=5)
        
        # Static central percentage label placed strictly on top of progress canvas (no wobble/vibration)
        self.lbl_percentage = ctk.CTkLabel(
            self.progress,
            text="0%",
            font=("Outfit", 9, "bold"),
            text_color="#F8FAFC",
            fg_color="transparent"
        )
        self.lbl_percentage.place(relx=0.5, rely=0.5, anchor="center")

        # Status loading label
        self.lbl_status = ctk.CTkLabel(
            self.card,
            text="INICIALIZANDO...",
            font=("Outfit", 8, "bold"),
            text_color="#64748B"
        )
        self.lbl_status.pack(pady=2)

        # Adventist/Biblical Quote Label
        autor, cita = random.choice(CITAS_ADVENTISTAS)
        self.lbl_quote = ctk.CTkLabel(
            self.card,
            text=f'"{cita}"\n— {autor}',
            font=("Outfit", 9, "italic"),
            text_color="#94A3B8",
            wraplength=280,
            justify="center"
        )
        self.lbl_quote.pack(pady=(12, 18))
        
        # Setup drag-and-drop window movement bindings
        self.x_drag = 0
        self.y_drag = 0
        self._configurar_arrastre_recursivo(self.card)

        self.tasks = []
        self.current_task_idx = 0
        self.on_complete_callback = None
        self.app_instance = None
        
        self.protocol("WM_DELETE_WINDOW", lambda: None)

    def _configurar_arrastre_recursivo(self, widget):
        widget.bind("<Button-1>", self.iniciar_arrastre)
        widget.bind("<B1-Motion>", self.realizar_arrastre)
        for child in widget.winfo_children():
            self._configurar_arrastre_recursivo(child)

    def iniciar_arrastre(self, event):
        self.x_drag = event.x_root - self.winfo_x()
        self.y_drag = event.y_root - self.winfo_y()

    def realizar_arrastre(self, event):
        x = event.x_root - self.x_drag
        y = event.y_root - self.y_drag
        self.geometry(f"+{x}+{y}")

    def set_tasks(self, tasks, on_complete):
        self.tasks = tasks
        self.on_complete_callback = on_complete

    def iniciar(self):
        # Settle Tkinter state and center window natively right before revealing it
        width = 330
        height = 420
        self.update_idletasks()
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        
        x = (screen_width - width) // 2
        y = (screen_height - height) // 2
        self.geometry(f"{width}x{height}+{x}+{y}")

        self.deiconify()
        
        # Force the borderless (overrideredirect) window onto the Windows taskbar
        if sys.platform == "win32":
            try:
                import ctypes
                hwnd = ctypes.windll.user32.GetParent(self.winfo_id())
                if not hwnd:
                    hwnd = self.winfo_id()
                    
                GWL_EXSTYLE = -20
                WS_EX_APPWINDOW = 0x00040000
                WS_EX_TOOLWINDOW = 0x00000080
                
                style = ctypes.windll.user32.GetWindowLongW(hwnd, GWL_EXSTYLE)
                style = (style & ~WS_EX_TOOLWINDOW) | WS_EX_APPWINDOW
                ctypes.windll.user32.SetWindowLongW(hwnd, GWL_EXSTYLE, style)
                
                # Refresh window to register style change with Windows taskbar
                self.withdraw()
                self.deiconify()
            except Exception as e:
                print(f"Error forzando icono en barra de tareas: {e}")
                
        self.fade_in(0.93, 0.0)

    def fade_in(self, target_alpha=0.93, current_alpha=0.0):
        if current_alpha < target_alpha:
            current_alpha += 0.08
            self.attributes("-alpha", min(current_alpha, target_alpha))
            self.after(15, lambda: self.fade_in(target_alpha, current_alpha))
        else:
            self.attributes("-alpha", target_alpha)
            self.after(100, self._ejecutar_siguiente_tarea)

    def fade_out(self, current_alpha=0.93):
        if current_alpha > 0.0:
            current_alpha -= 0.08
            self.attributes("-alpha", max(current_alpha, 0.0))
            self.after(15, lambda: self.fade_out(current_alpha))
        else:
            self.attributes("-alpha", 0.0)
            if self.on_complete_callback:
                self.on_complete_callback(self.app_instance)
            self.destroy()

    def _ejecutar_siguiente_tarea(self):
        if not self.tasks:
            self.fade_out(0.93)
            return
            
        if self.current_task_idx >= len(self.tasks):
            self.progress.actualizar_progreso(100.0)
            self.lbl_percentage.configure(text="100%")
            self.lbl_status.configure(text="SISTEMA COMPLETADO")
            self.update()
            self.after(300, lambda: self.fade_out(0.93))
            return

        percentage = (self.current_task_idx / len(self.tasks)) * 100.0
        self.progress.actualizar_progreso(percentage)
        self.lbl_percentage.configure(text=f"{int(percentage)}%")
        
        name, fn = self.tasks[self.current_task_idx]
        self.lbl_status.configure(text=name.upper())
        
        # Flush Tkinter event queue
        self.update_idletasks()
        self.update()

        # Run a micro-spin transition loop (20ms) to ensure loader spins fluidly before blocking on the next task
        self._spin_loader(2, name, fn)

    def _spin_loader(self, steps_remaining, name, fn):
        if not self.winfo_exists():
            return
        if steps_remaining > 0:
            self.update_idletasks()
            self.update()
            self.after(10, lambda: self._spin_loader(steps_remaining - 1, name, fn))
        else:
            try:
                fn(self)
            except Exception as e:
                print(f"Error executing preload task '{name}': {e}")
                import traceback
                traceback.print_exc()

            self.update_idletasks()
            self.update()

            self.current_task_idx += 1
            # Run the next task after a short 10ms pause
            self.after(10, self._ejecutar_siguiente_tarea)
