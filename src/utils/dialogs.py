import customtkinter as ctk
import tkinter as tk

class CTkCustomMessagebox(ctk.CTkToplevel):
    def __init__(self, parent, title, message, icon_type="info", option_type="ok", **kwargs):
        super().__init__(parent, **kwargs)
        self.title(title)
        self.resizable(False, False)
        if parent:
            self.transient(parent)
        try:
            self.grab_set()
        except Exception:
            pass
        
        # Tema visual
        from theme import C
        self.configure(fg_color=C.get("card", "#1A2638"))
        
        # Icono de ventana nativo
        from config import establecer_icono_ventana
        establecer_icono_ventana(self)
        
        self.result = None
        
        # Contenedor principal
        main_frame = ctk.CTkFrame(self, fg_color="transparent")
        main_frame.pack(fill="both", expand=True, padx=20, pady=20)
        
        # Tipo de icono y colores
        icons = {
            "info": ("ℹ️", C.get("cian", "#00DDEB")),
            "warning": ("⚠️", "#F59E0B"),
            "error": ("❌", "#EF4444"),
            "question": ("❓", "#3B82F6"),
            "success": ("✅", "#10B981")
        }
        icon_str, icon_color = icons.get(icon_type, ("ℹ️", C.get("cian", "#00DDEB")))
        
        # Contenedor de mensaje
        msg_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        msg_frame.pack(fill="both", expand=True, pady=(0, 15))
        
        # Icono gigante a la izquierda
        lbl_icon = ctk.CTkLabel(msg_frame, text=icon_str, font=("Segoe UI", 36), text_color=icon_color)
        lbl_icon.pack(side="left", padx=(5, 15), anchor="n")
        
        # Texto del mensaje con ajuste de línea
        lbl_text = ctk.CTkLabel(
            msg_frame, 
            text=message, 
            font=("Segoe UI", 12), 
            text_color=C.get("texto", "#FFFFFF"),
            justify="left",
            wraplength=380,
            anchor="w"
        )
        lbl_text.pack(side="left", fill="both", expand=True)
        
        # Botones
        btn_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        btn_frame.pack(fill="x", side="bottom")
        
        if option_type == "ok":
            btn_ok = ctk.CTkButton(
                btn_frame,
                text="Aceptar",
                font=("Segoe UI", 12, "bold"),
                fg_color=C.get("activo", "#00DDEB"),
                hover_color=C.get("hover", "#2563EB"),
                text_color="#1E2D42" if icon_type not in ["error", "warning"] else "#FFFFFF",
                width=100,
                height=32,
                command=self._on_ok
            )
            btn_ok.pack(side="right", padx=5)
            btn_ok.focus_set()
        elif option_type == "yesno":
            btn_yes = ctk.CTkButton(
                btn_frame,
                text="Sí",
                font=("Segoe UI", 12, "bold"),
                fg_color=C.get("verde", "#10B981"),
                hover_color="#059669",
                text_color="#FFFFFF",
                width=90,
                height=32,
                command=self._on_yes
            )
            btn_yes.pack(side="right", padx=5)
            btn_yes.focus_set()
            
            btn_no = ctk.CTkButton(
                btn_frame,
                text="No",
                font=("Segoe UI", 12),
                fg_color="#475569",
                hover_color="#334155",
                text_color="#FFFFFF",
                width=90,
                height=32,
                command=self._on_no
            )
            btn_no.pack(side="right", padx=5)
            
        # Centrado sobre el padre o la pantalla
        self.update_idletasks()
        
        # Calcular altura dinámica en función del texto y saltos de línea para evitar truncamiento
        lineas = message.split("\n")
        cant_lineas = 0
        for l in lineas:
            # Con wraplength=380, caben aprox. 55 caracteres en Segoe UI 12 por línea de texto
            cant_lineas += max(1, len(l) // 55)
            
        w = 500
        # 130px base para iconos/botones, más el alto de cada línea (18px)
        h = max(180, 130 + cant_lineas * 18)
        
        if parent:
            px = parent.winfo_toplevel().winfo_x() + (parent.winfo_toplevel().winfo_width() // 2) - (w // 2)
            py = parent.winfo_toplevel().winfo_y() + (parent.winfo_toplevel().winfo_height() // 2) - (h // 2)
        else:
            px = (self.winfo_screenwidth() // 2) - (w // 2)
            py = (self.winfo_screenheight() // 2) - (h // 2)
            
        self.geometry(f"{w}x{h}+{px}+{py}")
        
        self.protocol("WM_DELETE_WINDOW", self._on_close)
        try:
            self.wait_window(self)
        except Exception:
            pass
        finally:
            try:
                self.grab_release()
            except Exception:
                pass
        
    def _on_ok(self):
        self.result = True
        try:
            self.grab_release()
        except Exception:
            pass
        self.destroy()
        
    def _on_yes(self):
        self.result = True
        try:
            self.grab_release()
        except Exception:
            pass
        self.destroy()
        
    def _on_no(self):
        self.result = False
        try:
            self.grab_release()
        except Exception:
            pass
        self.destroy()
        
    def _on_close(self):
        self.result = False
        try:
            self.grab_release()
        except Exception:
            pass
        self.destroy()

def obtener_parent_activo():
    try:
        root = tk._default_root
        if not root:
            root = tk.Tk.nametowidget(".")
        if root:
            topmost = root
            for child in root.winfo_children():
                if isinstance(child, tk.Toplevel) and child.winfo_viewable():
                    topmost = child
            return topmost
    except Exception:
        pass
    return None

def _ejecutar_dialogo_en_hilo_principal(funcion_dialogo, *args, **kwargs):
    """
    Ejecuta una función de diálogo de Tkinter en el hilo principal
    y espera por su resultado de forma segura usando un threading.Event.
    """
    import threading
    
    # Si ya estamos en el hilo principal, ejecutar directamente
    if threading.current_thread() is threading.main_thread():
        return funcion_dialogo(*args, **kwargs)
        
    # De lo contrario, despachar al hilo principal
    evento = threading.Event()
    resultado = [None]
    
    def tarea():
        try:
            resultado[0] = funcion_dialogo(*args, **kwargs)
        finally:
            evento.set()
            
    # Obtener el root widget para usar after()
    try:
        import tkinter as tk
        root = tk._default_root
        if not root:
            root = tk.Tk.nametowidget(".")
        if root:
            root.after(0, tarea)
        else:
            evento.set()
    except Exception:
        evento.set()
        
    evento.wait(timeout=10)
    return resultado[0]

def show_info(title, message, **kwargs):
    def impl():
        parent = obtener_parent_activo()
        dialog = CTkCustomMessagebox(parent, title, message, icon_type="info", option_type="ok")
        return dialog.result
    return _ejecutar_dialogo_en_hilo_principal(impl)

def show_success(title, message, **kwargs):
    def impl():
        parent = obtener_parent_activo()
        dialog = CTkCustomMessagebox(parent, title, message, icon_type="success", option_type="ok")
        return dialog.result
    return _ejecutar_dialogo_en_hilo_principal(impl)

def show_warning(title, message, **kwargs):
    def impl():
        parent = obtener_parent_activo()
        dialog = CTkCustomMessagebox(parent, title, message, icon_type="warning", option_type="ok")
        return dialog.result
    return _ejecutar_dialogo_en_hilo_principal(impl)

def show_error(title, message, **kwargs):
    def impl():
        parent = obtener_parent_activo()
        dialog = CTkCustomMessagebox(parent, title, message, icon_type="error", option_type="ok")
        return dialog.result
    return _ejecutar_dialogo_en_hilo_principal(impl)

def ask_yes_no(title, message, **kwargs):
    def impl():
        parent = obtener_parent_activo()
        dialog = CTkCustomMessagebox(parent, title, message, icon_type="question", option_type="yesno")
        return dialog.result
    return _ejecutar_dialogo_en_hilo_principal(impl)

def patch_messagebox():
    """Realiza monkeypatch sobre tkinter.messagebox para usar los dialogos customizados."""
    import tkinter.messagebox
    tkinter.messagebox.showinfo = show_info
    tkinter.messagebox.showwarning = show_warning
    tkinter.messagebox.showerror = show_error
    tkinter.messagebox.askyesno = ask_yes_no
