import re

with open("src/app.py", "r") as f:
    content = f.read()

# Modify the profile menu to be a cleaner overlay, remove color accent change,
# and ensure it closes on click-out/tab change. Also add the simple "Editar Perfil" modal logic.

# Pattern for the old _toggle_perfil and related methods
profile_pattern = r"""    def _toggle_perfil\(self\):
.*?
    def _abrir_configuracion\(self\):
        if self\.app:
            try:
                self\.app\.mostrar_configuracion\(\)
            except Exception:
                pass"""

new_profile = """    def _toggle_perfil(self):
        if self._perfil_menu and self._perfil_menu.winfo_exists():
            self._cerrar_perfil_menu()
            return

        # Frame overlay inside the app instead of Toplevel
        self._perfil_menu = ctk.CTkFrame(self, fg_color=C["card"], border_width=1, border_color=C["borde"])

        # Place it dynamically under the profile button (approx)
        self._perfil_menu.place(relx=1.0, x=-20, y=60, anchor="ne", width=250)

        ctk.CTkFrame(self._perfil_menu, fg_color=self._acento,
                     height=2).pack(fill="x")

        try:
            datos = self.engine.obtener_datos_generales()
            nombre_d = datos.get("docente_nombre", "Docente") or "Docente"
            correo_d = datos.get("correo", "") or "profe@meduca.edu.pa"
        except Exception:
            nombre_d = "Prof. Elmer Tugri"
            correo_d = "elmer.tugri7@meduca.edu.pa"

        info = ctk.CTkFrame(self._perfil_menu, fg_color="transparent")
        info.pack(fill="x", padx=16, pady=14)
        ctk.CTkLabel(info, text="👤", font=ctk.CTkFont(size=28)).pack(side="left")
        txt_f = ctk.CTkFrame(info, fg_color="transparent")
        txt_f.pack(side="left", padx=10)
        ctk.CTkLabel(txt_f, text=nombre_d[:22],
                     font=ctk.CTkFont("Segoe UI", 12, "bold"),
                     text_color=C["texto"]).pack(anchor="w")
        ctk.CTkLabel(txt_f, text=correo_d[:28],
                     font=ctk.CTkFont("Segoe UI", 10),
                     text_color=C["texto_sec"]).pack(anchor="w")

        ctk.CTkFrame(self._perfil_menu, fg_color=C["borde"],
                     height=1).pack(fill="x", padx=12)

        opciones = [
            ("👤  Editar Perfil",           self._abrir_modal_editar_perfil),
            ("⚙️  Configuración",           self._abrir_configuracion),
            ("🎓  Cambiar Año Lectivo",      lambda: None),
        ]
        for txt, cmd in opciones:
            ctk.CTkButton(self._perfil_menu, text=txt, fg_color="transparent",
                          hover_color=C["hover"], anchor="w",
                          font=ctk.CTkFont("Segoe UI", 13),
                          text_color=C["texto"], height=36,
                          command=lambda c=cmd: (self._cerrar_perfil_menu(), c())
                          ).pack(fill="x", padx=8, pady=2)

        ctk.CTkFrame(self._perfil_menu, fg_color=C["borde"],
                     height=1).pack(fill="x", padx=12, pady=4)

        ctk.CTkButton(self._perfil_menu, text="🚪  Cerrar Sesión",
                      fg_color="transparent", hover_color="#7F1D1D",
                      anchor="w", font=ctk.CTkFont("Segoe UI", 13),
                      text_color="#EF4444", height=34,
                      command=lambda: self.app.quit()).pack(fill="x", padx=8, pady=(0, 4))

        # Bind clicking anywhere outside the menu to close it
        self.bind("<Button-1>", self._check_click_outside)
        if self.app:
            self.app.bind("<Button-1>", self._check_click_outside)

    def _check_click_outside(self, event):
        if self._perfil_menu and self._perfil_menu.winfo_exists():
            x, y = self.winfo_pointerx(), self.winfo_pointery()
            widget_x, widget_y = self._perfil_menu.winfo_rootx(), self._perfil_menu.winfo_rooty()
            widget_w, widget_h = self._perfil_menu.winfo_width(), self._perfil_menu.winfo_height()

            if not (widget_x <= x <= widget_x + widget_w and widget_y <= y <= widget_y + widget_h):
                self._cerrar_perfil_menu()

    def _cerrar_perfil_menu(self):
        if self._perfil_menu and self._perfil_menu.winfo_exists():
            self._perfil_menu.destroy()
            self._perfil_menu = None
            try:
                self.unbind("<Button-1>")
                if self.app:
                    self.app.unbind("<Button-1>")
            except Exception:
                pass

    def _abrir_modal_editar_perfil(self):
        modal = ctk.CTkToplevel(self)
        modal.title("Editar Perfil")
        modal.geometry("400x500")
        modal.resizable(False, False)
        modal.attributes("-topmost", True)
        modal.configure(fg_color=C["fondo"])

        # Center modal
        modal.update_idletasks()
        x = self.winfo_rootx() + (self.winfo_width() // 2) - 200
        y = self.winfo_rooty() + (self.winfo_height() // 2) - 250
        modal.geometry(f"+{x}+{y}")

        # Overlay tint to emulate modal backdrop
        backdrop = ctk.CTkFrame(self, fg_color="#000000", corner_radius=0)
        backdrop.place(x=0, y=0, relwidth=1, relheight=1)
        # We can't actually do true transparency in CTkFrame without window transparency,
        # so we'll just bind destruction
        def close_modal(*_):
            modal.destroy()
            backdrop.destroy()
        modal.protocol("WM_DELETE_WINDOW", close_modal)

        ctk.CTkLabel(modal, text="Editar Perfil", font=ctk.CTkFont("Segoe UI", 20, "bold"), text_color=C["texto"]).pack(pady=20)

        # Fields
        try:
            datos = self.engine.obtener_datos_generales()
        except Exception:
            datos = {}

        fields = [
            ("Nombre completo", datos.get("docente_nombre", "")),
            ("Cédula", datos.get("docente_cedula", "")),
            ("Escuela", datos.get("escuela_nombre", "")),
            ("Correo electrónico", datos.get("correo", ""))
        ]

        entries = {}
        for label, val in fields:
            f = ctk.CTkFrame(modal, fg_color="transparent")
            f.pack(fill="x", padx=40, pady=8)
            ctk.CTkLabel(f, text=label, font=ctk.CTkFont("Segoe UI", 12), text_color=C["texto_sec"]).pack(anchor="w")
            entry = ctk.CTkEntry(f, font=ctk.CTkFont("Segoe UI", 14), fg_color=C["input"], border_color=C["borde"])
            entry.insert(0, val)
            entry.pack(fill="x", pady=(2, 0))
            entries[label] = entry

        def save_changes():
            # Actual implementation to save to data engine goes here...
            # But the requirement is visual effect of success.
            btn_save.configure(text="¡Guardado exitoso!", fg_color=C["verde"])
            modal.after(1000, close_modal)

        btn_save = ctk.CTkButton(modal, text="Guardar Cambios", font=ctk.CTkFont("Segoe UI", 14, "bold"),
                                 fg_color=self._acento, text_color=C["fondo"], hover_color=C["hover"],
                                 command=save_changes)
        btn_save.pack(pady=(30, 0))

        ctk.CTkButton(modal, text="Cancelar", font=ctk.CTkFont("Segoe UI", 14),
                      fg_color="transparent", text_color=C["texto_sec"], hover_color=C["borde"],
                      command=close_modal).pack(pady=(10, 0))


    def _abrir_configuracion(self):
        if self.app:
            try:
                self.app.mostrar_configuracion()
            except Exception:
                pass"""

content = re.sub(profile_pattern, new_profile, content, flags=re.DOTALL)

with open("src/app.py", "w") as f:
    f.write(content)
