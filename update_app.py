import re

with open("src/app.py", "r") as f:
    content = f.read()

# Make the sidebar correctly switch to a hamburger button
sidebar_pattern = r"""    def _sb_renderizar\(self\):
.*?
        toggle_txt = "«" if self._sidebar_exp else "»"
        ctk\.CTkButton\(self\._sb, text=toggle_txt, width=36, height=28,
                      fg_color=C\["hover"\], hover_color=C\["borde"\],
                      font=ctk\.CTkFont\(size=14, weight="bold"\),
                      text_color=C\["texto_sec"\],
                      corner_radius=6,
                      command=self\._toggle_sidebar_dash\)\.pack\(
            pady=\(4, 10\), padx=4\)"""

new_sidebar = """    def _sb_renderizar(self):
        for w in self._sb.winfo_children():
            w.destroy()

        ancho = 200 if self._sidebar_exp else 50

        # Botón hamburguesa (Hamburger toggle) at the top
        toggle_frame = ctk.CTkFrame(self._sb, fg_color="transparent")
        toggle_frame.pack(fill="x", pady=(10, 0))

        btn_toggle = ctk.CTkButton(
            toggle_frame, text="≡", width=36, height=36,
            fg_color="transparent", hover_color=C["hover"],
            font=ctk.CTkFont(size=24, weight="bold"),
            text_color=C["texto_sec"],
            corner_radius=6,
            command=self._toggle_sidebar_dash
        )
        if self._sidebar_exp:
             btn_toggle.pack(side="right", padx=10)
        else:
             btn_toggle.pack(pady=0)

        # Add Logo at top of Sidebar
        logo_path = os.path.join(os.path.dirname(__file__), "..", "img", "icono.jpg")
        if not os.path.exists(logo_path):
            logo_path = os.path.join(os.path.dirname(__file__), "..", "img", "logo.png")

        if PIL_OK and os.path.exists(logo_path):
            try:
                if self._sidebar_exp:
                    # Logo completo para expandido
                    size = (120, 120)
                    pil_logo = Image.open(logo_path).resize(size, Image.LANCZOS)
                    self._sb_logo_img = ctk.CTkImage(pil_logo, size=size)
                    ctk.CTkLabel(self._sb, image=self._sb_logo_img, text="").pack(pady=(12, 8))
                else:
                    # Ícono de libro/registro minimalista
                    size = (32, 32)
                    pil_logo = Image.open(logo_path).resize(size, Image.LANCZOS)
                    self._sb_logo_img = ctk.CTkImage(pil_logo, size=size)
                    ctk.CTkLabel(self._sb, image=self._sb_logo_img, text="").pack(pady=(12, 8))
            except Exception:
                pass
        else:
             if not self._sidebar_exp:
                 ctk.CTkLabel(self._sb, text="📘", fg_color="transparent",
                              font=ctk.CTkFont(size=24),
                              text_color=self._acento).pack(pady=(14, 8))

        ctk.CTkFrame(self._sb, fg_color=C["borde"], height=1).pack(
            fill="x", padx=8, pady=4)

        # Items del menú
        items = [
            ("🏠", "Inicio",        self._ir_inicio,    True),
            ("👤", "Estudiantes",   self._ir_estudiantes, False),
            ("📝", "Notas",         self._ir_notas,     False),
            ("📅", "Asistencia",    self._ir_asistencia, False),
            ("📋", "Reportes",      self._ir_reportes,  False),
        ]

        for icono, texto, cmd, activo in items:
            if self._sidebar_exp:
                label_txt = f"  {icono}   {texto}"
                anc = "w"
            else:
                label_txt = icono
                anc = "center"

            bg = C["activo"] if activo else "transparent"
            tc = self._acento if activo else C["texto_sec"]
            bw = 2 if activo else 0

            btn = ctk.CTkButton(
                self._sb, text=label_txt,
                fg_color=bg,
                hover_color=C["hover"],
                font=ctk.CTkFont("Segoe UI", 14 if self._sidebar_exp else 20),
                text_color=tc, anchor=anc,
                height=40, corner_radius=6,
                border_width=bw,
                border_color=self._acento,
                command=cmd)
            btn.pack(fill="x", padx=6, pady=2)

        # Espacio flexible
        ctk.CTkFrame(self._sb, fg_color="transparent").pack(
            fill="both", expand=True)

        if self._sidebar_exp:
            ctk.CTkLabel(self._sb, text="v.Prov.22:6",
                         font=ctk.CTkFont("Segoe UI", 10),
                         text_color=C["texto_dim"]).pack(pady=(0, 6))"""

content = re.sub(sidebar_pattern, new_sidebar, content, flags=re.DOTALL)


# Make the sidebar responsive in size
sidebar_size_pattern = r"nuevo_ancho = 200 if self._sidebar_exp else 45"
new_sidebar_size = 'nuevo_ancho = 200 if self._sidebar_exp else 50'
content = content.replace(sidebar_size_pattern, new_sidebar_size)


with open("src/app.py", "w") as f:
    f.write(content)
