import calendar
import datetime
import customtkinter as ctk
from theme import C, FONT_TITLE, FONT_BODY

class CTkCalendarPopup(ctk.CTkToplevel):
    def __init__(self, master, target_entry, formato="MM-DD", **kwargs):
        super().__init__(master, **kwargs)
        self.target_entry = target_entry
        self.formato = formato
        
        # Intentar parsear la fecha actual del entry para iniciar el calendario en ese mes
        self.fecha_actual = datetime.date.today()
        entrada = target_entry.get().strip()
        if entrada:
            try:
                if formato == "MM-DD" and len(entrada) == 5:
                    mes, dia = map(int, entrada.split("-"))
                    self.fecha_actual = datetime.date(self.fecha_actual.year, mes, dia)
                elif formato == "DD-MM-YYYY" and len(entrada) == 10:
                    dia, mes, anio = map(int, entrada.split("-"))
                    self.fecha_actual = datetime.date(anio, mes, dia)
            except Exception:
                pass

        self.mes_seleccionado = self.fecha_actual.month
        self.anio_seleccionado = self.fecha_actual.year

        self.title("Seleccionar Fecha")
        self.geometry("320x340")
        self.resizable(False, False)
        
        # Evitar redimensionamiento y mantener arriba
        self.transient(master)
        self.grab_set()
        
        # Configurar icono y centrado
        from config import establecer_icono_ventana
        try:
            establecer_icono_ventana(self)
        except Exception:
            pass

        # Centrar la ventana respecto al master
        self.update_idletasks()
        w = self.winfo_width()
        h = self.winfo_height()
        x = master.winfo_toplevel().winfo_x() + (master.winfo_toplevel().winfo_width() // 2) - (w // 2)
        y = master.winfo_toplevel().winfo_y() + (master.winfo_toplevel().winfo_height() // 2) - (h // 2)
        self.geometry(f"+{x}+{y}")

        # Configurar colores
        self.configure(fg_color=C["card"][1] if ctk.get_appearance_mode().lower() == "dark" else C["card"][0])

        self.crear_interfaz()
        self.dibujar_dias()

        # Liberar grab al cerrar
        self.protocol("WM_DELETE_WINDOW", self.cerrar)

    def crear_interfaz(self):
        # ─── HEADER: MES Y AÑO CON NAVEGACIÓN ───
        header = ctk.CTkFrame(self, fg_color="transparent")
        header.pack(fill="x", padx=10, pady=10)

        self.btn_ant = ctk.CTkButton(
            header, text="◀", width=30, height=30, fg_color=C["badge_bg"],
            hover_color=C["hover"], text_color=C["texto"], font=(FONT_BODY, 12, "bold"),
            command=self.mes_anterior
        )
        self.btn_ant.pack(side="left")

        self.lbl_mes_anio = ctk.CTkLabel(
            header, text="", font=(FONT_TITLE, 14, "bold"), text_color=C["cian"]
        )
        self.lbl_mes_anio.pack(side="left", fill="x", expand=True)

        self.btn_sig = ctk.CTkButton(
            header, text="▶", width=30, height=30, fg_color=C["badge_bg"],
            hover_color=C["hover"], text_color=C["texto"], font=(FONT_BODY, 12, "bold"),
            command=self.mes_siguiente
        )
        self.btn_sig.pack(side="right")

        # ─── DÍAS DE LA SEMANA ───
        semana_frame = ctk.CTkFrame(self, fg_color="transparent")
        semana_frame.pack(fill="x", padx=10, pady=2)
        
        dias_semana = ["Lu", "Ma", "Mi", "Ju", "Vi", "Sá", "Do"]
        for ds in dias_semana:
            lbl = ctk.CTkLabel(
                semana_frame, text=ds, width=42, font=(FONT_BODY, 11, "bold"),
                text_color=C["texto_sec"]
            )
            lbl.pack(side="left")

        # ─── CONTENEDOR DE LA REJILLA DE DÍAS ───
        self.grid_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.grid_frame.pack(fill="both", expand=True, padx=10, pady=5)

    def dibujar_dias(self):
        # Limpiar rejilla anterior
        for w in self.grid_frame.winfo_children():
            w.destroy()

        meses_nombres = [
            "", "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
            "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"
        ]
        self.lbl_mes_anio.configure(text=f"{meses_nombres[self.mes_seleccionado]} {self.anio_seleccionado}")

        # Obtener calendario de semanas
        cal = calendar.monthcalendar(self.anio_seleccionado, self.mes_seleccionado)
        hoy = datetime.date.today()

        for r_idx, week in enumerate(cal):
            row_frame = ctk.CTkFrame(self.grid_frame, fg_color="transparent")
            row_frame.pack(fill="x", pady=2)
            for c_idx, day in enumerate(week):
                if day == 0:
                    # Celda vacía (días fuera de mes)
                    empty = ctk.CTkLabel(row_frame, text="", width=42, height=30)
                    empty.pack(side="left")
                else:
                    # Determinar colores especiales
                    es_hoy = (self.anio_seleccionado == hoy.year and 
                              self.mes_seleccionado == hoy.month and 
                              day == hoy.day)
                    
                    es_seleccionado = (self.anio_seleccionado == self.fecha_actual.year and 
                                      self.mes_seleccionado == self.fecha_actual.month and 
                                      day == self.fecha_actual.day)

                    if es_seleccionado:
                        fg = C["cian"]
                        txt_color = "#FFFFFF"
                    elif es_hoy:
                        fg = C["badge_bg"]
                        txt_color = C["cian"]
                    else:
                        fg = "transparent"
                        txt_color = C["texto"]

                    btn = ctk.CTkButton(
                        row_frame, text=str(day), width=42, height=30,
                        fg_color=fg, hover_color=C["hover"], text_color=txt_color,
                        font=(FONT_BODY, 11), corner_radius=6,
                        command=lambda d=day: self.seleccionar_dia(d)
                    )
                    btn.pack(side="left")

    def mes_anterior(self):
        self.mes_seleccionado -= 1
        if self.mes_seleccionado == 0:
            self.mes_seleccionado = 12
            self.anio_seleccionado -= 1
        self.dibujar_dias()

    def mes_siguiente(self):
        self.mes_seleccionado += 1
        if self.mes_seleccionado == 13:
            self.mes_seleccionado = 1
            self.anio_seleccionado += 1
        self.dibujar_dias()

    def seleccionar_dia(self, dia):
        if self.formato == "MM-DD":
            valor = f"{self.mes_seleccionado:02d}-{dia:02d}"
        else:
            valor = f"{dia:02d}-{self.mes_seleccionado:02d}-{self.anio_seleccionado}"

        # Guardar estado original y cambiar a normal si es necesario
        estado_original = self.target_entry.cget("state")
        if estado_original in ["readonly", "disabled"]:
            self.target_entry.configure(state="normal")

        # Actualizar entry
        self.target_entry.delete(0, "end")
        self.target_entry.insert(0, valor)

        # Restaurar estado original si aplica
        if estado_original in ["readonly", "disabled"]:
            self.target_entry.configure(state=estado_original)

        # Disparar eventos si existen (ej. validaciones al cambiar el texto)
        try:
            self.target_entry.event_generate("<KeyRelease>")
        except Exception:
            pass

        self.cerrar()

    def cerrar(self):
        try:
            self.grab_release()
        except Exception:
            pass
        self.destroy()


def crear_date_picker(parent, formato="MM-DD", val_defecto=None, on_key_release=None, **entry_kwargs):
    """
    Función helper que crea un frame contenedor con un CTkEntry y un CTkButton '📅' de calendario.
    Retorna la tupla (frame, entry_widget).
    """
    frame = ctk.CTkFrame(parent, fg_color="transparent")
    
    entry = ctk.CTkEntry(frame, **entry_kwargs)
    entry.pack(side="left", fill="x", expand=True)

    if val_defecto:
        entry.insert(0, val_defecto)

    # Botón de calendario
    btn_cal = ctk.CTkButton(
        frame, text="📅", width=34, height=30, fg_color=C["badge_bg"],
        hover_color=C["hover"], text_color=C["texto"], font=(FONT_BODY, 13),
        command=lambda: CTkCalendarPopup(frame, entry, formato=formato)
    )
    btn_cal.pack(side="left", padx=(5, 0))

    # ---- LÓGICA DE VALIDACIÓN DE ENTRADA TECLADO ----
    def validar_teclado(event):
        texto = entry.get()
        # Filtrar solo caracteres válidos (números y guión)
        filtrado = "".join(c for c in texto if c.isdigit() or c == "-")
        
        # Limitar longitud máxima y auto-formatear
        if formato == "MM-DD":
            # Máximo 5 caracteres: MM-DD
            if len(filtrado) > 5:
                filtrado = filtrado[:5]
            
            # Auto-insertar el guión al escribir 2 dígitos
            if len(filtrado) == 2 and not filtrado.endswith("-") and event.keysym != "BackSpace":
                filtrado += "-"
            
            # Si borró el guión y quedan 3 dígitos, reajustar
            if len(filtrado) == 3 and not "-" in filtrado:
                filtrado = filtrado[:2] + "-" + filtrado[2:]
        else:
            # Máximo 10 caracteres: DD-MM-YYYY
            if len(filtrado) > 10:
                filtrado = filtrado[:10]
            
            # Auto-insertar guiones en posiciones 2 y 5
            if len(filtrado) == 2 and not filtrado.endswith("-") and event.keysym != "BackSpace":
                filtrado += "-"
            if len(filtrado) == 5 and not filtrado.endswith("-") and event.keysym != "BackSpace":
                filtrado += "-"
                
        if texto != filtrado:
            entry.delete(0, "end")
            entry.insert(0, filtrado)

        # Si hay un callback personalizado para KeyRelease, ejecutarlo después de la validación
        if on_key_release:
            on_key_release(event)

    entry.bind("<KeyRelease>", validar_teclado)
    
    return frame, entry
