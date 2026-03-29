# dashapp.py
# Dashboard final - estilo Grok + layout exacto que pediste

import customtkinter as ctk
from datetime import datetime


class DashboardFrame(ctk.CTkScrollableFrame):
    def __init__(self, master, engine, app_principal=None, **kwargs):
        super().__init__(master, fg_color="#0F172A", **kwargs)
        
        self.engine = engine
        self.app_principal = app_principal

        self.grid_columnconfigure(0, weight=1)

        self._crear_header()
        self._crear_tarjetas_metricas()
        self._crear_estado_asistencia()
        self._crear_predicciones()
        self._crear_grados_y_accesos()

    # ====================== HEADER ======================
    def _crear_header(self):
        header = ctk.CTkFrame(self, fg_color="transparent")
        header.grid(row=0, column=0, padx=30, pady=(25, 15), sticky="ew")

        ctk.CTkLabel(header, text="Panel de Control Principal",
                     font=ctk.CTkFont(size=28, weight="bold"), text_color="#E2E8F0").pack(anchor="w")
        ctk.CTkLabel(header, text="RegistroDoc Pro v3.0 - MEDUCA Panamá",
                     font=ctk.CTkFont(size=15), text_color="#94A3B8").pack(anchor="w", pady=2)

        fecha = datetime.now().strftime("%A, %d de %B de %Y").capitalize()
        ctk.CTkLabel(header, text=fecha, font=ctk.CTkFont(size=13), text_color="#64748B").pack(anchor="w", pady=6)

    # ====================== TARJETAS ======================
    def _crear_tarjetas_metricas(self):
        stats = self.engine.get_dashboard_stats()
        frame = ctk.CTkFrame(self, fg_color="transparent")
        frame.grid(row=1, column=0, padx=30, pady=15, sticky="ew")
        frame.grid_columnconfigure((0,1,2,3), weight=1, uniform="card")

        self._crear_tarjeta(frame, 0, "👥 Total Alumnos", str(stats.get("total", 0)),
                            "Estudiantes matriculados", "#3B82F6")
        self._crear_tarjeta(frame, 1, "⚠️ En Riesgo", str(stats.get("riesgo", 0)),
                            "Requieren seguimiento especial", "#EF4444")
        self._crear_tarjeta(frame, 2, "🏆 Mejor Alumno General", stats.get("honor", "Sin datos"),
                            "Mayor promedio de todos los grados", "#FACC15")
        self._crear_tarjeta(frame, 3, "📊 Asistencia Global", stats.get("asistencia", "98%"),
                            "Promedio general", "#22C55E")

    def _crear_tarjeta(self, parent, col, titulo, valor, subtitulo, color):
        card = ctk.CTkFrame(parent, fg_color="#1E2D42", corner_radius=16, border_width=2, border_color=color)
        card.grid(row=0, column=col, padx=8, pady=8, sticky="nsew")

        ctk.CTkLabel(card, text=titulo, font=ctk.CTkFont(size=14, weight="bold"), text_color="#CBD5E1").pack(pady=(16,4), padx=20, anchor="w")
        ctk.CTkLabel(card, text=valor, font=ctk.CTkFont(size=26, weight="bold"), text_color=color).pack(pady=0, padx=20, anchor="w")
        ctk.CTkLabel(card, text=subtitulo, font=ctk.CTkFont(size=13), text_color="#94A3B8").pack(pady=(0,16), padx=20, anchor="w")

    # ====================== ESTADO ASISTENCIA + GRÁFICO DE BARRAS ======================
    def _crear_estado_asistencia(self):
        stats = self.engine.get_dashboard_stats()
        porc = float(stats.get("asistencia", "98%").replace("%", ""))

        frame = ctk.CTkFrame(self, fg_color="#1E2D42", corner_radius=16)
        frame.grid(row=2, column=0, padx=30, pady=15, sticky="ew")
        frame.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(frame, text="Estado de Asistencia General", font=ctk.CTkFont(size=16, weight="bold"), text_color="#E2E8F0").pack(pady=(20,10), padx=25, anchor="w")

        # Barra principal
        progress = ctk.CTkProgressBar(frame, height=22, corner_radius=999)
        progress.pack(fill="x", padx=25, pady=8)
        progress.set(porc / 100)

        ctk.CTkLabel(frame, text=f"✅ Presente: {porc:.1f}%    ❌ Ausente: {100-porc:.1f}%",
                     font=ctk.CTkFont(size=13), text_color="#94A3B8").pack(pady=(0,15), padx=25, anchor="w")

        # Gráfico de barras simple (ejemplo: asistencia por grado)
        ctk.CTkLabel(frame, text="Asistencia por Grado", font=ctk.CTkFont(size=14, weight="bold"), text_color="#E2E8F0").pack(pady=(10,5), padx=25, anchor="w")

        # Barras simuladas (puedes reemplazar con datos reales del engine)
        for grado, valor in [("7°", 96), ("8°", 94), ("9°", 99)]:
            bar_frame = ctk.CTkFrame(frame, fg_color="transparent")
            bar_frame.pack(fill="x", padx=25, pady=3)
            ctk.CTkLabel(bar_frame, text=grado, width=30, font=ctk.CTkFont(size=13)).pack(side="left")
            bar = ctk.CTkFrame(bar_frame, fg_color="#3B82F6", height=18, corner_radius=6)
            bar.pack(side="left", fill="x", expand=True, padx=8)
            bar_inner = ctk.CTkFrame(bar, fg_color="#60A5FA", width=int(valor*2.5), height=18, corner_radius=6)
            bar_inner.pack(side="left", fill="y")
            ctk.CTkLabel(bar_frame, text=f"{valor}%", font=ctk.CTkFont(size=12)).pack(side="right")

    # ====================== PREDICCIONES ======================
    def _crear_predicciones(self):
        frame = ctk.CTkFrame(self, fg_color="#1E2D42", corner_radius=16)
        frame.grid(row=3, column=0, padx=30, pady=15, sticky="ew")
        frame.grid_columnconfigure((0,1,2), weight=1)

        ctk.CTkLabel(frame, text="Predicciones y Tendencias Educativas", font=ctk.CTkFont(size=16, weight="bold"), text_color="#E2E8F0").grid(row=0, column=0, columnspan=3, pady=(20,10), padx=25, sticky="w")

        self._crear_card_prediccion(frame, 0, "📈 Tendencia Asistencia", "↑ +3.2%", "Mejorando vs mes anterior", "#22C55E")
        self._crear_card_prediccion(frame, 1, "⚠️ Riesgo Próximo Trimestre", "4 alumnos", "Posible reprobación", "#F59E0B")
        self._crear_card_prediccion(frame, 2, "🎯 Proyección Fin de Año", "92% Promedio", "Meta alcanzable", "#3B82F6")

    def _crear_card_prediccion(self, parent, col, titulo, valor, subtitulo, color):
        card = ctk.CTkFrame(parent, fg_color="#0F172A", corner_radius=12)
        card.grid(row=1, column=col, padx=8, pady=8, sticky="nsew")
        ctk.CTkLabel(card, text=titulo, font=ctk.CTkFont(size=13), text_color="#CBD5E1").pack(pady=(12,4), padx=16, anchor="w")
        ctk.CTkLabel(card, text=valor, font=ctk.CTkFont(size=22, weight="bold"), text_color=color).pack(pady=0, padx=16, anchor="w")
        ctk.CTkLabel(card, text=subtitulo, font=ctk.CTkFont(size=12), text_color="#94A3B8").pack(pady=(0,12), padx=16, anchor="w")

    # ====================== GRADOS ACTIVOS Y ACCESOS RÁPIDOS (lado a lado) ======================
    def _crear_grados_y_accesos(self):
        frame = ctk.CTkFrame(self, fg_color="transparent")
        frame.grid(row=4, column=0, padx=30, pady=(0,30), sticky="ew")
        frame.grid_columnconfigure((0,1), weight=1, uniform="section")

        # Cuadro Grados Activos
        g_frame = ctk.CTkFrame(frame, fg_color="#1E2D42", corner_radius=16)
        g_frame.grid(row=0, column=0, padx=(0,8), pady=0, sticky="nsew")
        ctk.CTkLabel(g_frame, text="Grados Activos", font=ctk.CTkFont(size=17, weight="bold"), text_color="#E2E8F0").pack(pady=(20,10), padx=25, anchor="w")
        grados_f = ctk.CTkFrame(g_frame, fg_color="transparent")
        grados_f.pack(pady=(0,20), padx=25, anchor="w")
        for i, grado in enumerate(self.engine.obtener_grados_activos()):
            lbl = ctk.CTkLabel(grados_f, text=grado, font=ctk.CTkFont(size=15, weight="bold"), text_color="#60A5FA", fg_color="#1E2937", corner_radius=8, padx=20, pady=8)
            lbl.grid(row=0, column=i, padx=6)

        # Cuadro Accesos Rápidos
        a_frame = ctk.CTkFrame(frame, fg_color="#1E2D42", corner_radius=16)
        a_frame.grid(row=0, column=1, padx=(8,0), pady=0, sticky="nsew")
        ctk.CTkLabel(a_frame, text="Accesos Rápidos", font=ctk.CTkFont(size=17, weight="bold"), text_color="#E2E8F0").pack(pady=(20,15), padx=25, anchor="w")

        accesos = ctk.CTkFrame(a_frame, fg_color="transparent")
        accesos.pack(pady=(0,25), padx=25, fill="x")
        accesos.grid_columnconfigure((0,1,2), weight=1)

        botones = [
            ("📝 Ir a Notas", "#3B82F6", lambda: self.app_principal.mostrar_notas() if self.app_principal else None),
            ("📅 Tomar Asistencia", "#22C55E", lambda: self.app_principal.mostrar_asistencia() if self.app_principal else None),
            ("⚙️ Configuración", "#8B5CF6", lambda: self.app_principal.mostrar_configuracion() if self.app_principal else None),
        ]
        for i, (txt, color, cmd) in enumerate(botones):
            btn = ctk.CTkButton(accesos, text=txt, height=58, font=ctk.CTkFont(size=15, weight="bold"),
                                fg_color=color, hover_color=self._oscurecer_color(color), corner_radius=12, command=cmd)
            btn.grid(row=0, column=i, padx=8, pady=5, sticky="ew")

    def _oscurecer_color(self, color: str) -> str:
        mapa = {"#3B82F6": "#2563EB", "#22C55E": "#16A34A", "#8B5CF6": "#7C3AED"}
        return mapa.get(color, "#1E40AF")


# ====================== PRUEBA ======================
if __name__ == "__main__":
    ctk.set_appearance_mode("dark")
    root = ctk.CTk()
    root.title("RegistroDoc Pro - Dashboard")
    root.geometry("1280x820")

    class MockEngine:
        def get_dashboard_stats(self): return {"total": 92, "riesgo": 2, "honor": "ANTOS FIDEL (4.9)", "asistencia": "98%"}
        def obtener_grados_activos(self): return ["7°", "8°", "9°"]

    class MockApp:
        def mostrar_notas(self): print("Notas")
        def mostrar_asistencia(self): print("Asistencia")
        def mostrar_configuracion(self): print("Configuración")

    DashboardFrame(root, MockEngine(), MockApp()).pack(fill="both", expand=True, padx=10, pady=10)
    root.mainloop()