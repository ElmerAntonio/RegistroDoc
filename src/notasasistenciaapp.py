import customtkinter as ctk
from theme import C

class NotasAsistenciaFrame(ctk.CTkFrame):
    def __init__(self, master, engine, **kwargs):
        super().__init__(master, fg_color="transparent", **kwargs)
        self.engine = engine
        
        # Tabview for unified grades & attendance
        self.tabview = ctk.CTkTabview(
            self,
            segmented_button_selected_color=C["cian"],
            segmented_button_selected_hover_color=C["hover"],
            segmented_button_unselected_hover_color=C["hover"],
            text_color=C["texto"]
        )
        self.tabview.pack(fill="both", expand=True, padx=10, pady=5)
        
        # Add tabs
        self.tab_notas = self.tabview.add("Notas")
        self.tab_asistencia = self.tabview.add("Asistencia")
        
        # Styling tab headers font
        self.tabview._segmented_button.configure(font=("Outfit", 12, "bold"))
        
        # Load child frames inside tabs
        from eapp import NotasFrame
        from fapp import AsistenciaFrame
        
        self.frame_notas = NotasFrame(self.tab_notas, self.engine)
        self.frame_notas.pack(fill="both", expand=True)
        
        self.frame_asistencia = AsistenciaFrame(self.tab_asistencia, self.engine)
        self.frame_asistencia.pack(fill="both", expand=True)
        
    def actualizar_vista(self):
        if hasattr(self.frame_notas, "actualizar_vista"):
            try:
                self.frame_notas.actualizar_vista()
            except Exception as e:
                print(f"[!] Error updating Notas view: {e}")
        if hasattr(self.frame_asistencia, "actualizar_vista"):
            try:
                self.frame_asistencia.actualizar_vista()
            except Exception as e:
                print(f"[!] Error updating Asistencia view: {e}")
