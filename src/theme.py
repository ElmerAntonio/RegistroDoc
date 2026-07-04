import sys
import customtkinter as ctk

# ─── PALETA PRINCIPAL ──────────────────────────────────────────────
# Todos los colores definen un par: (modo_claro_alto_contraste, modo_oscuro)
C = {
    "fondo":        ("#F8FAFC", "#0A1628"),
    "sidebar":      ("#E2E8F0", "#0D1F35"),
    "header":       ("#E2E8F0", "#0D1F35"),
    "card":         ("#FFFFFF", "#0F2744"),
    "card_borde":   ("#0EA5E9", "#00DDEB"),
    "cian":         ("#0284C7", "#00DDEB"),
    "verde":        ("#16A34A", "#00FF88"),
    "rojo":         ("#DC2626", "#FF4444"),
    "amarillo":     ("#D97706", "#FFD700"),
    "purpura":      ("#7C3AED", "#A855F7"),
    "texto":        ("#0F172A", "#E2E8F0"),
    "texto_sec":    ("#475569", "#64748B"),
    "texto_dim":    ("#64748B", "#334155"),
    "input":        ("#FFFFFF", "#1E3A5F"),
    "hover":        ("#E2E8F0", "#1A3352"),
    "activo":       ("#CBD5E1", "#1A3352"),
    "borde":        ("#CBD5E1", "#1E3A5F"),
    "glow":         ("#E2E8F0", "#1E3A5F"),
    # ─── NUEVOS COLORES ÚNICOS ────────────────────────────
    "acento2":      ("#0284C7", "#38BDF8"),   # Azul cielo — botones secundarios
    "naranja":      ("#EA580C", "#FB923C"),   # Naranja cálido — alertas suaves
    "rosa":         ("#DB2777", "#F472B6"),   # Rosa — badges y acentos
    "esmeralda":    ("#059669", "#34D399"),   # Esmeralda — éxito alternativo
    "card_hover":   ("#F1F5F9", "#132E4A"),   # Card hover sutil
    "card_alt":     ("#F8FAFC", "#0C2137"),   # Card alternativo (zebra)
    "badge_bg":     ("#E2E8F0", "#1E3A5F"),   # Fondo de badges/pills
    "tooltip_bg":   ("#FFFFFF", "#1A3352"),   # Fondo de tooltips
    "divider":      ("#E2E8F0", "#1E3A5F"),   # Divisores
    "success":      ("#16A34A", "#10B981"),   # Verde success (botones guardar)
    "warning":      ("#D97706", "#F59E0B"),   # Amarillo warning
    "danger":       ("#DC2626", "#EF4444"),   # Rojo danger
    "info":         ("#2563EB", "#3B82F6"),   # Azul info
}

ACENTO = C["cian"]

# ─── FUENTES CON FALLBACKS MULTIPLATAFORMA ──────────────────────────────────
if sys.platform.startswith("win"):
    FONT_TITLE = "Outfit"
    FONT_BODY = "Segoe UI"
elif sys.platform.startswith("darwin"):
    FONT_TITLE = "Helvetica Neue"
    FONT_BODY = "Helvetica"
else:
    FONT_TITLE = "DejaVu Sans"
    FONT_BODY = "sans-serif"

# ─── ESCALAS DE FUENTE ──────────────────────────────────────────────
FONT_SIZES = {
    "xs": 10,
    "sm": 11,
    "md": 12,
    "lg": 14,
    "xl": 16,
    "2xl": 20,
    "3xl": 24,
}

# ─── COLOR CONTEXT HELPERS ──────────────────────────────────────────
def obtener_color_por_modo(color_tuple_or_str):
    if isinstance(color_tuple_or_str, tuple):
        mode = ctk.get_appearance_mode().lower()
        if mode == "dark":
            return color_tuple_or_str[1]
        else:
            return color_tuple_or_str[0]
    return color_tuple_or_str

def obtener_C_res():
    return {k: obtener_color_por_modo(v) for k, v in C.items()}



