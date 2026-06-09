import sys
import customtkinter as ctk

# ─── PALETA PRINCIPAL ──────────────────────────────────────────────
C = {
    "fondo":        "#0A1628",
    "sidebar":      "#0D1F35",
    "header":       "#0D1F35",
    "card":         "#0F2744",
    "card_borde":   "#00DDEB",
    "cian":         "#00DDEB",
    "verde":        "#00FF88",
    "rojo":         "#FF4444",
    "amarillo":     "#FFD700",
    "purpura":      "#A855F7",
    "texto":        "#E2E8F0",
    "texto_sec":    "#64748B",
    "texto_dim":    "#334155",
    "input":        "#1E3A5F",
    "hover":        "#1A3352",
    "activo":       "#1A3352",
    "borde":        "#1E3A5F",
    "glow":         "#1E3A5F",
    # ─── NUEVOS COLORES ÚNICOS ────────────────────────────
    "acento2":      "#38BDF8",   # Azul cielo — botones secundarios
    "naranja":      "#FB923C",   # Naranja cálido — alertas suaves
    "rosa":         "#F472B6",   # Rosa — badges y acentos
    "esmeralda":    "#34D399",   # Esmeralda — éxito alternativo
    "card_hover":   "#132E4A",   # Card hover sutil
    "card_alt":     "#0C2137",   # Card alternativo (zebra)
    "badge_bg":     "#1E3A5F",   # Fondo de badges/pills
    "tooltip_bg":   "#1A3352",   # Fondo de tooltips
    "divider":      "#1E3A5F",   # Divisores
    "success":      "#10B981",   # Verde success (botones guardar)
    "warning":      "#F59E0B",   # Amarillo warning
    "danger":       "#EF4444",   # Rojo danger
    "info":         "#3B82F6",   # Azul info
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
    "hero": 32,
}
