"""
RegistroDoc Pro — Módulo de Impresión
======================================
Abre el Excel con su diseño original y lo envía a la impresora.
Compatible con Microsoft Excel y LibreOffice Calc.

© 2026 RegistroDoc Pro — Elmer Tugri — Panamá
"""

import os
import subprocess
import platform
from tkinter import messagebox


def _ruta_excel():
    from config import BASE_DIR
    from rdsecurity import cargar_config_segura

    base = BASE_DIR
    raiz = os.path.join(base, "..")

    modalidad = "premedia"
    try:
        cfg = cargar_config_segura({"modalidad": "premedia"})
        modalidad = str(cfg.get("modalidad", "premedia")).lower()
    except Exception:
        pass

    candidatos = []
    if modalidad == "primaria":
        candidatos.extend([
            os.path.join(raiz, "Registro_Primaria.xlsx"),
            os.path.join(base, "Registro_Primaria.xlsx"),
        ])
    else:
        candidatos.extend([
            os.path.join(raiz, "Registro_2026.xlsx"),
            os.path.join(base, "Registro_2026.xlsx"),
        ])

    # Fallback por si la modalidad del perfil no coincide.
    candidatos.extend([
        os.path.join(raiz, "Registro_2026.xlsx"),
        os.path.join(raiz, "Registro_Primaria.xlsx"),
        os.path.join(base, "Registro_2026.xlsx"),
        os.path.join(base, "Registro_Primaria.xlsx"),
    ])

    for ruta in candidatos:
        if os.path.exists(ruta):
            return ruta

    # Devuelve la ruta principal esperada aunque no exista para mantener mensajes consistentes.
    return candidatos[0] if candidatos else os.path.join(raiz, "Registro_2026.xlsx")


def _excel_disponible() -> bool:
    """Verifica si Microsoft Excel está instalado en Windows."""
    try:
        import winreg
        key = winreg.OpenKey(
            winreg.HKEY_LOCAL_MACHINE,
            r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\EXCEL.EXE"
        )
        winreg.CloseKey(key)
        return True
    except Exception:
        return False


def _libreoffice_disponible() -> bool:
    """Verifica si LibreOffice está instalado."""
    rutas = [
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
        "/usr/bin/libreoffice",
        "/usr/bin/soffice",
    ]
    return any(os.path.exists(r) for r in rutas)


def encontrar_hoja_impresion(wb, tipo, grado=None, materia=None) -> str:
    """
    Localiza dinámicamente el nombre de la hoja en el libro basándose en el tipo, grado y materia.
    Evita fallos por diferencias de formato (ej: Asistencia (7°) vs Asistencia 7° ).
    """
    tipo_upper = tipo.upper()
    grado_num = grado.replace("°", "").strip() if grado else ""
    materia_clean = materia.lower().replace(" ", "").replace(".", "").replace("á","a").replace("é","e").replace("í","i").replace("ó","o").replace("ú","u") if materia else ""

    for s in wb.sheetnames:
        s_upper = s.upper()
        if tipo == "Portada" and "PORTADA" in s_upper and "VISTOSA" not in s_upper:
            return s
        if tipo == "Caratula" and ("CARATULA" in s_upper or "CARÁTULA" in s_upper or "VISTOSA" in s_upper):
            return s
        if tipo == "Horarios" and "HORARIO" in s_upper:
            return s
        if tipo == "Resumen" and "RESUMEN" in s_upper and grado_num in s_upper:
            return s
        if tipo == "Asistencia" and "ASISTENCIA" in s_upper and grado_num in s_upper:
            return s
        if tipo == "Planilla" and "PLANILLA" in s_upper and grado_num in s_upper:
            if not materia_clean:
                return s
            sheet_clean = s.lower().replace(" ", "").replace(".", "").replace("á","a").replace("é","e").replace("í","i").replace("ó","o").replace("ú","u")
            if materia_clean in sheet_clean:
                return s
    return None


def abrir_para_imprimir(hoja: str = None) -> tuple[bool, str]:
    """
    Abre el Excel con Microsoft Excel o LibreOffice para imprimir.
    El usuario imprime desde la aplicación como siempre.
    Retorna (éxito: bool, mensaje: str)
    """
    ruta = _ruta_excel()
    if not os.path.exists(ruta):
        return False, "No se encontró Registro_2026.xlsx"

    sistema = platform.system()

    if sistema == "Windows":
        try:
            # Intentar abrir con Excel directamente
            os.startfile(ruta)
            return True, (
                "El archivo Excel se abrió correctamente.\n\n"
                "Para imprimir:\n"
                "1. Ve a la hoja que deseas imprimir\n"
                "2. Presiona Ctrl+P\n"
                "3. Ajusta márgenes si es necesario\n"
                "4. Haz clic en Imprimir"
            )
        except Exception as e:
            return False, f"No se pudo abrir el archivo: {e}"
    else:
        # Linux/Mac — LibreOffice
        if _libreoffice_disponible():
            try:
                subprocess.Popen(["libreoffice", "--calc", os.path.abspath(ruta)])
                return True, (
                    "Excel abierto en LibreOffice. "
                    "Usa Ctrl+P para imprimir."
                )
            except Exception as e:
                return False, f"Error al abrir LibreOffice: {e}"
        else:
            return False, (
                "No se encontró Microsoft Excel ni LibreOffice instalado."
            )


def imprimir_hoja_directo(nombre_hoja: str = "Portada") -> tuple[bool, str]:
    """
    En Windows con Excel instalado, imprime directamente sin abrir ventana.
    Usa el método de apertura directa que mantiene todo el formato.
    """
    ruta = _ruta_excel()
    if not os.path.exists(ruta):
        return False, "No se encontró el archivo Excel."

    if platform.system() != "Windows":
        return abrir_para_imprimir(nombre_hoja)

    try:
        # Import win32com here to prevent loading errors on non-Windows systems
        import win32com.client

        # Absolute path ensures we only use trusted, absolute file locations
        abs_ruta = ruta

        # Ensure we only interact with files that exist to
        # prevent path traversal issues
        if not os.path.isfile(abs_ruta):
            return False, "La ruta del archivo no es válida."

        # Conectar a Excel de manera segura
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False

        try:
            wb = excel.Workbooks.Open(abs_ruta, ReadOnly=True)
            try:
                hoja = wb.Sheets(nombre_hoja)
                hoja.PrintOut()
                mensaje = f"Hoja '{nombre_hoja}' enviada a la impresora."
                exito = True
            except Exception:
                mensaje = (
                    f"Error al imprimir la hoja '{nombre_hoja}'. "
                    "Verifica que exista."
                )
                exito = False
            finally:
                wb.Close(SaveChanges=False)
        except Exception as e_wb:
            mensaje = f"Error al abrir el archivo con Excel: {e_wb}"
            exito = False
        finally:
            excel.Quit()

        return exito, mensaje

    except Exception:
        # Fallback: abrir normalmente
        return abrir_para_imprimir(nombre_hoja)


class PanelImpresion:
    """
    Panel de impresión para agregar a la UI principal.
    Uso: nb.add(PanelImpresion(nb, app), text='🖨️  Imprimir')
    """
    def __new__(cls, parent, app):
        import tkinter as tk

        C = {
            "azul_osc": "#1C3557", "azul_med": "#2E6DA4",
            "azul_clar": "#D6E8FA", "blanco": "#FFFFFF",
            "gris_clar": "#F4F6FA", "amarillo": "#FFC000",
            "verde": "#2D6A4F", "rojo": "#C0392B",
            "texto": "#1A1A2E", "texto_med": "#4A5568",
        }

        frame = tk.Frame(parent, bg=C["blanco"])

        # Header
        hdr = tk.Frame(frame, bg=C["azul_osc"], height=52)
        hdr.pack(fill="x")
        hdr.pack_propagate(False)
        tk.Label(hdr, text="  🖨️  Impresión de Planillas",
                 font=("Segoe UI", 14, "bold"), fg=C["amarillo"],
                 bg=C["azul_osc"]).pack(side="left", pady=8)

        cuerpo = tk.Frame(frame, bg=C["blanco"], padx=30, pady=20)
        cuerpo.pack(fill="both", expand=True)

        tk.Label(cuerpo,
                 text=("El programa abre el Excel con su diseño y "
                       "formato originales.\n"
                       "Tú imprimes directamente desde ahí, "
                       "exactamente como MEDUCA lo requiere."),
                 font=("Segoe UI", 10), fg=C["texto_med"], bg=C["blanco"],
                 justify="center").pack(pady=(0, 20))

        # Hojas disponibles (Cargadas dinámicamente)
        import openpyxl
        hojas_comunes = []
        try:
            wb = openpyxl.load_workbook(app.engine.ruta, read_only=True)
            sheet_names = wb.sheetnames
            wb.close()
        except Exception:
            sheet_names = []

        for sheet in sheet_names:
            sheet_upper = sheet.upper()
            if "PORTADA" in sheet_upper and "VISTOSA" not in sheet_upper:
                hojas_comunes.append((sheet, "📋 Portada / Carátula oficial"))
            elif "CARATULA" in sheet_upper or "CARÁTULA" in sheet_upper or "VISTOSA" in sheet_upper:
                hojas_comunes.append((sheet, "📄 Carátula del registro"))
            elif "ASISTENCIA" in sheet_upper:
                grade = sheet.replace("Asistencia", "").replace("(", "").replace(")", "").strip()
                hojas_comunes.append((sheet, f"📅 Asistencia — Grado {grade}"))
            elif "PROM" in sheet_upper:
                subj = sheet.replace("PROM", "").replace("(", "").replace(")", "").strip()
                hojas_comunes.append((sheet, f"📝 PROM — {subj}"))
            elif "PLANILLA" in sheet_upper:
                subj = sheet.replace("Planilla", "").replace("(", "").replace(")", "").strip()
                hojas_comunes.append((sheet, f"📊 Planilla — {subj}"))
            elif "HORARIO" in sheet_upper:
                hojas_comunes.append((sheet, "🕐 Horario de clases"))
            elif "RESUMEN" in sheet_upper:
                grade = sheet.replace("RESUMEN", "").replace("(", "").replace(")", "").strip()
                hojas_comunes.append((sheet, f"📈 RESUMEN — {grade if grade else 'General'}"))

        if not hojas_comunes:
            hojas_comunes = [
                ("Portada",           "📋 Portada / Carátula oficial"),
                ("Caratula",          "📄 Carátula del registro"),
                ("Horarios",          "🕐 Horario de clases"),
            ]

        sel_f = tk.LabelFrame(cuerpo, text="  Selecciona qué imprimir  ",
                              font=("Segoe UI", 10, "bold"), bg=C["blanco"],
                              fg=C["azul_osc"], padx=20, pady=15)
        sel_f.pack(fill="x", pady=10)

        # Set default value safely to the first available sheet or "Portada"
        default_val = hojas_comunes[0][0] if hojas_comunes else "Portada"
        var_hoja = tk.StringVar(value=default_val)

        for i, (hoja_id, hoja_lbl) in enumerate(hojas_comunes):
            col = i % 2
            fila = i // 2
            tk.Radiobutton(
                sel_f, text=hoja_lbl, variable=var_hoja,
                value=hoja_id, font=("Segoe UI", 10),
                bg=C["blanco"], fg=C["texto"],
                selectcolor=C["azul_clar"], cursor="hand2"
            ).grid(row=fila, column=col, sticky="w",
                   padx=10, pady=3)

        btn_f = tk.Frame(cuerpo, bg=C["blanco"])
        btn_f.pack(pady=20)

        def abrir_excel():
            ok, msg = abrir_para_imprimir()
            if ok:
                messagebox.showinfo("✓ Excel abierto", msg)
            else:
                messagebox.showerror("Error", msg)

        def imprimir_sel():
            hoja = var_hoja.get()
            ok, msg = imprimir_hoja_directo(hoja)
            if ok:
                messagebox.showinfo("✓ Imprimiendo", msg)
            else:
                messagebox.showerror("Error", msg)

        tk.Button(btn_f, text="📂  Abrir Excel completo",
                  command=abrir_excel,
                  bg=C["azul_med"], fg=C["blanco"],
                  font=("Segoe UI", 11, "bold"), relief="flat",
                  cursor="hand2", padx=20, pady=10,
                  width=22).pack(side="left", padx=8)

        tk.Button(btn_f, text="🖨️  Imprimir hoja seleccionada",
                  command=imprimir_sel,
                  bg=C["verde"], fg=C["blanco"],
                  font=("Segoe UI", 11, "bold"), relief="flat",
                  cursor="hand2", padx=20, pady=10,
                  width=26).pack(side="left", padx=8)

        nota = tk.Label(cuerpo,
                        text="💡 El formato, fórmulas y diseño "
                             "del Excel nunca se modifican.\n"
                             "    Lo que ves en Excel es exactamente "
                             "lo que se imprime.",
                        font=("Segoe UI", 9), fg=C["texto_med"],
                        bg=C["blanco"], justify="left")
        nota.pack(pady=10)

        return frame
