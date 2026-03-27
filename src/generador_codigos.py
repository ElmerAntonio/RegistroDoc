"""
RegistroDoc Pro — Generador de Códigos de Activación
=====================================================
HERRAMIENTA PRIVADA DEL VENDEDOR — NO DISTRIBUIR JAMÁS

Seguridad:
  • SHA3-256 + PBKDF2-HMAC-SHA256 para generación de códigos
  • AES-256-GCM para el registro de ventas cifrado
  • Log de auditoría de todas las generaciones

© 2026 RegistroDoc Pro — Todos los derechos reservados
"""

import tkinter as tk
from tkinter import ttk, messagebox
import hashlib, re, json, os, datetime

from rdsecurity import (
    generar_codigo_licencia,
    guardar_cifrado, cargar_cifrado,
    registrar_auditoria,
    cifrar, descifrar,
    _hw_fingerprint,
    PBKDF2_ITERS,
)
from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC
from cryptography.hazmat.primitives import hashes
from cryptography.hazmat.backends import default_backend

VENTAS_FILE = "ventas_privado.enc"

# ── Clave de cifrado del registro de ventas (solo con HW del vendedor) ──
def _clave_ventas():
    hw   = _hw_fingerprint()
    salt = hashlib.sha3_256(hw + b"VENTAS_RD_2026").digest()
    kdf  = PBKDF2HMAC(algorithm=hashes.SHA256(), length=32, salt=salt,
                       iterations=PBKDF2_ITERS, backend=default_backend())
    return kdf.derive(hw)

def guardar_ventas(ventas: list) -> None:
    datos = {"ventas": ventas}
    blob  = cifrar(json.dumps(datos, ensure_ascii=False).encode(), _clave_ventas())
    with open(VENTAS_FILE, "wb") as f:
        f.write(blob)

def cargar_ventas() -> list:
    if not os.path.exists(VENTAS_FILE):
        return []
    try:
        with open(VENTAS_FILE, "rb") as f:
            blob = f.read()
        datos = json.loads(descifrar(blob, _clave_ventas()).decode())
        return datos.get("ventas", [])
    except Exception:
        return []

def registrar_venta(cedula, nombre, codigo, precio="", telefono="", nota=""):
    ventas = cargar_ventas()
    ventas.append({
        "fecha":    datetime.datetime.now().strftime("%d/%m/%Y %H:%M"),
        "cedula":   cedula,
        "nombre":   nombre,
        "codigo":   codigo,
        "precio":   precio,
        "telefono": telefono,
        "nota":     nota,
    })
    guardar_ventas(ventas)
    registrar_auditoria("VENTA","CODIGO_GENERADO",
                        f"Cédula {cedula[:4]}*** Nombre: {nombre[:20]}")

# ──────────────────────────────────────────────
C = {
    "azul_osc":"#1C3557","azul_med":"#2E6DA4","azul_clar":"#D6E8FA",
    "blanco":"#FFFFFF","gris_clar":"#F4F6FA","amarillo":"#FFC000",
    "verde":"#2D6A4F","rojo":"#C0392B","texto":"#1A1A2E","texto_med":"#4A5568",
    "naranja":"#E67E22",
}

def mk_btn(p,txt,cmd=None,color=None,w=18,**kw):
    return tk.Button(p,text=txt,command=cmd,bg=color or C["azul_med"],fg=C["blanco"],
                     font=("Segoe UI",10,"bold"),relief="flat",cursor="hand2",
                     padx=10,pady=7,width=w,**kw)

def mk_entry(p,var=None,w=32,**kw):
    return tk.Entry(p,textvariable=var,font=("Segoe UI",11),
                    bg=C["gris_clar"],relief="flat",bd=4,width=w,**kw)

# ──────────────────────────────────────────────

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("RegistroDoc Pro — Generador de Códigos  [PRIVADO DEL VENDEDOR]")
        self.geometry("860x720")
        self.configure(bg=C["gris_clar"])
        self.resizable(True,True)
        self._build()

    def _build(self):
        hdr=tk.Frame(self,bg=C["azul_osc"],height=65)
        hdr.pack(fill="x"); hdr.pack_propagate(False)
        tk.Label(hdr,text="  🔑  Generador de Códigos de Activación",
                 font=("Segoe UI",16,"bold"),fg=C["amarillo"],bg=C["azul_osc"]).pack(side="left",pady=10)
        tk.Label(hdr,text="⚠ HERRAMIENTA PRIVADA — NUNCA DISTRIBUIR",
                 font=("Segoe UI",9,"bold"),fg=C["rojo"],bg=C["azul_osc"]).pack(side="right",padx=20)

        nb=ttk.Notebook(self)
        nb.pack(fill="both",expand=True,padx=10,pady=10)
        nb.add(self._tab_generar(nb), text="🔑  Generar código")
        nb.add(self._tab_ventas(nb),  text="📋  Registro de ventas")
        nb.add(self._tab_verificar(nb), text="🔍  Verificar / Recuperar")

    def _tab_generar(self,parent):
        f=tk.Frame(parent,bg=C["blanco"],padx=30,pady=20)
        tk.Label(f,text="Genera el código único de activación para cada docente que compre el programa.",
                 font=("Segoe UI",10),fg=C["texto_med"],bg=C["blanco"]).pack(pady=(0,18))

        campos=[("N° Cédula del docente *:","var_ced"),
                ("Nombre completo del docente *:","var_nom"),
                ("Precio cobrado (B/.):", "var_precio"),
                ("Teléfono / WhatsApp:","var_tel"),
                ("Notas adicionales:","var_nota")]
        for etq,attr in campos:
            tf=tk.Frame(f,bg=C["blanco"]); tf.pack(fill="x",pady=4)
            tk.Label(tf,text=etq,width=34,anchor="w",font=("Segoe UI",10),
                     fg=C["texto"],bg=C["blanco"]).pack(side="left")
            v=tk.StringVar(); setattr(self,attr,v)
            mk_entry(tf,v,w=30).pack(side="left")

        tk.Frame(f,bg=C["gris_clar"],height=1).pack(fill="x",pady=14)
        mk_btn(f,"🔑  Generar código de activación",self._generar,color=C["verde"],w=34).pack(pady=8)

        res=tk.LabelFrame(f,text="  Código generado  ",
                           font=("Segoe UI",10,"bold"),bg=C["blanco"],
                           fg=C["azul_osc"],padx=20,pady=15)
        res.pack(fill="x",pady=10)

        self.lbl_codigo=tk.Label(res,text="—",font=("Courier New",20,"bold"),
                                  fg=C["azul_osc"],bg=C["blanco"])
        self.lbl_codigo.pack(pady=6)

        btns=tk.Frame(res,bg=C["blanco"]); btns.pack()
        mk_btn(btns,"📋  Copiar código",self._copiar,w=18).pack(side="left",padx=5)
        mk_btn(btns,"✉️  Preparar mensaje",self._msg,color=C["azul_osc"],w=22).pack(side="left",padx=5)

        self.lbl_status=tk.Label(f,text="",font=("Segoe UI",9),fg=C["verde"],bg=C["blanco"])
        self.lbl_status.pack()
        return f

    def _generar(self):
        ced   =self.var_ced.get().strip()
        nombre=self.var_nom.get().strip()
        if not ced:   messagebox.showwarning("Falta cédula","Ingresa el N° de cédula."); return
        if not nombre:messagebox.showwarning("Falta nombre","Ingresa el nombre del docente."); return

        codigo=generar_codigo_licencia(ced)
        self.lbl_codigo.config(text=codigo,fg=C["verde"])
        registrar_venta(ced,nombre,codigo,
                        precio   =self.var_precio.get().strip(),
                        telefono =self.var_tel.get().strip(),
                        nota     =self.var_nota.get().strip())
        self.lbl_status.config(
            text=f"✓ Guardado en registro cifrado  |  {datetime.datetime.now().strftime('%d/%m/%Y %H:%M')}")
        self._codigo_actual=codigo; self._nombre_actual=nombre; self._ced_actual=ced
        self._recargar_ventas()

    def _copiar(self):
        if not hasattr(self,"_codigo_actual") or not self._codigo_actual:
            messagebox.showwarning("Sin código","Genera un código primero."); return
        self.clipboard_clear(); self.clipboard_append(self._codigo_actual)
        messagebox.showinfo("✓ Copiado",f"Código copiado:\n{self._codigo_actual}")

    def _msg(self):
        if not hasattr(self,"_codigo_actual") or not self._codigo_actual:
            messagebox.showwarning("Sin código","Genera un código primero."); return
        nombre=getattr(self,"_nombre_actual","Docente")
        ced   =getattr(self,"_ced_actual","")
        msg=(f"Estimado(a) Prof. {nombre},\n\n"
             f"Su código de activación para RegistroDoc Pro es:\n\n"
             f"  {self._codigo_actual}\n\n"
             f"Instrucciones de activación:\n"
             f"1. Abra la aplicación RegistroDoc Pro\n"
             f"2. En la pantalla de activación, ingrese:\n"
             f"   • Su N° de cédula: {ced}\n"
             f"   • El código de activación de arriba\n"
             f"3. Haga clic en 'Activar programa'\n\n"
             f"⚠ Importante:\n"
             f"  • El programa se activa en UN SOLO dispositivo\n"
             f"  • Si cambia de computadora, contacte al proveedor\n"
             f"  • No comparta su código con otras personas\n\n"
             f"Para soporte técnico contacte al proveedor.\n\n"
             f"¡Gracias por su compra!")
        vent=tk.Toplevel(self)
        vent.title("Mensaje para el docente")
        vent.geometry("540,400".replace(",","x"))
        vent.configure(bg=C["blanco"])
        tk.Label(vent,text="Copia este mensaje y envíaselo al docente:",
                 font=("Segoe UI",10,"bold"),fg=C["azul_osc"],bg=C["blanco"]).pack(pady=10)
        txt=tk.Text(vent,font=("Segoe UI",10),bg=C["gris_clar"],relief="flat",bd=5,wrap="word")
        txt.insert("1.0",msg); txt.pack(fill="both",expand=True,padx=15,pady=(0,10))
        def cp():
            self.clipboard_clear(); self.clipboard_append(msg)
            messagebox.showinfo("✓","Mensaje copiado.",parent=vent)
        mk_btn(vent,"📋  Copiar mensaje completo",cp,w=28).pack(pady=8)

    def _tab_ventas(self,parent):
        f=tk.Frame(parent,bg=C["blanco"])
        top=tk.Frame(f,bg=C["azul_clar"],padx=12,pady=8); top.pack(fill="x")
        tk.Label(top,text="Registro cifrado de todas las licencias generadas · AES-256-GCM",
                 font=("Segoe UI",10),fg=C["azul_osc"],bg=C["azul_clar"]).pack(side="left")
        mk_btn(top,"🔄 Actualizar",self._recargar_ventas,w=12).pack(side="right")

        cols=("fecha","nombre","cedula","codigo","precio","tel")
        self.tree_v=ttk.Treeview(f,columns=cols,show="headings",height=25)
        for c,lbl_t,w in [("fecha","Fecha",120),("nombre","Nombre docente",200),
                           ("cedula","Cédula",100),("codigo","Código",210),
                           ("precio","Precio",80),("tel","Teléfono",110)]:
            self.tree_v.heading(c,text=lbl_t)
            self.tree_v.column(c,width=w,anchor="w" if c in ("nombre","codigo") else "center")
        sc=ttk.Scrollbar(f,orient="vertical",command=self.tree_v.yview)
        self.tree_v.configure(yscrollcommand=sc.set)
        sc.pack(side="right",fill="y"); self.tree_v.pack(fill="both",expand=True,padx=8,pady=8)
        self.lbl_tot_v=tk.Label(f,text="Total: 0",font=("Segoe UI",9),
                                 fg=C["texto_med"],bg=C["blanco"])
        self.lbl_tot_v.pack(pady=4)
        self._recargar_ventas()
        return f

    def _recargar_ventas(self):
        if not hasattr(self,"tree_v"): return
        self.tree_v.delete(*self.tree_v.get_children())
        ventas=cargar_ventas()
        for v in reversed(ventas):
            self.tree_v.insert("","end",values=(
                v.get("fecha",""),v.get("nombre",""),v.get("cedula",""),
                v.get("codigo",""),v.get("precio",""),v.get("telefono","")
            ))
        self.lbl_tot_v.config(text=f"Total: {len(ventas)} licencias generadas")

    def _tab_verificar(self,parent):
        f=tk.Frame(parent,bg=C["blanco"],padx=30,pady=20)
        tk.Label(f,text="Verifica o recupera el código de cualquier cédula.",
                 font=("Segoe UI",10),fg=C["texto_med"],bg=C["blanco"]).pack(pady=(0,18))

        # Verificar
        tk.Label(f,text="━━━  Verificar si un código es válido  ━━━",
                 font=("Segoe UI",10,"bold"),fg=C["azul_osc"],bg=C["blanco"]).pack(anchor="w",pady=(0,8))
        for etq,attr in [("N° Cédula:","var_v_ced"),("Código a verificar:","var_v_cod")]:
            tf=tk.Frame(f,bg=C["blanco"]); tf.pack(fill="x",pady=4)
            tk.Label(tf,text=etq,width=22,anchor="w",font=("Segoe UI",10),
                     fg=C["texto"],bg=C["blanco"]).pack(side="left")
            v=tk.StringVar(); setattr(self,attr,v)
            mk_entry(tf,v,w=32).pack(side="left")
        mk_btn(f,"🔍  Verificar código",self._verificar,color=C["azul_med"],w=22).pack(pady=10)
        self.lbl_result=tk.Label(f,text="",font=("Segoe UI",13,"bold"),
                                  fg=C["texto_med"],bg=C["blanco"],wraplength=700)
        self.lbl_result.pack(pady=6)

        # Recuperar
        tk.Frame(f,bg=C["gris_clar"],height=1).pack(fill="x",pady=16)
        tk.Label(f,text="━━━  Ver / recuperar el código de una cédula  ━━━",
                 font=("Segoe UI",10,"bold"),fg=C["azul_osc"],bg=C["blanco"]).pack(anchor="w",pady=(0,8))
        tf3=tk.Frame(f,bg=C["blanco"]); tf3.pack(fill="x",pady=4)
        tk.Label(tf3,text="Cédula:",width=12,anchor="w",font=("Segoe UI",10),
                 fg=C["texto"],bg=C["blanco"]).pack(side="left")
        self.var_regen=tk.StringVar()
        mk_entry(tf3,self.var_regen,w=25).pack(side="left",padx=(0,10))
        mk_btn(tf3,"Ver código",self._regen,w=14).pack(side="left")
        self.lbl_regen=tk.Label(f,text="",font=("Courier New",18,"bold"),
                                 fg=C["azul_osc"],bg=C["blanco"])
        self.lbl_regen.pack(pady=8)
        self.lbl_regen_cp=tk.Label(f,text="",font=("Segoe UI",9),
                                    fg=C["verde"],bg=C["blanco"])
        self.lbl_regen_cp.pack()
        return f

    def _verificar(self):
        ced=self.var_v_ced.get().strip()
        cod=self.var_v_cod.get().strip().upper()
        if not ced or not cod:
            messagebox.showwarning("Faltan datos","Ingresa cédula y código."); return
        esperado=generar_codigo_licencia(ced)
        if cod==esperado:
            self.lbl_result.config(text="✅  CÓDIGO VÁLIDO",fg=C["verde"])
        else:
            self.lbl_result.config(
                text=f"❌  CÓDIGO INVÁLIDO\nEl código correcto para esta cédula es:\n{esperado}",
                fg=C["rojo"])
        registrar_auditoria("VERIFICACION","VERIFICAR_CODIGO",f"Cédula {ced[:4]}***")

    def _regen(self):
        ced=self.var_regen.get().strip()
        if not ced: return
        codigo=generar_codigo_licencia(ced)
        self.lbl_regen.config(text=codigo)
        self.clipboard_clear(); self.clipboard_append(codigo)
        self.lbl_regen_cp.config(text="✓ Copiado al portapapeles")
        registrar_auditoria("VERIFICACION","RECUPERAR_CODIGO",f"Cédula {ced[:4]}***")

if __name__=="__main__":
    app=App()
    app.mainloop()
