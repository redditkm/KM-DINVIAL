import os
import glob
import re
import requests
from tkinter import messagebox, scrolledtext, StringVar
import customtkinter as ctk
from PyPDF2 import PdfMerger, PdfReader

# Rutas
DESCARGAS = os.path.join(os.path.expanduser("~"), "Downloads")
DESKTOP = os.path.join(os.path.expanduser("~"), "Desktop")
BIENVENIDAS = os.path.join(DESKTOP, "BIENVENIDAS")

PLANES_FL_TX = {
    "DVH": "DENTAL WISE MAX DVH EN ESPAÑOL.pdf",
    "DENTAL": "DENTAL WISE TX FL ESPAÑOL.pdf",
    "VISUAL": "VisionWise FL TX ESPAÑOL.pdf"
}

PLANES_VA = {
    "DVH": "PRIME DVH VA ESPAÑOL.pdf",
    "DENTAL": "DENTAL PPO VA ESPAÑOL.pdf",
    "VISUAL": "PREMIER VISION VA ESPAÑOL.pdf"
}

UPLOAD_URL = "https://links.dinivalsas.com/upload.php"
BASE_URL = "https://links.dinivalsas.com/"

def limpiar_archivos_temporales(archivos):
    return [f for f in archivos if not os.path.basename(f).startswith("~$")]

def extraer_primer_nombre(pdf_path):
    try:
        reader = PdfReader(pdf_path)
        texto = ""
        for page in reader.pages:
            texto += page.extract_text() or ""
        match = re.search(r"Estimado\(a\)\s+([^.]+)\.", texto)
        if match:
            nombres = match.group(1).strip().split()
            primer_nombre = nombres[0] if nombres else "Bienvenida"
            primer_nombre = re.sub(r'[^A-Za-zÁÉÍÓÚáéíóúÑñüÜ]', '', primer_nombre)
            return primer_nombre.capitalize()
        return "Bienvenida"
    except Exception as e:
        print("Error al leer nombre del PDF:", e)
        return "Bienvenida"

def subir_archivo(file_path):
    try:
        with open(file_path, "rb") as f:
            files = {"files[]": f}
            data = {"submit": "Subir Archivos"}
            response = requests.post(UPLOAD_URL, files=files, data=data)
        match = re.search(r"href='([^']+)'", response.text)
        if match:
            enlace_relativo = match.group(1)
            if not enlace_relativo.startswith("http"):
                return BASE_URL + enlace_relativo.lstrip("/")
            return enlace_relativo
        return None
    except Exception as e:
        messagebox.showerror("Error", f"Error al subir archivo:\n{e}")
        return None

def mostrar_mensaje_final(mensaje):
    ventana_msg = ctk.CTkToplevel()
    ventana_msg.title("Mensaje de Bienvenida")
    ventana_msg.geometry("600x400")
    ventana_msg.resizable(True, True)

    text_area = scrolledtext.ScrolledText(
        ventana_msg,
        wrap="word",
        font=("Consolas", 11),
        bg="#ffffff",
        fg="#333333"
    )
    text_area.pack(expand=True, fill="both", padx=10, pady=10)
    text_area.insert("end", mensaje)

    def copiar():
        ventana_msg.clipboard_clear()
        ventana_msg.clipboard_append(mensaje)
        ventana_msg.update()
        messagebox.showinfo("Copiado", "El mensaje se copió al portapapeles.")

    btn_copiar = ctk.CTkButton(ventana_msg, text="Copiar al portapapeles", command=copiar)
    btn_copiar.pack(pady=5)

def confirmar_datos(tratamiento, nombre, plan, fecha, poliza, archivo_bienvenida, extra_pdf):
    salida = os.path.join(DESCARGAS, f"Bienvenida Suplementario.pdf")
    try:
        merger = PdfMerger()
        merger.append(archivo_bienvenida)
        merger.append(extra_pdf)
        merger.write(salida)
        merger.close()
        enlace = subir_archivo(salida)
        if enlace:
            mensaje = (
                f"{tratamiento} {nombre} nos complace brindarle la bienvenida a su \"{plan}\" ✨ . "
                f"Activo desde el día {fecha}, a continuación le explicamos el proceso para solicitar "
                f"alguna cita y la información que debe tener en cuenta al momento de asistir a ellas 🩺.  "
                f"Recuerde que su número de póliza es: {poliza} Si tiene alguna duda o requerimiento por favor "
                f"no dude en contactarse al  800-683-9337 o aquí mismo nos puede dejar el mensaje para atenderlo "
                f"lo antes posible. ☺🙏\n\n\n"
                f"En el siguiente enlace encontrará la carta de bienvenida con toda la información de su poliza:  {enlace}"
            )
            mostrar_mensaje_final(mensaje)
    except Exception as e:
        messagebox.showerror("Error", f"Error en Confirmar:\n{e}")

def unir_pdfs_bienvenida():
    estado = entrada_estado.get().strip().upper()
    if not estado:
        estado = "TX"  # por defecto

    archivos = limpiar_archivos_temporales(glob.glob(os.path.join(DESCARGAS, "*.pdf")))
    seleccionados = []
    for archivo in archivos:
        nombre = os.path.basename(archivo).upper()
        if "BIENVENIDA PLAN DENTAL" in nombre:
            seleccionados.append((archivo, "DENTAL"))
        elif "BIENVENIDA PLAN DVH" in nombre:
            seleccionados.append((archivo, "DVH"))
        elif "BIENVENIDA PLAN VISUAL" in nombre:
            seleccionados.append((archivo, "VISUAL"))

    if not seleccionados:
        estado_label.configure(text="No hay elementos para BIENVENIDA SUPLE", text_color="#e44")
        return

    archivo_bienvenida, tipo = seleccionados[0]
    extra_pdf = os.path.join(BIENVENIDAS, PLANES_VA[tipo] if estado == "VA" else PLANES_FL_TX[tipo])
    if not os.path.exists(extra_pdf):
        estado_label.configure(text=f"No se encontró:\n{os.path.basename(extra_pdf)}", text_color="#e44")
        return

    # Modal responsive con grid
    modal = ctk.CTkToplevel()
    modal.title("Datos de Bienvenida")
    modal.geometry("380x400")
    modal.resizable(True, True)

    modal.transient(app)
    modal.lift()
    modal.grab_set()
    modal.focus_force()

    for i in range(10):
        modal.grid_rowconfigure(i, weight=1)
    modal.grid_columnconfigure(0, weight=1)

    fuente_lbl = ("Segoe UI", 12)

    ctk.CTkLabel(modal, text="Tratamiento (Sr / Sra)", font=fuente_lbl).grid(row=0, column=0, pady=(8, 2))
    tratamiento_var = StringVar(value="Sra")
    frame_tratamiento = ctk.CTkFrame(modal)
    frame_tratamiento.grid(row=1, column=0, pady=(0, 8))
    ctk.CTkRadioButton(frame_tratamiento, text="Sr", variable=tratamiento_var, value="Sr").pack(side="left", padx=10)
    ctk.CTkRadioButton(frame_tratamiento, text="Sra", variable=tratamiento_var, value="Sra").pack(side="left", padx=10)

    ctk.CTkLabel(modal, text="Nombre Titular", font=fuente_lbl).grid(row=2, column=0, pady=(8, 2))
    entry_nombre = ctk.CTkEntry(modal, width=240)
    entry_nombre.grid(row=3, column=0, pady=(0, 8))

    ctk.CTkLabel(modal, text="Nombre del Plan", font=fuente_lbl).grid(row=4, column=0, pady=(8, 2))
    entry_plan = ctk.CTkEntry(modal, width=240)
    entry_plan.grid(row=5, column=0, pady=(0, 8))

    ctk.CTkLabel(modal, text="Fecha de Efectividad", font=fuente_lbl).grid(row=6, column=0, pady=(8, 2))
    entry_fecha = ctk.CTkEntry(modal, width=240)
    entry_fecha.grid(row=7, column=0, pady=(0, 8))

    ctk.CTkLabel(modal, text="Número de Póliza", font=fuente_lbl).grid(row=8, column=0, pady=(8, 2))
    entry_poliza = ctk.CTkEntry(modal, width=240)
    entry_poliza.grid(row=9, column=0, pady=(0, 8))

    def on_confirm():
        nombre = entry_nombre.get().strip()
        plan = entry_plan.get().strip()
        fecha = entry_fecha.get().strip()
        poliza = entry_poliza.get().strip()
        tratamiento = tratamiento_var.get()
        if not all([nombre, plan, fecha, poliza]):
            messagebox.showerror("Error", "Todos los campos son obligatorios")
            return
        modal.destroy()
        confirmar_datos(tratamiento, nombre, plan, fecha, poliza, archivo_bienvenida, extra_pdf)

    ctk.CTkButton(modal, text="Confirmar", command=on_confirm, width=160, height=35).grid(row=10, column=0, pady=12)

def unir_pdfs_bienvenidas():
    archivos = limpiar_archivos_temporales(glob.glob(os.path.join(DESCARGAS, "*.pdf")))
    pdf_bienvenida = next((f for f in archivos if "CARTA BIENVENIDA KM COMPLETA" in f.upper()), None)
    pdfs_plan = [f for f in archivos if "PLAN" in f.upper()]
    otros = [f for f in archivos if f != pdf_bienvenida and f not in pdfs_plan]
    if not pdf_bienvenida or not pdfs_plan:
        estado_label.configure(text="Faltan elementos para BIENVENIDAS", text_color="#e44")
        return
    orden_final = [pdf_bienvenida] + pdfs_plan + otros
    primer_nombre = extraer_primer_nombre(pdf_bienvenida)
    salida = os.path.join(DESCARGAS, f"Bienvenida {primer_nombre}.pdf")
    try:
        merger = PdfMerger()
        for pdf in orden_final:
            merger.append(pdf)
        merger.write(salida)
        merger.close()
        estado_label.configure(text=f"PDF creado:\n{os.path.basename(salida)}", text_color="#43ea4b")
    except Exception as e:
        estado_label.configure(text=f"Error al guardar:\n{e}", text_color="#e44")

def unir_todos_los_pdfs():
    archivos = limpiar_archivos_temporales(glob.glob(os.path.join(DESCARGAS, "*.pdf")))
    if not archivos:
        estado_label.configure(text="No hay PDFs en Descargas", text_color="#e44")
        return
    elegibles = []
    otros = []
    for pdf in archivos:
        nombre = os.path.basename(pdf).lower()
        if "eligibility" in nombre or "elegibilidad" in nombre:
            elegibles.append(pdf)
        else:
            otros.append(pdf)
    orden_final = otros + elegibles
    salida = os.path.join(DESCARGAS, "PDF UNIDO.pdf")
    try:
        merger = PdfMerger()
        for pdf in orden_final:
            merger.append(pdf)
        merger.write(salida)
        merger.close()
        estado_label.configure(text=f"PDF creado:\n{os.path.basename(salida)}", text_color="#43ea4b")
    except Exception as e:
        estado_label.configure(text=f"Error al unir:\n{e}", text_color="#e44")

# GUI
ctk.set_appearance_mode("system")
ctk.set_default_color_theme("green")

class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Unir PDFs en Descargas")
        self.geometry("300x270")
        self.resizable(False, False)

        fuente_titulo = ("Segoe UI", 22, "bold")
        fuente_btn = ("Segoe UI", 16, "bold")
        fuente_estado = ("Segoe UI", 14, "italic")

        ctk.CTkLabel(self, text="Unir PDFs - Opciones", font=fuente_titulo, text_color="#287a4f").pack(pady=(18, 12))

        ctk.CTkButton(
            self,
            text="BIENVENIDA",
            font=fuente_btn,
            fg_color="#43ea4b",
            hover_color="#339933",
            text_color="#fff",
            width=180,
            height=40,
            corner_radius=17,
            command=unir_pdfs_bienvenidas
        ).pack(pady=4)

        ctk.CTkButton(
            self,
            text="BIENVENIDA SUPLE",
            font=fuente_btn,
            fg_color="#2e82ff",
            hover_color="#205ecf",
            text_color="#fff",
            width=180,
            height=40,
            corner_radius=17,
            command=unir_pdfs_bienvenida
        ).pack(pady=4)

        ctk.CTkButton(
            self,
            text="UNIR PDF",
            font=fuente_btn,
            fg_color="#ff9e00",
            hover_color="#cc7a00",
            text_color="#fff",
            width=180,
            height=40,
            corner_radius=17,
            command=unir_todos_los_pdfs
        ).pack(pady=4)

        global entrada_estado
        entrada_estado = ctk.CTkEntry(self, placeholder_text="ESTADO", width=64, justify="center")
        entrada_estado.pack(pady=(6, 4))
        entrada_estado.bind("<FocusOut>", self.reiniciar_placeholder)

        global estado_label
        estado_label = ctk.CTkLabel(self, text="", font=fuente_estado, text_color="#888")
        estado_label.pack(pady=(20, 0))

    def reiniciar_placeholder(self, event):
        if entrada_estado.get().strip() == "":
            entrada_estado.delete(0, "end")
            entrada_estado.configure(placeholder_text="ESTADO")

if __name__ == "__main__":
    app = App()
    app.mainloop()
