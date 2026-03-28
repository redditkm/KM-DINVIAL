import tkinter as tk
from tkinter import messagebox, scrolledtext
import requests
import re
import os
import glob
from PyPDF2 import PdfMerger, PdfReader

UPLOAD_URL = "https://links.dinivalsas.com/upload.php"
BASE_URL = "https://links.dinivalsas.com/"

# Carpetas
DESCARGAS = os.path.join(os.path.expanduser("~"), "Downloads")
CARPETA_BIENVENIDA = DESCARGAS

# Mensajes
MENSAJE_BIENVENIDA = """¡Buen día! Nos llena de alegría que ahora forme parte de nuestra gran familia en la agencia KM DINIVAL 🎉🤗. Estamos aquí para apoyarle en todo lo que necesite y hacer que su experiencia con nosotros sea increíble. Si tiene alguna pregunta, no dude en contactarnos.

En este documento encontrará su carta de bienvenida, la captura de su plan y su carta de elegibilidad a través del siguiente enlace:
{enlace_bienvenida}

¡Trabajamos para usted!
Por favor, confirme si ha recibido toda la información.
¡Estamos a su disposición!
"""

MENSAJE_TARJETA = """BIENVENIDA WS // SE ENVIA MENSAJE CON EL LINK EN DONDE SE ENCUENTRA LA CAPTURA DEL PLAN // LA CARTA DE BIENVENIDA // LA ELIGIBILIDAD // SE ADJUNTA CAPTURA // SE CIERRA CASO"""

# ---------- FUNCIONES PDF BIENVENIDA ----------

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
    except:
        return "Bienvenida"

def unir_pdfs_bienvenidas():
    archivos = limpiar_archivos_temporales(glob.glob(os.path.join(DESCARGAS, "*.pdf")))

    pdf_bienvenida = next(
        (f for f in archivos if "CARTA BIENVENIDA KM COMPLETA" in f.upper()),
        None
    )
    pdfs_plan = [f for f in archivos if "PLAN" in f.upper()]
    otros = [f for f in archivos if f != pdf_bienvenida and f not in pdfs_plan]

    if not pdf_bienvenida or not pdfs_plan:
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
    except:
        pass

# ---------- FUNCIONES DE SUBIDA ----------

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
                enlace = BASE_URL + enlace_relativo.lstrip("/")
            else:
                enlace = enlace_relativo
            return enlace
        return None
    except Exception as e:
        messagebox.showerror("Error", f"Ocurrió un error:\n{str(e)}")
        return None

def borrar_descargas():
    try:
        archivos = glob.glob(os.path.join(CARPETA_BIENVENIDA, "*"))
        for archivo in archivos:
            try:
                os.remove(archivo)
            except:
                pass
    except:
        pass

# ---------- BOTÓN PRINCIPAL ----------

def subir_mensajes():
    # 1) Primero unir PDFs
    unir_pdfs_bienvenidas()

    # 2) Luego continuar con lo anterior
    archivos = glob.glob(os.path.join(CARPETA_BIENVENIDA, "Bienvenida*"))
    if not archivos:
        messagebox.showerror("Error", f"No se encontró ningún archivo que empiece con 'Bienvenida' en {CARPETA_BIENVENIDA}")
        return

    archivo_bienvenida = archivos[0]
    enlace_bienvenida = subir_archivo(archivo_bienvenida)

    if enlace_bienvenida:
        mensaje_final_bienvenida = MENSAJE_BIENVENIDA.format(enlace_bienvenida=enlace_bienvenida)
        mensaje_final = mensaje_final_bienvenida + "\n\n" + MENSAJE_TARJETA

        text_area.delete(1.0, tk.END)
        text_area.insert(tk.END, mensaje_final)

        borrar_descargas()

        btn_subir.config(bg="green")
        root.after(2000, lambda: btn_subir.config(bg="#1a73e8"))
    else:
        messagebox.showerror("Error", "No se pudo obtener el enlace de subida.")

# ==============================
# INTERFAZ
# ==============================
root = tk.Tk()
root.title("✨ Generador de Mensajes KM DINIVAL ✨")
root.geometry("550x550")
root.config(bg="#f5f7fa")

titulo = tk.Label(
    root, text="Generador de Mensajes KM DINIVAL",
    font=("Helvetica", 18, "bold"), fg="#1a73e8", bg="#f5f7fa"
)
titulo.pack(pady=15)

btn_subir = tk.Button(
    root, text="📤 Crear Mensaje", command=subir_mensajes,
    bg="#1a73e8", fg="white", font=("Arial", 14, "bold"),
    relief="flat", padx=20, pady=10
)
btn_subir.pack(pady=10)

frame_texto = tk.Frame(root, bd=2, relief="groove", bg="white")
frame_texto.pack(padx=20, pady=15, fill="both", expand=True)

text_area = scrolledtext.ScrolledText(
    frame_texto, width=95, height=25, font=("Consolas", 11),
    wrap=tk.WORD, relief="flat", bg="#ffffff", fg="#333333"
)
text_area.pack(padx=10, pady=10, fill="both", expand=True)

root.mainloop()
