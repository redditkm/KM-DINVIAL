import tkinter as tk
from tkinter import messagebox, scrolledtext, simpledialog
import requests
import re
import os
from PyPDF2 import PdfMerger, PdfReader
from PyPDF2.errors import PdfReadError
import sys
import unicodedata

UPLOAD_URL = "https://links.dinivalsas.com/upload.php"
BASE_URL = "https://links.dinivalsas.com/"

USER_HOME = os.path.expanduser("~")
DOWNLOADS_DIR = os.path.join(USER_HOME, "Downloads")
IMAGEN_PUNTOS = os.path.join(USER_HOME, "Documents", "TARJETA MEMBRESIA.PNG")

PDF_UNIDO = os.path.join(DOWNLOADS_DIR, "documento_unido.pdf")

# ======================
# COLORES
# ======================
COLOR_NORMAL = "#1a73e8"
COLOR_OK = "#34a853"
DURACION_VERDE_MS = 3000

# ======================
# MENSAJES
# ======================
MENSAJE_COMPROBANTE = """👋🏼 ¡Muy buen día!
Esperamos que se encuentre muy bien.
Le compartimos el enlace donde podrá consultar y descargar su comprobante de pago.
Muchas gracias por confiar en nosotros 🙏🏼
¡Que tenga un excelente y bendecido día!
👉 {enlace}


REQ // SE ENVIA MENSAJE DE COMPROBANTE DE PAGO CON LA CAPTURA 
"""

MENSAJE_ID_CARD = """👋🏼 Muy buen día.
En el siguiente enlace podrá consultar y descargar su ID CARD virtual, la cual podrá utilizar para ser atendido de la misma forma que con su ID CARD física.
Si tiene alguna duda o requiere asistencia adicional, no dude en comunicarse al 800-683-9337.
👉 {enlace} 🙏😊


REQ // SE EVIA ID CARD // SE CIERRA CASO 
"""

MENSAJE_1095A = """👋🏼 Buen día.
Le informamos que en el siguiente enlace podrá encontrar y descargar su forma 1095-A 📄🔗:
👉 {enlace}
Si tiene alguna duda o requiere asistencia adicional, no dude en contactarnos 📞💬; con gusto estaremos atentos para ayudarle 🤝✨.
Gracias por confiar en nosotros 🙏❤.


REQ // SE ENVIA FORMA 1095-A // SE CIERRA CASO
"""

MENSAJE_PUNTOS = """👋🏼 ¡Muy buen día!
En el siguiente enlace podrá encontrar su Tarjeta de Membresía.
Esta tarjeta de Membresía como cliente de nuestra agencia le hace merecedor de 100 dólares cuando logre sumar 100 puntos.
Actualmente ya cuenta con {PUNTOS} puntos 🥳🥳🥳.
Llámenos al 800-683-9337 para que pueda sumar más.
👉 {enlace}



REQ // SE ENVIA TARJETA DE MEMBRESIA // SE CIERRA CASO
"""

MENSAJE_PLAN = """👋🏼 Buen día.
Buen día, le informamos que su póliza se encuentra actualmente activa;
en el siguiente enlace podrá consultar la captura del plan vigente y los detalles de elegibilidad:👉 {enlace} Si tiene alguna duda o requiere asistencia adicional, puede contactarnos al 800-683-9337 o al correo info@kmdinival.com
. Gracias por confiar en KM DINIVAL y por ser parte de nuestra familia.

REQ // SE ENVÍA INFORMACIÓN DE PLAN/ELEGIBILIDAD // SE CIERRA CASO
"""

# ======================
# UNIR PDFs (VERSIÓN SEGURA)
# ======================
def unir_pdfs(pdfs):
    merger = PdfMerger()
    errores = []

    for pdf in pdfs:
        try:
            if not os.path.exists(pdf):
                errores.append(f"No existe: {os.path.basename(pdf)}")
                continue

            if os.path.getsize(pdf) == 0:
                errores.append(f"PDF vacío: {os.path.basename(pdf)}")
                continue

            with open(pdf, "rb") as f:
                reader = PdfReader(f, strict=False)

                if len(reader.pages) == 0:
                    errores.append(f"PDF sin páginas: {os.path.basename(pdf)}")
                    continue

                merger.append(reader)

        except PdfReadError:
            errores.append(f"PDF corrupto ignorado: {os.path.basename(pdf)}")
        except Exception as e:
            errores.append(f"Error con {os.path.basename(pdf)}: {e}")

    if not merger.pages:
        merger.close()
        raise Exception(
            "No se pudo unir ningún PDF válido.\n\n" + "\n".join(errores)
        )

    with open(PDF_UNIDO, "wb") as f:
        merger.write(f)

    merger.close()

    if errores:
        messagebox.showwarning(
            "Advertencia",
            "Algunos archivos fueron ignorados:\n\n" + "\n".join(errores)
        )

    return PDF_UNIDO

# ======================
# NORMALIZAR NOMBRES
# ======================
def _quitar_acentos(texto):
    nfkd = unicodedata.normalize("NFKD", texto)
    return "".join(c for c in nfkd if not unicodedata.combining(c))

def archivo_es_eligibility(file_path):
    nombre = os.path.basename(file_path)
    nombre_sin_acentos = _quitar_acentos(nombre).casefold()
    return ("eligibility" in nombre_sin_acentos) or ("elegibilidad" in nombre_sin_acentos)

# ======================
# OBTENER PDF FINAL
# ======================
def obtener_pdf_final():
    pdfs = [
        os.path.join(DOWNLOADS_DIR, f)
        for f in os.listdir(DOWNLOADS_DIR)
        if f.lower().endswith(".pdf")
        and os.path.isfile(os.path.join(DOWNLOADS_DIR, f))
    ]

    if not pdfs:
        return None, [], "No se encontró ningún PDF en Descargas."

    if len(pdfs) == 1:
        return pdfs[0], pdfs, None

    non_matching = [p for p in pdfs if not archivo_es_eligibility(p)]
    matching = [p for p in pdfs if archivo_es_eligibility(p)]
    ordered_pdfs = non_matching + matching

    pdf_unido = unir_pdfs(ordered_pdfs)
    return pdf_unido, pdfs + [pdf_unido], None

# ======================
# SUBIR ARCHIVO
# ======================
def subir_archivo(file_path):
    try:
        with open(file_path, "rb") as f:
            files = {"files[]": f}
            data = {"submit": "Subir Archivos"}
            response = requests.post(UPLOAD_URL, files=files, data=data)

        match = re.search(r"href='([^']+)'", response.text)
        if match:
            enlace = match.group(1)
            if not enlace.startswith("http"):
                enlace = BASE_URL + enlace.lstrip("/")
            return enlace
        return None
    except Exception as e:
        messagebox.showerror("Error", f"Error al subir archivo: {e}")
        return None

# ======================
# FEEDBACK BOTÓN
# ======================
def boton_feedback(boton):
    boton.config(bg=COLOR_OK)
    boton.after(DURACION_VERDE_MS, lambda: boton.config(bg=COLOR_NORMAL))

# ======================
# GENERAR DESDE PDF
# ======================
def generar_desde_pdf(template, boton):
    try:
        pdf_final, archivos_a_borrar, err = obtener_pdf_final()
        if err:
            messagebox.showerror("Error", err)
            return

        enlace = subir_archivo(pdf_final)
        if not enlace:
            messagebox.showerror("Error", "No se pudo generar el enlace.")
            return

        mensaje = template.format(enlace=enlace)
        text_area.delete(1.0, tk.END)
        text_area.insert(tk.END, mensaje)

        for archivo in archivos_a_borrar:
            try:
                os.remove(archivo)
            except:
                pass

        boton_feedback(boton)

    except Exception as e:
        messagebox.showerror("Error", str(e))

# ======================
# GENERAR PUNTOS
# ======================
def generar_puntos(boton):
    if not os.path.exists(IMAGEN_PUNTOS):
        messagebox.showerror("Error", "No se encontró la imagen de Tarjeta de Membresía.")
        return

    puntos = simpledialog.askstring("Puntos del cliente", "Ingrese la cantidad de puntos actuales:")
    if not puntos:
        return

    enlace = subir_archivo(IMAGEN_PUNTOS)
    if not enlace:
        messagebox.showerror("Error", "No se pudo generar el enlace.")
        return

    mensaje = MENSAJE_PUNTOS.format(enlace=enlace, PUNTOS=puntos)
    text_area.delete(1.0, tk.END)
    text_area.insert(tk.END, mensaje)

    boton_feedback(boton)

# ======================
# INTERFAZ
# ======================
root = tk.Tk()
root.title("📄 Generador de Enlaces")
root.geometry("700x520")
root.config(bg="#f5f7fa")

titulo = tk.Label(
    root,
    text="Generador de Mensajes y Enlaces",
    font=("Helvetica", 17, "bold"),
    fg="#1a73e8",
    bg="#f5f7fa"
)
titulo.pack(pady=10)

frame_botones = tk.Frame(root, bg="#f5f7fa")
frame_botones.pack(pady=10)

BTN_STYLE = {
    "bg": COLOR_NORMAL,
    "fg": "white",
    "font": ("Arial", 9, "bold"),
    "padx": 8,
    "pady": 4,
    "relief": "flat",
    "width": 18
}

btn_pago = tk.Button(frame_botones, text="PAGO",
    command=lambda: generar_desde_pdf(MENSAJE_COMPROBANTE, btn_pago), **BTN_STYLE)
btn_pago.grid(row=0, column=0, padx=6, pady=6)

btn_id = tk.Button(frame_botones, text="ID CARD",
    command=lambda: generar_desde_pdf(MENSAJE_ID_CARD, btn_id), **BTN_STYLE)
btn_id.grid(row=0, column=1, padx=6, pady=6)

btn_1095 = tk.Button(frame_botones, text="FORMA 1095-A",
    command=lambda: generar_desde_pdf(MENSAJE_1095A, btn_1095), **BTN_STYLE)
btn_1095.grid(row=1, column=0, padx=6, pady=6)

btn_puntos = tk.Button(frame_botones, text="PUNTOS",
    command=lambda: generar_puntos(btn_puntos), **BTN_STYLE)
btn_puntos.grid(row=1, column=1, padx=6, pady=6)

btn_plan = tk.Button(frame_botones, text="PLAN/ELEGIBILIDAD",
    command=lambda: generar_desde_pdf(MENSAJE_PLAN, btn_plan), **BTN_STYLE)
btn_plan.grid(row=2, column=0, columnspan=2, padx=6, pady=6)

frame_texto = tk.Frame(root, bd=2, relief="groove", bg="white")
frame_texto.pack(padx=15, pady=10, fill="both", expand=True)

text_area = scrolledtext.ScrolledText(frame_texto, font=("Consolas", 11), wrap=tk.WORD)
text_area.pack(padx=8, pady=8, fill="both", expand=True)

def on_closing():
    if messagebox.askokcancel("Salir", "¿Desea salir?"):
        root.destroy()

root.protocol("WM_DELETE_WINDOW", on_closing)

try:
    root.mainloop()
except KeyboardInterrupt:
    try:
        root.destroy()
    except:
        pass
    sys.exit(0)
