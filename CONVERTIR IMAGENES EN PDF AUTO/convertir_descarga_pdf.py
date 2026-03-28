import os
import time
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
from PIL import Image
from docx2pdf import convert
import customtkinter as ctk

CARPETA = os.path.join(os.path.expanduser("~"), "Downloads")
EXT_IMAGEN = (".jpg", ".jpeg", ".png", ".bmp", ".gif")
EXT_WORD = (".doc", ".docx")
EXT_TEMPORALES = (".crdownload", ".tmp", ".part")

activo = False
observer = None


class PDFHandler(FileSystemEventHandler):
    def on_created(self, event):
        self.procesar_archivo(event)

    def on_modified(self, event):
        self.procesar_archivo(event)

    def procesar_archivo(self, event):
        global activo
        if not activo or event.is_directory:
            return

        archivo = event.src_path
        nombre, ext = os.path.splitext(os.path.basename(archivo))
        ext = ext.lower()

        # Ignorar temporales
        if ext in EXT_TEMPORALES:
            return

        # Si aún existe su crdownload, no procesar
        if os.path.exists(archivo + ".crdownload"):
            return

        print(f"📁 Detectado archivo: {archivo}")

        for intento in range(30):
            try:
                if os.path.exists(archivo) and os.path.getsize(archivo) > 0:

                    if ext in EXT_IMAGEN:
                        self.convertir_imagen_a_pdf(archivo, nombre)
                    elif ext in EXT_WORD:
                        self.convertir_word_a_pdf(archivo)

                    return

            except Exception as e:
                print(f"Error intento {intento + 1}: {e}")

            time.sleep(1)

    def eliminar_temporal(self, archivo):
        temp = archivo + ".crdownload"
        if os.path.exists(temp):
            try:
                os.remove(temp)
                print(f"🗑️ Temporal eliminado: {temp}")
            except:
                pass

    def convertir_imagen_a_pdf(self, archivo, nombre):
        try:
            pdf_destino = os.path.join(CARPETA, f"{nombre}.pdf")

            imagen = Image.open(archivo)
            if imagen.mode != "RGB":
                imagen = imagen.convert("RGB")

            imagen.save(pdf_destino, "PDF", resolution=100.0)
            print(f"✅ Imagen convertida a PDF: {pdf_destino}")

            os.remove(archivo)
            print(f"🗑️ Imagen original eliminada")

            self.eliminar_temporal(archivo)

        except Exception as e:
            print(f"❌ Error al convertir imagen: {e}")

    def convertir_word_a_pdf(self, archivo):
        try:
            pdf_destino = os.path.join(
                CARPETA, os.path.splitext(os.path.basename(archivo))[0] + ".pdf"
            )

            convert(archivo, pdf_destino)
            print(f"✅ Word convertido a PDF: {pdf_destino}")

            os.remove(archivo)
            print(f"🗑️ Documento Word eliminado")

            self.eliminar_temporal(archivo)

        except Exception as e:
            print(f"❌ Error al convertir Word: {e}")


def iniciar_monitoreo():
    global observer
    if observer is None:
        event_handler = PDFHandler()
        observer = Observer()
        observer.schedule(event_handler, path=CARPETA, recursive=False)
        observer.start()
        print("📂 Monitoreo activado")


def detener_monitoreo():
    global observer
    if observer is not None:
        observer.stop()
        observer.join()
        observer = None
        print("🛑 Monitoreo detenido")


def actualizar_estado():
    if activo:
        estado_label.configure(text="Estado: ENCENDIDO", text_color="#43ea4b")
        toggle_btn.configure(text="APAGAR", fg_color="#f44336", hover_color="#b71c1c")
    else:
        estado_label.configure(text="Estado: APAGADO", text_color="#d32f2f")
        toggle_btn.configure(text="ENCENDER", fg_color="#43ea4b", hover_color="#339933")


def alternar_estado():
    global activo
    if activo:
        activo = False
        detener_monitoreo()
    else:
        activo = True
        iniciar_monitoreo()
    actualizar_estado()


ctk.set_appearance_mode("system")
ctk.set_default_color_theme("green")


class App(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("Auto PDF Monitor")
        self.geometry("321x300")
        self.resizable(False, False)

        fuente_titulo = ("Segoe UI Semibold", 25)
        fuente_estado = ("Segoe UI", 16, "bold")
        fuente_btn = ("Segoe UI", 16, "bold")
        fuente_sub = ("Segoe UI", 13)
        fuente_footer = ("Segoe UI", 11, "italic")
        fuente_c = ("Segoe UI", 13, "bold")

        box = ctk.CTkFrame(self, fg_color="#ffffff", corner_radius=18)
        box.pack(padx=26, pady=26, fill="both", expand=True)

        header = ctk.CTkLabel(box, text="Auto PDF Monitor", font=fuente_titulo, text_color="#287a4f")
        header.pack(pady=(18, 2))

        sub = ctk.CTkLabel(box, text="Convierte imágenes y Word a PDF.", font=fuente_sub, text_color="#888")
        sub.pack(pady=(0, 24))

        global estado_label
        estado_label = ctk.CTkLabel(box, text="Estado: APAGADO", font=fuente_estado, text_color="#d32f2f")
        estado_label.pack(pady=(0, 20))

        global toggle_btn
        toggle_btn = ctk.CTkButton(
            box, text="ENCENDER",
            font=fuente_btn,
            fg_color="#43ea4b",
            hover_color="#339933",
            width=180,
            height=48,
            command=alternar_estado,
            corner_radius=17
        )
        toggle_btn.pack(pady=(0, 10))

        footer_frame = ctk.CTkFrame(self, fg_color="transparent")
        footer_frame.pack(side="bottom", fill="x")

        footer_row = ctk.CTkFrame(footer_frame, fg_color="transparent")
        footer_row.pack(pady=(9, 13))

        footer_c = ctk.CTkLabel(footer_row, text="©", font=fuente_c, text_color="#888")
        footer_c.pack(side="left", padx=(0, 2))

        footer_label = ctk.CTkLabel(
            footer_row,
            text="Hecho por FELIPE KM DINIVAL",
            font=fuente_footer,
            text_color="#888"
        )
        footer_label.pack(side="left")

        actualizar_estado()


if __name__ == "__main__":
    app = App()
    app.mainloop()
