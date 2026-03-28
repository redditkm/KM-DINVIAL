#!/usr/bin/env python3
"""
recortes.py

Versión con GUI y listener del portapapeles que incluye fallbacks cuando
pywin32 no expone AddClipboardFormatListener. Requiere:
  pip install pillow pywin32

Ejecución sin consola:
  pythonw.exe recortes.py
"""
import os
import io
import time
import hashlib
import threading
import ctypes
from datetime import datetime
from PIL import Image, ImageGrab
import tkinter as tk
from tkinter import filedialog, messagebox
from pathlib import Path

# Obtener carpeta Documentos del usuario actual (Windows)
def get_documents_path():
    return Path.home() / "Documents"

try:
    import win32con
    import win32gui
    import win32api
    WIN32_AVAILABLE = True
except Exception:
    WIN32_AVAILABLE = False
    raise RuntimeError("Este script requiere pywin32 (pip install pywin32) y funcionar en Windows.")

# Fallback para WM_CLIPBOARDUPDATE (constante Windows)
try:
    WM_CLIPBOARDUPDATE = win32con.WM_CLIPBOARDUPDATE
except AttributeError:
    WM_CLIPBOARDUPDATE = 0x031D

# Configuración
SHOW_INFO_ON_SAVE = False    # <- notificación desactivada (solo guardar)
AUTO_CLEAR_CLIPBOARD = False
TOPMOST_DIALOG = True
DUPLICATE_DEBOUNCE_SECONDS = 0.5
POLL_INTERVAL_SECONDS = 0.5  # usado solo si se cae al polling

# Rutas fijas solicitadas
BASE_DOCS = get_documents_path()

TARGET_PATHS = {
    "TARJETA MEMBRESIA": str(BASE_DOCS / "TARJETA MEMBRESIA.PNG"),
    "BIENVENIDA WPP":   str(BASE_DOCS / "BIENVENIDA WPP.PNG"),
    "BIENVENIDA SUPLE": str(BASE_DOCS / "BIENVENIDA SUPLE.PNG"),
    "CUMPLEAÑOS":       str(BASE_DOCS / "CUMPLEAÑOS.PNG"),
}


def image_hash(img: Image.Image) -> str:
    b = io.BytesIO()
    img.save(b, format="PNG")
    return hashlib.sha256(b.getvalue()).hexdigest()

def ask_save_path(default_ext=".png"):
    root = tk.Tk()
    root.withdraw()
    if TOPMOST_DIALOG:
        root.attributes("-topmost", True)
    f = filedialog.asksaveasfilename(
        defaultextension=default_ext,
        filetypes=[
            ("PNG image", "*.png"),
            ("JPEG image", "*.jpg;*.jpeg"),
            ("Bitmap", "*.bmp"),
            ("All files", "*.*"),
        ],
        title="Guardar recorte como..."
    )
    root.destroy()
    return f

def ensure_dir_for_file(path):
    d = os.path.dirname(path)
    if d and not os.path.exists(d):
        os.makedirs(d, exist_ok=True)

class ClipListener:
    """
    Escucha el portapapeles. Intenta usar AddClipboardFormatListener; si no está
    disponible, usa ctypes para llamar a user32.AddClipboardFormatListener; si
    eso falla, cae a un polling loop.
    """
    def __init__(self, target_getter):
        self.last_hash = None
        self.last_time = 0
        self._target_getter = target_getter
        self._running = True
        self._use_polling = False
        self._poll_stop_event = None
        self._register_window()

    def _register_window(self):
        # Registrar clase de ventana y crear window message-only
        message_map = {
            win32con.WM_DESTROY: self.on_destroy,
            WM_CLIPBOARDUPDATE: self.on_clipboard_update,
        }

        wc = win32gui.WNDCLASS()
        hinst = wc.hInstance = win32gui.GetModuleHandle(None)
        wc.lpszClassName = "PythonClipboardListener"
        wc.lpfnWndProc = message_map
        try:
            classAtom = win32gui.RegisterClass(wc)
        except Exception:
            classAtom = wc.lpszClassName

        self.hwnd = win32gui.CreateWindowEx(
            0, classAtom, "ClipListenerWindow", 0,
            0, 0, 0, 0, 0, 0, hinst, None
        )

        # Intentar registrar AddClipboardFormatListener por varios medios
        registered = False
        # 1) si win32gui lo expone, usarlo
        if hasattr(win32gui, "AddClipboardFormatListener"):
            try:
                win32gui.AddClipboardFormatListener(self.hwnd)
                registered = True
            except Exception:
                registered = False

        # 2) intentar via ctypes -> user32.AddClipboardFormatListener
        if not registered:
            try:
                user32 = ctypes.windll.user32
                res = user32.AddClipboardFormatListener(self.hwnd)
                if res != 0:
                    registered = True
            except Exception:
                registered = False

        # 3) Si todo falla, activar polling (fallback)
        if not registered:
            self._use_polling = True
            self._start_poller()
        else:
            self._use_polling = False

    def _start_poller(self):
        self._poll_stop_event = threading.Event()
        t = threading.Thread(target=self._poll_loop, daemon=True)
        t.start()

    def _poll_loop(self):
        while not self._poll_stop_event.is_set():
            try:
                self._check_clipboard_and_handle()
            except Exception:
                pass
            self._poll_stop_event.wait(POLL_INTERVAL_SECONDS)

    def _check_clipboard_and_handle(self):
        try:
            clip = ImageGrab.grabclipboard()
        except Exception:
            clip = None

        if not isinstance(clip, Image.Image):
            return

        try:
            h = image_hash(clip)
        except Exception:
            h = None

        now = time.time()
        if not h:
            return
        if not (h != self.last_hash or (now - self.last_time) > DUPLICATE_DEBOUNCE_SECONDS):
            return

        # actualizar debounce
        self.last_hash = h
        self.last_time = now

        # obtener path destino
        try:
            path = self._target_getter()
        except Exception:
            path = None

        if not path:
            path = ask_save_path(".png")

        if path:
            ext = path.lower().rsplit(".", 1)[-1] if "." in path else "png"
            fmt = "PNG" if ext in ("png",) else ("JPEG" if ext in ("jpg", "jpeg") else "BMP")
            try:
                ensure_dir_for_file(path)
                img_to_save = clip.convert("RGB") if fmt in ("JPEG",) else clip
                img_to_save.save(path, fmt)
                # NOTA: notificación de guardado desactivada (SHOW_INFO_ON_SAVE=False)
                if AUTO_CLEAR_CLIPBOARD:
                    try:
                        import win32clipboard
                        win32clipboard.OpenClipboard()
                        win32clipboard.EmptyClipboard()
                        win32clipboard.CloseClipboard()
                    except Exception:
                        pass
            except Exception as e:
                try:
                    root = tk.Tk()
                    root.withdraw()
                    if TOPMOST_DIALOG:
                        root.attributes("-topmost", True)
                    messagebox.showerror("Error al guardar", str(e))
                    root.destroy()
                except Exception:
                    pass

    def on_clipboard_update(self, hwnd, msg, wparam, lparam):
        # Solo usado si AddClipboardFormatListener está activo
        self._check_clipboard_and_handle()
        return 0

    def on_destroy(self, hwnd, msg, wparam, lparam):
        # Quitar listener o detener poller
        if not self._use_polling:
            # intentar win32gui.RemoveClipboardFormatListener, sino ctypes
            try:
                if hasattr(win32gui, "RemoveClipboardFormatListener"):
                    win32gui.RemoveClipboardFormatListener(self.hwnd)
                else:
                    user32 = ctypes.windll.user32
                    user32.RemoveClipboardFormatListener(self.hwnd)
            except Exception:
                pass
        else:
            if self._poll_stop_event:
                self._poll_stop_event.set()
        try:
            win32gui.PostQuitMessage(0)
        except Exception:
            pass
        self._running = False
        return 0

    def run(self):
        try:
            if self._use_polling:
                # Si estamos en polling no necesitamos PumpMessages; mantener hilo vivo
                while self._running:
                    time.sleep(0.1)
            else:
                win32gui.PumpMessages()
        finally:
            try:
                win32gui.DestroyWindow(self.hwnd)
            except Exception:
                pass

    def stop(self):
        # Cerrar de manera limpia: enviar WM_DESTROY o detener poller
        if self._use_polling:
            if self._poll_stop_event:
                self._poll_stop_event.set()
        else:
            try:
                win32gui.PostMessage(self.hwnd, win32con.WM_DESTROY, 0, 0)
            except Exception:
                # como fallback, llamar DestroyWindow directo
                try:
                    win32gui.DestroyWindow(self.hwnd)
                except Exception:
                    pass

class ModeState:
    def __init__(self):
        self.lock = threading.Lock()
        self.active_mode = None

    def set_mode(self, mode_key):
        with self.lock:
            self.active_mode = mode_key

    def clear_mode(self):
        with self.lock:
            self.active_mode = None

    def get_path(self):
        with self.lock:
            if self.active_mode and self.active_mode in TARGET_PATHS:
                return TARGET_PATHS[self.active_mode]
            return None

def build_gui_and_run(mode_state, stop_callback):
    root = tk.Tk()
    root.title("Save Snip Quick Modes")

    lbl = tk.Label(root, text="Modo activo: Ninguno", font=("Segoe UI", 10))
    lbl.pack(padx=10, pady=(10, 6))

    btn_frame = tk.Frame(root)
    btn_frame.pack(padx=10, pady=6)

    buttons = {}

    def update_label():
        m = mode_state.active_mode
        lbl.config(text=f"Modo activo: {m if m else 'Ninguno'}")

    def on_mode_press(key):
        if mode_state.active_mode == key:
            mode_state.clear_mode()
        else:
            mode_state.set_mode(key)
        for k, b in buttons.items():
            if mode_state.active_mode == k:
                b.config(relief=tk.SUNKEN, bg="lightgreen")
            else:
                b.config(relief=tk.RAISED, bg=root.cget("bg"))
        update_label()

    for k in TARGET_PATHS.keys():
        b = tk.Button(btn_frame, text=k, width=18, command=lambda k=k: on_mode_press(k))
        b.pack(side=tk.LEFT, padx=4, pady=4)
        buttons[k] = b

    ctrl_frame = tk.Frame(root)
    ctrl_frame.pack(fill=tk.X, padx=10, pady=(6,10))

    def on_deactivate():
        mode_state.clear_mode()
        for b in buttons.values():
            b.config(relief=tk.RAISED, bg=root.cget("bg"))
        update_label()

    def on_quit():
        try:
            stop_callback()
        except Exception:
            pass
        root.after(150, root.destroy)

    btn_off = tk.Button(ctrl_frame, text="Desactivar", command=on_deactivate)
    btn_off.pack(side=tk.LEFT, padx=4)
    btn_quit = tk.Button(ctrl_frame, text="Salir", command=on_quit)
    btn_quit.pack(side=tk.RIGHT, padx=4)

    update_label()
    root.mainloop()

def main():
    print("Iniciado watcher rápido del portapapeles con GUI.")
    mode_state = ModeState()

    def current_target_getter():
        return mode_state.get_path()

    listener = ClipListener(current_target_getter)
    listener_thread = threading.Thread(target=listener.run, daemon=True)
    listener_thread.start()

    def stop_listener():
        listener.stop()
        listener_thread.join(timeout=1.0)

    try:
        build_gui_and_run(mode_state, stop_listener)
    except KeyboardInterrupt:
        print("Saliendo por Ctrl+C...")
    finally:
        stop_listener()

if __name__ == "__main__":
    main()