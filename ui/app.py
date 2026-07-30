import customtkinter as ctk
from PIL import Image
import os

from ui.views.convertidor import ConvertidorView
from ui.views.configuracion import ConfiguracionView
from ui.views.informacion import InformacionView

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

import sys

def ruta_recurso(ruta_relativa):
    """Devuelve la ruta absoluta a un recurso, tanto en desarrollo como empaquetado con PyInstaller."""
    try:
        base_path = sys._MEIPASS  # carpeta temporal que usa PyInstaller en --onefile
    except AttributeError:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, ruta_relativa)

def cargar_icono(nombre, size=(20, 20)):
    ruta = ruta_recurso(os.path.join("ui", "assets", "icons", nombre))
    img_blanca = Image.open(ruta).convert("RGBA")

    # Generamos la versión negra invirtiendo el color (no el alfa) para modo claro
    r, g, b, a = img_blanca.split()
    img_negra = Image.merge("RGBA", (
        Image.eval(r, lambda x: 255 - x),
        Image.eval(g, lambda x: 255 - x),
        Image.eval(b, lambda x: 255 - x),
        a,
    ))

    return ctk.CTkImage(light_image=img_negra, dark_image=img_blanca, size=size)

class App(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("Lumen")
        self.geometry("900x600")
        self.minsize(700, 500)

        # ── Barra de menú superior ──────────────────────────────────────
        self.menu_bar = ctk.CTkFrame(self, height=50, corner_radius=0)
        self.menu_bar.pack(side="top", fill="x")

        # ── Contenedor de vistas ─────────────────────────────────────────
        self.container = ctk.CTkFrame(self, fg_color="transparent")
        self.container.pack(expand=True, fill="both")

        # Vistas disponibles: nombre -> (clase, icono)
        self.views = {
            "Convertidor": (ConvertidorView, "convertidor.png"),
            "Configuración": (ConfiguracionView, "configuracion.png"),
            "Información": (InformacionView, "informacion.png"),
        }

        self.buttons = {}
        self.frames = {}

        for name, (view_class, icon_name) in self.views.items():
            icono = cargar_icono(icon_name)

            btn = ctk.CTkButton(
                self.menu_bar,
                text=name,
                image=icono,
                compound="left",
                font=ctk.CTkFont(size=13, weight="bold"),
                fg_color="transparent",
                text_color=("gray20", "gray80"),
                hover_color=("gray80", "gray22"),
                corner_radius=8,
                height=34,
                command=lambda n=name: self.show_view(n),
            )
            btn.pack(side="left", padx=4, pady=10)
            self.buttons[name] = btn

            frame = view_class(self.container)
            self.frames[name] = frame

        self.show_view("Convertidor")

    def show_view(self, name):
        # Oculta todas
        for frame in self.frames.values():
            frame.pack_forget()

        # Muestra la seleccionada
        self.frames[name].pack(expand=True, fill="both")

        # Resalta el botón activo
        for n, btn in self.buttons.items():
            btn.configure(fg_color=("gray75", "gray30") if n == name else "transparent")
