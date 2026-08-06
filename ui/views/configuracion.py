import os
import sys
import customtkinter as ctk
from tkinter import filedialog
from PIL import Image

from functions.config_manager import guardar_valor, obtener_valor, eliminar_valor

CLAVE_RUTA_STOCK_FACIL = "ruta_exportacion_stock_facil"
CLAVE_RUTA_VENDEDOR = "ruta_exportacion_vendedor"
CLAVE_API_KEY_GEMINI = "api_key_gemini"
CLAVE_API_KEY_BORRADOR_FONDOS = "api_key_borrador_fondos"


def _ruta_recurso(ruta_relativa):
    """Devuelve la ruta absoluta a un recurso, tanto en desarrollo como empaquetado con PyInstaller."""
    try:
        base_path = sys._MEIPASS
    except AttributeError:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, ruta_relativa)


def _cargar_logo(nombre, size=(20, 20)):
    """Carga una imagen de assets/icons SIN invertir colores (para logos de marca, a diferencia de los íconos monocromáticos del menú)."""
    ruta = _ruta_recurso(os.path.join("ui", "assets", "icons", nombre))
    imagen = Image.open(ruta).convert("RGBA")
    return ctk.CTkImage(light_image=imagen, dark_image=imagen, size=size)


class ConfiguracionView(ctk.CTkFrame):
    def __init__(self, master):
        super().__init__(master, fg_color="transparent")

        contenido = ctk.CTkFrame(self, fg_color="transparent", width=480)
        contenido.pack(expand=True, pady=20)

        self._crear_titulo(contenido)
        self._crear_divisor(contenido)

        self.boton_stock_facil = self._crear_fila_ruta(
            contenido,
            titulo="Ruta de exportación para \"Stock Fácil\"",
            clave_config=CLAVE_RUTA_STOCK_FACIL,
        )
        self.boton_vendedor = self._crear_fila_ruta(
            contenido,
            titulo="Ruta de exportación para \"Vendedor\"",
            clave_config=CLAVE_RUTA_VENDEDOR,
        )

        self._crear_divisor(contenido)

        self._crear_fila_api_key(
            contenido,
            titulo="API Key de Gemini",
            clave_config=CLAVE_API_KEY_GEMINI,
            icono_nombre="gemini.png",
            texto_boton="Establecer API Key de Gemini",
            texto_dialogo="Pegá tu API Key de Gemini:",
        )

        self._crear_fila_api_key(
            contenido,
            titulo="API Key de Borrador ",
            clave_config=CLAVE_API_KEY_BORRADOR_FONDOS,
            icono_nombre="fondo_borrar.png",
            texto_boton="Establecer API Key de Borrador",
            texto_dialogo="Pegá tu API Key para borrar fondos:",
        )

    # ==================================================================
    #  TÍTULO Y DIVISOR
    # ==================================================================

    def _crear_titulo(self, padre):
        ctk.CTkLabel(
            padre,
            text="Configuración",
            font=ctk.CTkFont(size=22, weight="bold"),
        ).pack(anchor="w", pady=(0, 12))

    def _crear_divisor(self, padre):
        ctk.CTkFrame(
            padre,
            height=2,
            fg_color=("gray80", "gray30"),
        ).pack(fill="x", pady=(0, 24))

    # ==================================================================
    #  FILA DE RUTA (título + botón + borrar), reutilizada para ambos casos
    # ==================================================================

    def _crear_fila_ruta(self, padre, titulo, clave_config):
        bloque = ctk.CTkFrame(padre, fg_color="transparent")
        bloque.pack(fill="x", pady=(0, 24))

        ctk.CTkLabel(
            bloque,
            text=titulo,
            font=ctk.CTkFont(size=13),
            text_color="gray",
        ).pack(anchor="w", pady=(0, 8))

        fila_boton = ctk.CTkFrame(bloque, fg_color="transparent")
        fila_boton.pack(anchor="w")

        boton = ctk.CTkButton(
            fila_boton,
            text="📁  Elegir carpeta",
            font=ctk.CTkFont(size=14, weight="bold"),
            width=320,
            height=42,
            corner_radius=10,
            fg_color="gray30",
            hover_color="gray20",
            command=lambda: self._elegir_ruta(boton, clave_config),
        )
        boton.pack(side="left")

        boton_borrar = ctk.CTkButton(
            fila_boton,
            text="✕",
            font=ctk.CTkFont(size=14, weight="bold"),
            width=42,
            height=42,
            corner_radius=10,
            fg_color="gray25",
            hover_color="#C0392B",
            command=lambda: self._borrar_ruta(boton, clave_config),
        )
        boton_borrar.pack(side="left", padx=(8, 0))

        ruta_guardada = obtener_valor(clave_config)
        if ruta_guardada and os.path.isdir(ruta_guardada):
            self._marcar_boton_con_ruta(boton, ruta_guardada)

        return boton

    # ==================================================================
    #  ACCIONES — RUTAS
    # ==================================================================

    def _elegir_ruta(self, boton, clave_config):
        ruta = filedialog.askdirectory(title="Seleccionar carpeta de exportación")

        if not ruta:
            return

        guardar_valor(clave_config, ruta)
        self._marcar_boton_con_ruta(boton, ruta)

    def _marcar_boton_con_ruta(self, boton, ruta):
        boton.configure(
            text=f"✓  {ruta}",
            fg_color="#2FA572",
            hover_color="#268A5F",
        )

    def _borrar_ruta(self, boton, clave_config):
        eliminar_valor(clave_config)
        boton.configure(
            text="📁  Elegir carpeta",
            fg_color="gray30",
            hover_color="gray20",
        )

    # ==================================================================
    #  API KEYS (genérico: sirve para Gemini y Borrador de Fondos)
    # ==================================================================

    def _crear_fila_api_key(self, padre, titulo, clave_config, icono_nombre, texto_boton, texto_dialogo):
        bloque = ctk.CTkFrame(padre, fg_color="transparent")
        bloque.pack(fill="x", pady=(0, 24))

        ctk.CTkLabel(
            bloque,
            text=titulo,
            font=ctk.CTkFont(size=13),
            text_color="gray",
        ).pack(anchor="w", pady=(0, 8))

        fila_boton = ctk.CTkFrame(bloque, fg_color="transparent")
        fila_boton.pack(anchor="w")

        icono = _cargar_logo(icono_nombre, size=(20, 20))

        boton = ctk.CTkButton(
            fila_boton,
            text=f"  {texto_boton}",
            image=icono,
            compound="left",
            font=ctk.CTkFont(size=14, weight="bold"),
            width=320,
            height=42,
            corner_radius=10,
            fg_color="gray30",
            hover_color="gray20",
            anchor="w",
            command=lambda: self._abrir_dialogo_api_key(boton, clave_config, texto_dialogo, texto_boton),
        )
        boton.pack(side="left")

        boton_borrar = ctk.CTkButton(
            fila_boton,
            text="✕",
            font=ctk.CTkFont(size=14, weight="bold"),
            width=42,
            height=42,
            corner_radius=10,
            fg_color="gray25",
            hover_color="#C0392B",
            command=lambda: self._borrar_api_key(boton, clave_config, texto_boton, icono),
        )
        boton_borrar.pack(side="left", padx=(8, 0))

        clave_guardada = obtener_valor(clave_config)
        if clave_guardada:
            self._marcar_boton_con_api_key(boton, clave_guardada, icono)

        return boton

    def _abrir_dialogo_api_key(self, boton, clave_config, texto_dialogo, texto_boton):
        dialogo = ctk.CTkInputDialog(
            text=texto_dialogo,
            title="API Key",
        )
        clave_ingresada = dialogo.get_input()

        if not clave_ingresada:
            return

        clave_ingresada = clave_ingresada.strip()
        guardar_valor(clave_config, clave_ingresada)

        icono_actual = boton.cget("image")
        self._marcar_boton_con_api_key(boton, clave_ingresada, icono_actual)

    def _marcar_boton_con_api_key(self, boton, clave, icono):
        clave_oculta = self._ocultar_clave(clave)
        boton.configure(
            text=f"  ✓  {clave_oculta}",
            image=icono,
            fg_color="#2FA572",
            hover_color="#268A5F",
        )

    def _borrar_api_key(self, boton, clave_config, texto_boton, icono):
        eliminar_valor(clave_config)
        boton.configure(
            text=f"  {texto_boton}",
            image=icono,
            fg_color="gray30",
            hover_color="gray20",
        )

    def _ocultar_clave(self, clave):
        """Muestra solo el principio y el final de la clave, ocultando el resto con puntos."""
        if not clave or len(clave) <= 10:
            return "•" * len(clave or "")
        return f"{clave[:6]}{'•' * 8}{clave[-4:]}"