"""
configuracion.py

Vista de Configuración: permite elegir y guardar las carpetas de
destino donde se exportan los Excel de "Stock Fácil" y de "Vendedor".
Las rutas quedan persistidas con config_manager, así que se recuerdan
entre sesiones.
"""

import customtkinter as ctk
from tkinter import filedialog

from functions.config_manager import guardar_valor, obtener_valor

# Claves usadas para guardar cada ruta en el archivo de configuración
CLAVE_RUTA_STOCK_FACIL = "ruta_exportacion_stock_facil"
CLAVE_RUTA_VENDEDOR = "ruta_exportacion_vendedor"


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
    #  FILA DE RUTA (título + botón), reutilizada para ambos casos
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

        boton = ctk.CTkButton(
            bloque,
            text="📁  Elegir carpeta",
            font=ctk.CTkFont(size=14, weight="bold"),
            width=320,
            height=42,
            corner_radius=10,
            fg_color="gray30",
            hover_color="gray20",
            command=lambda: self._elegir_ruta(boton, clave_config),
        )
        boton.pack(anchor="w")

        # Si ya había una ruta guardada de una sesión anterior, se muestra ya
        ruta_guardada = obtener_valor(clave_config)
        if ruta_guardada:
            self._marcar_boton_con_ruta(boton, ruta_guardada)

        return boton

    # ==================================================================
    #  ACCIONES
    # ==================================================================

    def _elegir_ruta(self, boton, clave_config):
        ruta = filedialog.askdirectory(title="Seleccionar carpeta de exportación")

        if not ruta:
            return  # el usuario cerró el diálogo sin elegir nada

        guardar_valor(clave_config, ruta)
        self._marcar_boton_con_ruta(boton, ruta)

    def _marcar_boton_con_ruta(self, boton, ruta):
        boton.configure(
            text=f"✓  {ruta}",
            fg_color="#2FA572",
            hover_color="#268A5F",
        )