import os
from tkinter import filedialog

import customtkinter as ctk

from functions.config_manager import guardar_valor, obtener_valor, eliminar_valor
from functions.prompt_config import CLAVE_PROMPT_CATALOGO, PROMPT_CATALOGO_POR_DEFECTO

CLAVE_RUTA_STOCK_FACIL = "ruta_exportacion_stock_facil"
CLAVE_RUTA_VENDEDOR = "ruta_exportacion_vendedor"

class ConfiguracionView(ctk.CTkFrame):
    def __init__(self, master):
        super().__init__(master, fg_color="transparent")

        exterior = ctk.CTkScrollableFrame(self, fg_color="transparent")
        exterior.pack(expand=True, fill="both")

        contenido = ctk.CTkFrame(exterior, fg_color="transparent", width=520, height=750)
        contenido.pack(pady=20)
        contenido.pack_propagate(False)

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
        self._crear_seccion_prompt(contenido)

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
    #  PROMPT DE TIENDA WEB (editable, persistente)
    # ==================================================================

    def _crear_seccion_prompt(self, padre):
        bloque = ctk.CTkFrame(padre, fg_color="transparent")
        bloque.pack(fill="x", pady=(0, 24))

        ctk.CTkLabel(
            bloque,
            text="Prompt para clasificar el catálogo (Tienda Web)",
            font=ctk.CTkFont(size=13),
            text_color="gray",
        ).pack(anchor="w", pady=(0, 8))

        self.caja_prompt = ctk.CTkTextbox(
            bloque,
            height=280,
            corner_radius=8,
            font=ctk.CTkFont(family="Consolas", size=12),
            wrap="word",
        )
        self.caja_prompt.pack(fill="x", pady=(0, 10))

        prompt_guardado = obtener_valor(CLAVE_PROMPT_CATALOGO, PROMPT_CATALOGO_POR_DEFECTO)
        self.caja_prompt.insert("1.0", prompt_guardado)

        fila_botones = ctk.CTkFrame(bloque, fg_color="transparent")
        fila_botones.pack(fill="x")

        self.boton_guardar_prompt = ctk.CTkButton(
            fila_botones,
            text="💾  Guardar prompt",
            font=ctk.CTkFont(size=14, weight="bold"),
            height=42,
            corner_radius=10,
            fg_color="#3B8ED0",
            hover_color="#2A6FA8",
            command=self._guardar_prompt,
        )
        self.boton_guardar_prompt.pack(side="left", fill="x", expand=True, padx=(0, 8))

        boton_restaurar = ctk.CTkButton(
            fila_botones,
            text="↺  Restaurar original",
            font=ctk.CTkFont(size=14, weight="bold"),
            height=42,
            corner_radius=10,
            fg_color="gray30",
            hover_color="gray20",
            command=self._restaurar_prompt_por_defecto,
        )
        boton_restaurar.pack(side="left")

    def _guardar_prompt(self):
        texto = self.caja_prompt.get("1.0", "end-1c")
        guardar_valor(CLAVE_PROMPT_CATALOGO, texto)

        self.boton_guardar_prompt.configure(
            text="✓  ¡Guardado con éxito!",
            fg_color="#2FA572",
            hover_color="#268A5F",
        )
        self.after(2000, self._resetear_boton_guardar)

    def _resetear_boton_guardar(self):
        self.boton_guardar_prompt.configure(
            text="💾  Guardar prompt",
            fg_color="#3B8ED0",
            hover_color="#2A6FA8",
        )

    def _restaurar_prompt_por_defecto(self):
        self.caja_prompt.delete("1.0", "end")
        self.caja_prompt.insert("1.0", PROMPT_CATALOGO_POR_DEFECTO)