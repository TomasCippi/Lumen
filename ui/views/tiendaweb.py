import os
from tkinter import filedialog

import customtkinter as ctk

from functions.prompt_config import obtener_prompt_catalogo

COLOR_TITULO_SECCION = "gray"
COLOR_TARJETA = ("gray90", "#242424")
COLOR_BORDE = ("gray80", "#333333")


class TiendawebView(ctk.CTkFrame):
    def __init__(self, master):
        super().__init__(master, fg_color="transparent")

        self._id_reset_boton = None
        self.ruta_archivo = None

        exterior = ctk.CTkScrollableFrame(self, fg_color="transparent")
        exterior.pack(expand=True, fill="both")

        interior = ctk.CTkFrame(exterior, fg_color="transparent", width=520, height=700)
        interior.pack(pady=10)
        interior.pack_propagate(False)

        self._crear_titulo(interior)
        self._crear_divisor(interior)
        self._crear_seccion_prompt(interior)
        self._crear_seccion_archivo(interior)
        self._crear_log(interior)
        self._crear_boton_subir(interior)

    # ==================================================================
    #  TÍTULO Y DIVISOR
    # ==================================================================

    def _crear_titulo(self, padre):
        ctk.CTkLabel(
            padre,
            text="Tienda Web",
            font=ctk.CTkFont(size=22, weight="bold"),
        ).pack(anchor="w", pady=(0, 12))

    def _crear_divisor(self, padre):
        ctk.CTkFrame(
            padre,
            height=2,
            fg_color=("gray80", "gray30"),
        ).pack(fill="x", pady=(0, 24))

    # ==================================================================
    #  PROMPT PARA LA IA
    # ==================================================================

    def _crear_seccion_prompt(self, padre):
        ctk.CTkLabel(
            padre,
            text="Prompt para la IA",
            font=ctk.CTkFont(size=13),
            text_color=COLOR_TITULO_SECCION,
        ).pack(anchor="w", pady=(0, 8))

        self.boton_copiar = ctk.CTkButton(
            padre,
            text="📋  Copiar prompt",
            font=ctk.CTkFont(size=14, weight="bold"),
            height=42,
            corner_radius=10,
            fg_color="#3B8ED0",
            hover_color="#2A6FA8",
            command=self._copiar_prompt,
        )
        self.boton_copiar.pack(fill="x", pady=(0, 24))

    def _copiar_prompt(self):
        prompt = obtener_prompt_catalogo()

        self.clipboard_clear()
        self.clipboard_append(prompt)

        self.boton_copiar.configure(
            text="✓  ¡Se copió con éxito!",
            fg_color="#2FA572",
            hover_color="#268A5F",
        )

        if self._id_reset_boton is not None:
            self.after_cancel(self._id_reset_boton)

        self._id_reset_boton = self.after(2000, self._resetear_boton)

    def _resetear_boton(self):
        self.boton_copiar.configure(
            text="📋  Copiar prompt",
            fg_color="#3B8ED0",
            hover_color="#2A6FA8",
        )
        self._id_reset_boton = None

    # ==================================================================
    #  SUBIR A LA TIENDA WEB — SELECCIÓN DE ARCHIVO
    # ==================================================================

    def _crear_seccion_archivo(self, padre):
        ctk.CTkLabel(
            padre,
            text="Subir a la tienda web",
            font=ctk.CTkFont(size=13),
            text_color=COLOR_TITULO_SECCION,
        ).pack(anchor="w", pady=(0, 8))

        self.boton_archivo = ctk.CTkButton(
            padre,
            text="📂  Seleccionar Excel del catálogo",
            font=ctk.CTkFont(size=14, weight="bold"),
            height=42,
            corner_radius=10,
            fg_color="gray30",
            hover_color="gray20",
            command=self.seleccionar_archivo,
        )
        self.boton_archivo.pack(fill="x", pady=(0, 16))

    def seleccionar_archivo(self):
        ruta = filedialog.askopenfilename(
            title="Seleccionar Excel del catálogo",
            filetypes=[("Archivos de Excel", "*.xlsx *.xls")],
        )
        if not ruta:
            return

        self.ruta_archivo = ruta
        nombre_archivo = os.path.basename(ruta)

        self.boton_archivo.configure(
            text=f"✓  {nombre_archivo}",
            fg_color="#2FA572",
            hover_color="#268A5F",
        )

        self._log(f"Archivo seleccionado: {nombre_archivo}")

    # ==================================================================
    #  LOG
    # ==================================================================

    def _crear_log(self, padre):
        ctk.CTkLabel(
            padre,
            text="Registro del proceso",
            font=ctk.CTkFont(size=13),
            text_color=COLOR_TITULO_SECCION,
        ).pack(anchor="w", pady=(0, 8))

        self.caja_log = ctk.CTkTextbox(
            padre,
            height=180,
            corner_radius=10,
            fg_color=COLOR_TARJETA,
            border_width=1,
            border_color=COLOR_BORDE,
            font=ctk.CTkFont(family="Consolas", size=12),
        )
        self.caja_log.pack(fill="x", pady=(0, 16))
        self.caja_log.configure(state="disabled")

    def _log(self, texto):
        self.caja_log.configure(state="normal")
        self.caja_log.insert("end", f"{texto}\n")
        self.caja_log.see("end")
        self.caja_log.configure(state="disabled")

    # ==================================================================
    #  BOTÓN SUBIR (sin funcionalidad todavía)
    # ==================================================================

    def _crear_boton_subir(self, padre):
        self.boton_subir = ctk.CTkButton(
            padre,
            text="⬆  Subir a la tienda",
            font=ctk.CTkFont(size=16, weight="bold"),
            height=48,
            corner_radius=12,
            fg_color="#3B8ED0",
            hover_color="#2A6FA8",
            command=self._subir_a_la_tienda,
        )
        self.boton_subir.pack(fill="x")

    def _subir_a_la_tienda(self):
        # TODO: acá va la lógica real de subida al servidor, más adelante.
        self._log("⚠️ La subida a la tienda todavía no está implementada.")