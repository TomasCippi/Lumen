"""
informacion.py

Vista de Información: pantalla "Acerca de" con el nombre de la app,
versión, una descripción breve y el link al autor.
"""

import webbrowser

import customtkinter as ctk

COLOR_TARJETA = ("gray90", "#242424")
COLOR_BORDE = ("gray80", "#333333")


class InformacionView(ctk.CTkFrame):
    def __init__(self, master):
        super().__init__(master, fg_color="transparent")

        # Contenedor central (estilo tarjeta moderna, tamaño fijo)
        tarjeta = ctk.CTkFrame(
            self,
            fg_color=COLOR_TARJETA,
            border_width=1,
            border_color=COLOR_BORDE,
            corner_radius=16,
            width=420,
            height=340,
        )
        tarjeta.pack(expand=True, pady=40, padx=20)
        tarjeta.pack_propagate(False)  # mantiene el tamaño fijo y simétrico de la tarjeta

        # Nombre de la app
        lbl_titulo = ctk.CTkLabel(
            tarjeta,
            text="LUMEN",
            font=ctk.CTkFont(family="Segoe UI", size=32, weight="bold"),
            text_color=("#3B8ED0", "#569cd6"),  # azul destacado adaptativo (claro/oscuro)
        )
        lbl_titulo.pack(pady=(35, 2))

        # Versión
        lbl_version = ctk.CTkLabel(
            tarjeta,
            text="Versión 1.0.0",
            font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
            text_color="gray50",
        )
        lbl_version.pack(pady=(0, 20))

        # Descripción de la app
        descripcion_texto = (
            "Herramienta integral para el procesamiento, análisis y\n"
            "conversión inteligente de listas de precios de proveedores.\n"
            "Diseñada para optimizar los flujos de actualización de catálogos,\n"
            "garantizando consistencia absoluta con el ecosistema Stock Fácil."
        )
        lbl_descripcion = ctk.CTkLabel(
            tarjeta,
            text=descripcion_texto,
            font=ctk.CTkFont(family="Segoe UI", size=13),
            text_color=("gray40", "gray70"),
            justify="center",
        )
        lbl_descripcion.pack(pady=(0, 25), padx=24)

        # Línea divisoria sutil
        divisor = ctk.CTkFrame(tarjeta, height=1, fg_color=COLOR_BORDE)
        divisor.pack(fill="x", padx=45, pady=(0, 20))

        # Créditos / autoría
        lbl_creditos = ctk.CTkLabel(
            tarjeta,
            text="Desarrollado por",
            font=ctk.CTkFont(family="Segoe UI", size=11),
            text_color="gray50",
        )
        lbl_creditos.pack()

        # Botón enlace a GitHub del autor
        btn_link = ctk.CTkButton(
            tarjeta,
            text="TC",
            font=ctk.CTkFont(family="Segoe UI", size=14, weight="bold"),
            width=85,
            height=32,
            corner_radius=8,
            fg_color="#3B8ED0",
            hover_color="#2A6FA8",
            command=self._abrir_github,
        )
        btn_link.pack(pady=(6, 25))

    def _abrir_github(self):
        """Abre el repositorio o perfil del desarrollador en el navegador por defecto."""
        webbrowser.open("https://github.com/TomasCippi")