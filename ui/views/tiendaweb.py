"""
tiendaweb.py

Vista para clasificar el catálogo de productos con IA: se elige el
Excel del proveedor, se convierte con Gemini (agrupando variantes y
asignando categoría), y al terminar se guarda el resultado como JSON
donde el usuario elija.
"""

import os
import threading
from tkinter import filedialog, messagebox

import customtkinter as ctk

from functions.lector_excel import detectar_columnas
from functions.preparar_excel import preparar_excel
from functions.ia_convertidor import clasificar_excel

COLOR_TARJETA = ("gray90", "#242424")
COLOR_BORDE = ("gray80", "#333333")


class TiendawebView(ctk.CTkFrame):
    def __init__(self, master):
        super().__init__(master, fg_color="transparent")

        self.ruta_archivo = None
        self.resultado_clasificacion = None

        exterior = ctk.CTkScrollableFrame(self, fg_color="transparent")
        exterior.pack(expand=True, fill="both")

        interior = ctk.CTkFrame(exterior, fg_color="transparent", width=500)
        interior.pack(pady=10, padx=10, fill="x")

        self._crear_paso_1(interior)
        self._crear_boton_convertir(interior)
        self._crear_log(interior)

    # ==================================================================
    #  PASO 1 — Selección de archivo
    # ==================================================================

    def _crear_paso_1(self, padre):
        ctk.CTkLabel(
            padre,
            text="Paso 1 — Seleccioná el Excel del proveedor",
            font=ctk.CTkFont(size=14, weight="bold"),
        ).pack(anchor="w", pady=(0, 6))

        ctk.CTkFrame(
            padre, height=2, fg_color=("gray80", "gray30"),
        ).pack(fill="x", pady=(0, 16))

        self.boton_archivo = ctk.CTkButton(
            padre,
            text="📂  Seleccionar archivo Excel",
            font=ctk.CTkFont(size=14, weight="bold"),
            height=42,
            corner_radius=10,
            fg_color="gray30",
            hover_color="gray20",
            command=self.seleccionar_archivo,
        )
        self.boton_archivo.pack(fill="x", pady=(0, 16))

    # ==================================================================
    #  BOTÓN CONVERTIR
    # ==================================================================

    def _crear_boton_convertir(self, padre):
        self.boton_convertir = ctk.CTkButton(
            padre,
            text="Convertir",
            font=ctk.CTkFont(size=16, weight="bold"),
            height=48,
            corner_radius=12,
            fg_color="#3B8ED0",
            hover_color="#2A6FA8",
            command=self.convertir,
        )
        self.boton_convertir.pack(fill="x", pady=(0, 16))

    # ==================================================================
    #  LOG
    # ==================================================================

    def _crear_log(self, padre):
        ctk.CTkLabel(
            padre,
            text="Registro del proceso",
            font=ctk.CTkFont(size=13),
            text_color="gray",
        ).pack(anchor="w", pady=(0, 6))

        self.caja_log = ctk.CTkTextbox(
            padre,
            height=260,
            corner_radius=10,
            fg_color=COLOR_TARJETA,
            border_width=1,
            border_color=COLOR_BORDE,
            font=ctk.CTkFont(family="Consolas", size=12),
        )
        self.caja_log.pack(fill="both", expand=True, pady=(0, 10))
        self.caja_log.configure(state="disabled")

    def _log(self, texto):
        """Agrega una línea al registro de forma segura desde cualquier hilo."""
        self.after(0, self._escribir_log, texto)

    def _escribir_log(self, texto):
        self.caja_log.configure(state="normal")
        self.caja_log.insert("end", f"{texto}\n")
        self.caja_log.see("end")
        self.caja_log.configure(state="disabled")

    # ==================================================================
    #  SELECCIÓN DE ARCHIVO
    # ==================================================================

    def seleccionar_archivo(self):
        ruta = filedialog.askopenfilename(
            title="Seleccionar archivo Excel",
            filetypes=[("Archivos de Excel", "*.xlsx *.xls")],
        )
        if not ruta:
            return

        nombre_archivo = os.path.basename(ruta)
        self.boton_archivo.configure(text="⏳  Preparando archivo...", state="disabled")
        self.update_idletasks()

        try:
            ruta = preparar_excel(ruta)
        except Exception as e:
            messagebox.showerror(
                "No se pudo preparar el archivo",
                f"Ocurrió un problema al leer/recalcular el Excel:\n\n{e}",
            )
            self.boton_archivo.configure(text="📂  Seleccionar archivo Excel", state="normal")
            return

        self.ruta_archivo = ruta
        self.boton_archivo.configure(
            text=f"✓  {nombre_archivo}", fg_color="#2FA572", hover_color="#268A5F",
            state="normal",
        )
        self._log(f"Archivo seleccionado: {nombre_archivo}")

    # ==================================================================
    #  CONVERSIÓN
    # ==================================================================

    def convertir(self):
        if not self.ruta_archivo:
            self._log("⚠️ Primero seleccioná un archivo Excel.")
            return

        columnas = detectar_columnas(self.ruta_archivo)

        if not columnas["codigo"] or not columnas["producto"] or not columnas["precio"]:
            self._log("❌ No se pudieron detectar automáticamente las columnas de Código, Producto o Precio en este Excel.")
            return

        self.boton_convertir.configure(state="disabled", text="Procesando...")
        self._log("🚀 Iniciando clasificación con IA...")

        hilo = threading.Thread(target=self._ejecutar_clasificacion, args=(columnas,))
        hilo.start()

    def _ejecutar_clasificacion(self, columnas):
        try:
            def callback_progreso(lote_actual, total_lotes):
                self._log(f"Procesando lote {lote_actual} de {total_lotes}...")

            resultado = clasificar_excel(
                ruta_archivo=self.ruta_archivo,
                columna_codigo=columnas["codigo"],
                columna_producto=columnas["producto"],
                columna_precio=columnas["precio"],
                columna_familia=columnas["familia"],
                callback_progreso=callback_progreso,
            )

            self.resultado_clasificacion = resultado
            cantidad_productos = len(resultado["productos"])
            self._log(f"✅ Clasificación completa: {cantidad_productos} productos procesados.")

            if resultado["errores_lectura_excel"]:
                self._log(f"⚠️ {len(resultado['errores_lectura_excel'])} filas del Excel no se pudieron leer (código o producto faltante).")

            self.after(0, self._guardar_json)

        except Exception as e:
            self._log(f"❌ Error durante la clasificación: {e}")

        finally:
            self.after(0, self._finalizar_conversion)

    # ==================================================================
    #  GUARDADO DEL JSON
    # ==================================================================

    def _guardar_json(self):
        ruta_guardado = filedialog.asksaveasfilename(
            title="Guardar resultado como JSON",
            defaultextension=".json",
            filetypes=[("Archivo JSON", "*.json")],
        )

        if not ruta_guardado:
            self._log("⚠️ Guardado cancelado por el usuario.")
            return

        try:
            import json
            with open(ruta_guardado, "w", encoding="utf-8") as f:
                json.dump(self.resultado_clasificacion, f, indent=2, ensure_ascii=False)
            self._log(f"💾 Guardado correctamente en: {ruta_guardado}")
        except Exception as e:
            self._log(f"❌ No se pudo guardar el archivo: {e}")

    def _finalizar_conversion(self):
        self.boton_convertir.configure(state="normal", text="Convertir")