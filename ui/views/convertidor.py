"""
convertidor.py

Vista principal de la app: el formulario donde el usuario elige el
Excel del proveedor, indica qué columna es cada cosa (Código, Producto,
Precio, Familia), configura descuento/aumento/% vendedor/dólar, y
dispara la conversión.

La conversión en sí (leer el Excel, generar los distintos formatos de
salida y validarlos) se corre en un hilo aparte para no congelar la
interfaz, y el progreso se muestra en un pop-up con barra de progreso.
"""

import os
import threading
from tkinter import filedialog, messagebox, ttk

import customtkinter as ctk

from functions.config_manager import obtener_valor
from functions.exportador import exportar_stock_facil, exportar_vendedor
from functions.lector_excel import detectar_columnas
from functions.preparar_excel import preparar_excel
from functions.procesador import procesar_excel
from functions.calculadora_precios import calcular_precio_total, parsear_porcentajes_encadenados

COLOR_TARJETA = ("gray90", "#242424")
COLOR_BORDE = ("gray80", "#333333")
COLOR_TITULO_SECCION = "gray"


class ConvertidorView(ctk.CTkFrame):
    def __init__(self, master):
        super().__init__(master, fg_color="transparent")

        self.ruta_archivo = None  # se completa al elegir el Excel

        self._estilizar_combobox()

        self._validar_numero_cmd = (self.register(self._validar_numero), "%P")
        self._validar_porcentaje_encadenado_cmd = (
            self.register(self._validar_porcentaje_encadenado), "%P"
        )

        exterior = ctk.CTkScrollableFrame(self, fg_color="transparent")
        exterior.pack(expand=True, fill="both")

        interior = ctk.CTkFrame(exterior, fg_color="transparent", width=420)
        interior.pack(pady=10)

        scroll = interior

        self._crear_tarjeta_archivo(scroll)
        self._crear_tarjeta_columnas(scroll)
        self._crear_tarjeta_datos_origen(scroll)
        self._crear_tarjeta_precios(scroll)
        self._crear_boton_convertir(scroll)

    # ==================================================================
    #  TARJETA 1 — Selección de archivo
    # ==================================================================

    def _crear_tarjeta_archivo(self, padre):
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
    #  TARJETA 2 — Columnas del Excel (Código, Producto, Precio)
    # ==================================================================

    def _crear_tarjeta_columnas(self, padre):
        tarjeta = ctk.CTkFrame(
            padre, fg_color=COLOR_TARJETA, corner_radius=14,
            border_width=1, border_color=COLOR_BORDE,
        )
        tarjeta.pack(fill="x", pady=(0, 16), ipadx=14, ipady=12)

        letras_excel = [chr(i) for i in range(ord("A"), ord("Z") + 1)]

        self.dropdown_codigo = self._crear_dropdown_columna(
            tarjeta, "Código", letras_excel, columna=0,
        )
        self.dropdown_producto = self._crear_dropdown_columna(
            tarjeta, "Producto", letras_excel, columna=1,
        )
        self.dropdown_precio = self._crear_dropdown_columna(
            tarjeta, "Precio", letras_excel, columna=2,
        )

        tarjeta.grid_columnconfigure(0, weight=1)
        tarjeta.grid_columnconfigure(1, weight=1)
        tarjeta.grid_columnconfigure(2, weight=1)

    def _crear_dropdown_columna(self, padre, titulo, letras, columna):
        """Label arriba + dropdown abajo, ambos centrados. Devuelve el dropdown."""
        ctk.CTkLabel(
            padre, text=titulo, font=ctk.CTkFont(size=12), text_color="gray",
        ).grid(row=0, column=columna, pady=(0, 4))

        dropdown = ttk.Combobox(
            padre,
            values=letras,
            width=4,
            state="readonly",
            style="Lumen.TCombobox",
            font=("Segoe UI", 13),
            justify="center",
        )
        dropdown.grid(row=1, column=columna, ipady=4)
        dropdown.bind("<<ComboboxSelected>>", lambda e: dropdown.selection_clear())

        return dropdown

    # ==================================================================
    #  TARJETA 3 — Proveedor + columna de Familia (sin título de tarjeta)
    # ==================================================================

    def _crear_tarjeta_datos_origen(self, padre):
        tarjeta = ctk.CTkFrame(padre, fg_color="transparent")
        tarjeta.pack(fill="x", pady=(0, 16))

        # -- Proveedor --
        self._label_campo(tarjeta, "Proveedor", fila=0)
        self.input_proveedor = ctk.CTkEntry(
            tarjeta,
            placeholder_text="Ej: Distribuidora XYZ",
            height=36,
            corner_radius=8,
        )
        self.input_proveedor.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(0, 12))

        # -- Columna de Familia (checkbox + dropdown) --
        self.check_familia = ctk.CTkCheckBox(
            tarjeta, text="Usar columna de Familia", command=self.toggle_familia,
        )
        self.check_familia.grid(row=2, column=0, sticky="w", pady=(0, 4))

        letras_excel = [chr(i) for i in range(ord("A"), ord("Z") + 1)]
        self.dropdown_familia = ttk.Combobox(
            tarjeta,
            values=letras_excel,
            width=4,
            state="disabled",
            style="Lumen.TCombobox",
            font=("Segoe UI", 13),
            justify="center",
        )
        self.dropdown_familia.grid(row=2, column=1, sticky="e", pady=(0, 4), ipady=4)
        self.dropdown_familia.bind(
            "<<ComboboxSelected>>", lambda e: self.dropdown_familia.selection_clear(),
        )

        tarjeta.grid_columnconfigure(0, weight=1)
        tarjeta.grid_columnconfigure(1, weight=0)

    # ==================================================================
    #  TARJETA 4 — Precios (dólar, descuento, aumento, vendedor)
    # ==================================================================

    def _crear_tarjeta_precios(self, padre):
        tarjeta = self._nueva_tarjeta(padre, "Precios")

        # -- Precio en dólares: checkbox a la izquierda, input a la derecha --
        self.check_dolar = ctk.CTkCheckBox(
            tarjeta, text="Precio en dólares", command=self.toggle_dolar,
        )
        self.check_dolar.grid(row=1, column=0, sticky="w", pady=(0, 12))

        self.input_dolar = self._input_con_prefijo(
            tarjeta, prefijo="$", valor_inicial="1300", fila=1, columna=1, deshabilitado=True,
        )

        # -- Descuento y Aumento, uno al lado del otro --
        self._label_campo(tarjeta, "Descuento", fila=2, columna=0)
        self._label_campo(tarjeta, "Aumento", fila=2, columna=1)

        self.input_descuento = self._input_con_sufijo(
            tarjeta, sufijo="%", fila=3, columna=0, encadenado=True,
        )
        self.input_aumento = self._input_con_sufijo(
            tarjeta, sufijo="%", fila=3, columna=1, encadenado=True,
        )

        # -- % Vendedor --
        self._label_campo(tarjeta, "Porcentaje del vendedor", fila=4, columna=0)
        self.input_vendedor = self._input_con_sufijo(
            tarjeta, sufijo="%", fila=5, columna=0, valor_inicial="181.5",
        )

        tarjeta.grid_columnconfigure(0, weight=1)
        tarjeta.grid_columnconfigure(1, weight=1)

    # ==================================================================
    #  BOTÓN FINAL
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
        self.boton_convertir.pack(fill="x", pady=(4, 10))

    # ==================================================================
    #  HELPERS DE UI — para no repetir código entre tarjetas
    # ==================================================================

    def _nueva_tarjeta(self, padre, titulo):
        tarjeta = ctk.CTkFrame(
            padre, fg_color=COLOR_TARJETA, corner_radius=14,
            border_width=1, border_color=COLOR_BORDE,
        )
        tarjeta.pack(fill="x", pady=(0, 16), ipadx=14, ipady=12)

        ctk.CTkLabel(
            tarjeta, text=titulo, font=ctk.CTkFont(size=13, weight="bold"),
            text_color=COLOR_TITULO_SECCION,
        ).grid(row=0, column=0, columnspan=3, sticky="w", pady=(0, 10))

        return tarjeta

    def _label_campo(self, padre, texto, fila, columna=0):
        ctk.CTkLabel(
            padre, text=texto, font=ctk.CTkFont(size=12), text_color="gray",
        ).grid(row=fila, column=columna, sticky="w", padx=(0, 8), pady=(0, 4))

    def _input_con_prefijo(self, padre, prefijo, fila, columna, valor_inicial="", deshabilitado=False):
        """Caja tipo '$ [___]', usada para el valor del dólar."""
        caja = ctk.CTkFrame(padre, height=34, corner_radius=8, fg_color=("gray85", "gray20"))
        caja.grid(row=fila, column=columna, sticky="ew", padx=(0, 8), pady=(0, 12))
        caja.grid_propagate(False)

        ctk.CTkLabel(
            caja, text=prefijo, font=ctk.CTkFont(size=14, weight="bold"), text_color="gray50",
        ).pack(side="left", padx=(8, 0))

        entry = ctk.CTkEntry(
            caja, corner_radius=0, border_width=0, fg_color="transparent",
            state="disabled" if deshabilitado else "normal",
            validate="key", validatecommand=self._validar_numero_cmd,
        )
        entry.pack(side="left", fill="both", expand=True, padx=(4, 8))
        if valor_inicial:
            entry.insert(0, valor_inicial)
        return entry

    def _input_con_sufijo(self, padre, sufijo, fila, columna, valor_inicial="", encadenado=False):
        """Caja tipo '[___] %', usada para descuento, aumento y % vendedor."""
        validador = (
            self._validar_porcentaje_encadenado_cmd if encadenado else self._validar_numero_cmd
        )

        caja = ctk.CTkFrame(padre, height=34, corner_radius=8, fg_color=("gray85", "gray20"))
        caja.grid(row=fila, column=columna, sticky="ew", padx=(0, 8), pady=(0, 12))
        caja.grid_propagate(False)

        entry = ctk.CTkEntry(
            caja, placeholder_text="0", corner_radius=0, border_width=0,
            fg_color="transparent", validate="key", validatecommand=validador,
        )
        entry.pack(side="left", fill="both", expand=True, padx=(8, 4))
        if valor_inicial:
            entry.insert(0, valor_inicial)

        ctk.CTkLabel(
            caja, text=sufijo, font=ctk.CTkFont(size=14, weight="bold"), text_color="gray50",
        ).pack(side="left", padx=(0, 8))
        return entry

    # ==================================================================
    #  ESTILO DEL COMBOBOX
    # ==================================================================

    def _estilizar_combobox(self):
        """ttk no respeta el tema de customtkinter por defecto, así que se
        pisa manualmente el estilo del Combobox para que combine con el resto."""
        estilo = ttk.Style()
        estilo.theme_use("clam")
        estilo.configure(
            "Lumen.TCombobox", fieldbackground="#2b2b2b", background="#2b2b2b",
            foreground="white", arrowcolor="white", bordercolor="#3d3d3d",
            lightcolor="#2b2b2b", darkcolor="#2b2b2b", padding=6, relief="flat",
        )
        estilo.map(
            "Lumen.TCombobox",
            fieldbackground=[("readonly", "#2b2b2b"), ("disabled", "#242424")],
            foreground=[("disabled", "gray50")], bordercolor=[("focus", "#565b5e")],
        )
        self.option_add("*TCombobox*Listbox.background", "#2b2b2b")
        self.option_add("*TCombobox*Listbox.foreground", "white")
        self.option_add("*TCombobox*Listbox.selectBackground", "#3B8ED0")
        self.option_add("*TCombobox*Listbox.selectForeground", "white")
        self.option_add("*TCombobox*Listbox.font", ("Segoe UI", 13))

    # ==================================================================
    #  SELECCIÓN DE ARCHIVO Y AUTOCOMPLETADO
    # ==================================================================

    def seleccionar_archivo(self):
        """Abre el diálogo de selección, prepara el Excel (convierte/recalcula) y autocompleta columnas."""
        ruta = filedialog.askopenfilename(
            title="Seleccionar archivo Excel",
            filetypes=[("Archivos de Excel", "*.xlsx *.xls")],
        )
        if not ruta:
            return

        nombre_archivo = os.path.basename(ruta)
        self.boton_archivo.configure(text="⏳  Preparando archivo...", state="disabled")
        self.update_idletasks()

        # Convierte .xls -> .xlsx si hace falta, y recalcula todas las
        # fórmulas abriendo el archivo con Excel real. Esto garantiza
        # que columnas con fórmulas (ej. BUSCARV) siempre den un valor
        # numérico y no la fórmula cruda ni un vacío.
        try:
            ruta = preparar_excel(ruta)
        except Exception as e:
            messagebox.showerror(
                "No se pudo preparar el archivo",
                f"Ocurrió un problema al leer/recalcular el Excel:\n\n{e}",
            )
            self.boton_archivo.configure(
                text="📂  Seleccionar archivo Excel", state="normal",
            )
            return

        self.ruta_archivo = ruta
        self.boton_archivo.configure(
            text=f"✓  {nombre_archivo}", fg_color="#2FA572", hover_color="#268A5F",
            state="normal",
        )
        self._autocompletar_columnas(ruta)

    def _autocompletar_columnas(self, ruta_archivo):
        """Usa detectar_columnas para precargar los dropdowns; el usuario puede corregirlos a mano."""
        columnas = detectar_columnas(ruta_archivo)
        if columnas["codigo"]:
            self.dropdown_codigo.set(columnas["codigo"])
        if columnas["producto"]:
            self.dropdown_producto.set(columnas["producto"])
        if columnas["precio"]:
            self.dropdown_precio.set(columnas["precio"])
        if columnas["familia"]:
            self.dropdown_familia.set(columnas["familia"])
            self.check_familia.select()
            self.toggle_familia()

    # ==================================================================
    #  TOGGLES Y VALIDACIONES DE INPUT
    # ==================================================================

    def toggle_familia(self):
        estado = "readonly" if self.check_familia.get() == 1 else "disabled"
        self.dropdown_familia.configure(state=estado)

    def toggle_dolar(self):
        estado = "normal" if self.check_dolar.get() == 1 else "disabled"
        self.input_dolar.configure(state=estado)

    def _validar_numero(self, texto_nuevo):
        """Valida mientras se tipea: solo dígitos y un único punto decimal."""
        if texto_nuevo == "":
            return True
        return texto_nuevo.replace(".", "", 1).isdigit()

    def _validar_porcentaje_encadenado(self, texto_nuevo):
        """Valida mientras se tipea: dígitos, puntos y guiones (para porcentajes encadenados tipo "29-3-3")."""
        if texto_nuevo == "":
            return True
        caracteres_validos = set("0123456789.-")
        return all(c in caracteres_validos for c in texto_nuevo)

    # ==================================================================
    #  CONVERSIÓN: DISPARO DEL HILO Y POP-UP DE PROGRESO
    # ==================================================================

    def convertir(self):
        """Punto de entrada principal. Valida campos en la UI y lanza el hilo de conversión."""
        if not self.ruta_archivo:
            print("Falta seleccionar un archivo Excel.")
            return

        columna_codigo = self.dropdown_codigo.get()
        columna_producto = self.dropdown_producto.get()
        columna_precio = self.dropdown_precio.get()
        columna_familia = self.dropdown_familia.get() if self.check_familia.get() == 1 else None

        if not (columna_codigo and columna_producto and columna_precio):
            print("Faltan elegir columnas de Código, Producto o Precio.")
            return

        # Capturar todos los valores de la UI antes de ir al hilo independiente
        nombre_proveedor = self.input_proveedor.get() or "Proveedor sin nombre"
        incluir_familia = self.check_familia.get() == 1
        porcentaje_vendedor = float(self.input_vendedor.get() or 0)
        texto_desc = self.input_descuento.get().strip()
        texto_aum = self.input_aumento.get().strip()
        esta_en_dolares = self.check_dolar.get() == 1

        valor_dolar = None
        if esta_en_dolares:
            try:
                valor_dolar = float(self.input_dolar.get())
            except ValueError:
                valor_dolar = None

        # Deshabilitar botón principal para evitar dobles clics
        self.boton_convertir.configure(state="disabled", text="Procesando...")

        # 1) Crear ventana pop-up de progreso
        popup = ctk.CTkToplevel(self)
        popup.title("Procesando Archivos")
        popup.geometry("380x160")
        popup.resizable(False, False)
        popup.transient(self.master)
        popup.grab_set()  # bloquea interacciones con la ventana de atrás

        # Centrar el pop-up relativo a la app principal
        popup.update_idletasks()
        x = self.winfo_screenwidth() // 2 - 190
        y = self.winfo_screenheight() // 2 - 80
        popup.geometry(f"+{x}+{y}")

        # Elementos del pop-up
        lbl_estado = ctk.CTkLabel(
            popup, text="⏳ Iniciando exportación...",
            font=ctk.CTkFont(size=14, weight="bold"),
        )
        lbl_estado.pack(pady=(24, 8))

        progress_bar = ctk.CTkProgressBar(popup, width=280)
        progress_bar.pack(pady=8)
        progress_bar.set(0.1)

        btn_cerrar = ctk.CTkButton(
            popup, text="Cerrar", width=100,
            command=popup.destroy, state="disabled",
        )
        btn_cerrar.pack(pady=(12, 0))

        # 2) Arrancar la tarea pesada en un hilo separado para no congelar la UI
        args = (
            columna_codigo, columna_producto, columna_precio, columna_familia,
            nombre_proveedor, incluir_familia, porcentaje_vendedor, texto_desc,
            texto_aum, esta_en_dolares, valor_dolar, lbl_estado, progress_bar, btn_cerrar,
        )
        hilo = threading.Thread(target=self._ejecutar_conversion_background, args=args)
        hilo.start()

    def _ejecutar_conversion_background(
        self, col_cod, col_prod, col_prec, col_fam, proveedor, inc_fam,
        pct_vendedor, t_desc, t_aum, en_dolar, v_dolar, lbl, pbar, btn,
    ):
        try:
            self.after(0, self._actualizar_progreso, lbl, pbar, "📦 Leyendo y procesando Excel original...", 0.25)

            resultado = procesar_excel(
                self.ruta_archivo, columna_codigo=col_cod,
                columna_producto=col_prod, columna_precio=col_prec, columna_familia=col_fam,
            )
            productos = resultado["productos"]

            if not productos:
                self.after(0, self._actualizar_progreso, lbl, pbar, "⚠️ No hay productos para exportar.", 1.0, "orange")
                self.after(0, self._finalizar_conversion, btn)
                return

            self.after(0, self._actualizar_progreso, lbl, pbar, "🗂️ Generando listas de Vendedor...", 0.50)

            carpeta_destino = obtener_valor("ruta_exportacion_vendedor")
            if not carpeta_destino:
                self.after(0, self._actualizar_progreso, lbl, pbar, "⚠️ Ruta de exportación vendedor no configurada.")
                self.after(0, self._finalizar_conversion, btn)
                return

            rutas_exportadas = exportar_vendedor(
                productos=productos, nombre_proveedor=proveedor,
                ruta_carpeta_destino=carpeta_destino, incluir_familia=bool(col_fam),
                porcentaje_vendedor=pct_vendedor, valor_dolar=v_dolar,
                texto_descuento=t_desc, texto_aumento=t_aum,
            )
            print(f"Excel generado en:\n{rutas_exportadas['individual']}")

            self.after(0, self._actualizar_progreso, lbl, pbar, "⚙️ Exportando formato Stock Fácil...", 0.75)

            descuentos = parsear_porcentajes_encadenados(t_desc)
            aumentos = parsear_porcentajes_encadenados(t_aum)

            precio_total_por_producto = {
                prod["codigo"]: calcular_precio_total(
                    prod["precio"], descuentos, aumentos, pct_vendedor, v_dolar if en_dolar else None,
                )
                for prod in productos
            }

            ruta_stock = obtener_valor("ruta_exportacion_stock_facil")
            if not ruta_stock:
                self.after(0, self._actualizar_progreso, lbl, pbar, "⚠️ Ruta de Stock Fácil no configurada.")
                self.after(0, self._finalizar_conversion, btn)
                return

            ruta_generada = exportar_stock_facil(
                productos=productos, nombre_proveedor=proveedor, ruta_carpeta_destino=ruta_stock,
                incluir_familia=inc_fam, precio_total_por_producto=precio_total_por_producto,
            )
            print(f"Exportado correctamente a Stock Fácil en:\n{ruta_generada}")

            self.after(0, self._actualizar_progreso, lbl, pbar, "🔍 Ejecutando validación de consistencia...", 0.90)

            try:
                from functions.validador_stock import validar_exportacion
                validar_exportacion(
                    ruta_original=self.ruta_archivo, ruta_stock_facil=ruta_generada,
                    col_codigo=col_cod, col_producto=col_prod, col_precio=col_prec,
                    texto_descuento=t_desc, texto_aumento=t_aum,
                    porcentaje_vendedor=pct_vendedor, valor_dolar=v_dolar, esta_en_dolares=en_dolar,
                )
            except Exception as error_validacion:
                print(f"⚠️ No se pudo ejecutar la validación: {error_validacion}")

            self.after(0, self._actualizar_progreso, lbl, pbar, "✨ ¡Excels exportados con éxito!", 1.0, "#2FA572")

        except Exception as e:
            self.after(0, self._actualizar_progreso, lbl, pbar, "❌ Ocurrió un error en el proceso.", 1.0, "red")
            print(f"Error crítico en hilo de conversión: {e}")

        finally:
            self.after(0, self._finalizar_conversion, btn)

    def _actualizar_progreso(self, lbl, pbar, texto=None, valor=None, color=None):
        """Actualiza los widgets del popup de forma segura desde el hilo principal."""
        if texto is not None:
            lbl.configure(text=texto)
        if color is not None:
            pbar.configure(progress_color=color)
        if valor is not None:
            pbar.set(valor)

    def _finalizar_conversion(self, btn):
        """Reactiva los botones al terminar (éxito, error o corte temprano)."""
        btn.configure(state="normal")
        self.boton_convertir.configure(state="normal", text="Convertir")