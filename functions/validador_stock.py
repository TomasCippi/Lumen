"""
validador_stock.py

Validación cruzada entre el Excel original del proveedor y el Excel
exportado en formato Stock Fácil.

La idea es reconstruir, a partir del original, la misma lista de
productos que arma `procesador.py` (mismo criterio de inicio/corte de
tabla), y después comparar cada producto contra la fila equivalente
del Stock Fácil generado:
    - que el texto (descripción) se haya limpiado igual,
    - que "precio1" sea el precio de origen,
    - que "precio2" sea el resultado correcto de aplicar
      descuentos -> aumentos -> % vendedor -> dólar (en ese orden),
      replicando exactamente las cuentas del exportador.

Sirve como chequeo de humo antes de entregarle el Stock Fácil al
usuario: si algo en la cadena de cálculo se rompió, esto lo detecta.
"""

import os
import re

import openpyxl
import pandas as pd

from functions.calculadora_precios import calcular_precio_total, parsear_porcentajes_encadenados

CANTIDAD_FILAS_A_REVISAR_ANTES_DE_CORTAR = 3
TOLERANCIA_COMPARACION_PRECIOS = 0.01
MAXIMO_ERRORES_A_MOSTRAR = 20


# ─────────────────────────────────────────────────────────────────────────
# Helpers de limpieza (mismo criterio que procesador.py / exportador.py)
# ─────────────────────────────────────────────────────────────────────────

def _limpiar_precio(valor):
    """Convierte el valor de una celda a número. Soporta texto con símbolos ($, comas, etc)."""
    if valor is None:
        return None
    if isinstance(valor, (int, float)):
        return float(valor)

    texto = str(valor).strip()
    if not texto:
        return None

    texto = re.sub(r"[^\d.,-]", "", texto)
    if not texto:
        return None

    if "," in texto and "." in texto:
        texto = texto.replace(".", "").replace(",", ".")
    elif "," in texto:
        texto = texto.replace(",", ".")

    try:
        return float(texto)
    except ValueError:
        return None


def limpiar_texto(valor):
    """Misma normalización de texto que usa Stock Fácil: comas -> puntos, comillas simples -> dobles."""
    if valor is None:
        return ""
    return str(valor).replace(",", ".").replace("'", '"').strip()


# ─────────────────────────────────────────────────────────────────────────
# Reconstrucción de la lista de productos esperados (a partir del original)
# ─────────────────────────────────────────────────────────────────────────

def _extraer_productos_esperados(todas_las_filas, idx_codigo, idx_producto, idx_precio):
    """
    Replica la lógica de `procesador.procesar_excel`: busca dónde arranca
    la tabla real (primera fila con precio válido) y la recorre hasta
    que se corta (3 filas seguidas sin precio). Devuelve la lista de
    productos esperados, ignorando las filas con código o producto vacío.
    """
    total_filas = len(todas_las_filas)

    def valor_en(fila_valores, indice):
        if indice is None or indice - 1 >= len(fila_valores):
            return None
        return fila_valores[indice - 1]

    def tiene_precio(numero_fila_1indexed):
        indice_lista = numero_fila_1indexed - 1
        if indice_lista < 0 or indice_lista >= total_filas:
            return False
        return _limpiar_precio(valor_en(todas_las_filas[indice_lista], idx_precio)) is not None

    # Buscar inicio real de la tabla
    fila_actual = 1
    while fila_actual <= total_filas and not tiene_precio(fila_actual):
        fila_actual += 1

    productos_esperados = []
    while fila_actual <= total_filas:
        valores = todas_las_filas[fila_actual - 1]
        precio = _limpiar_precio(valor_en(valores, idx_precio))

        if precio is None:
            hay_precio_mas_adelante = any(
                tiene_precio(fila_actual + offset)
                for offset in range(1, CANTIDAD_FILAS_A_REVISAR_ANTES_DE_CORTAR + 1)
            )
            if not hay_precio_mas_adelante:
                break
            fila_actual += 1
            continue

        codigo = limpiar_texto(valor_en(valores, idx_codigo))
        producto = limpiar_texto(valor_en(valores, idx_producto))

        if not codigo or not producto:
            fila_actual += 1
            continue

        productos_esperados.append({
            "fila": fila_actual,
            "codigo": codigo,
            "producto": producto,
            "precio": precio,
        })
        fila_actual += 1

    return productos_esperados

# ─────────────────────────────────────────────────────────────────────────
# Validación principal
# ─────────────────────────────────────────────────────────────────────────

def validar_exportacion(
    ruta_original, ruta_stock_facil, col_codigo, col_producto, col_precio,
    texto_descuento="", texto_aumento="", porcentaje_vendedor=181.5,
    valor_dolar=None, esta_en_dolares=False,
):
    """
    Compara el Excel original contra el Stock Fácil exportado y reporta
    por consola cualquier inconsistencia (productos faltantes, texto que
    no coincide, o precios mal calculados). Devuelve True si todo está
    OK, False si hay al menos un error.
    """
    print("=" * 60)
    print("        🔍 INICIANDO VALIDACIÓN DE EXCEL STOCK FÁCIL        ")
    print("=" * 60)

    if not os.path.exists(ruta_original):
        print(f"❌ Error: El archivo original no existe en: {ruta_original}")
        return False
    if not os.path.exists(ruta_stock_facil):
        print(f"❌ Error: El archivo de Stock Fácil no existe en: {ruta_stock_facil}")
        return False

    # 1) Leer el archivo original con openpyxl, usando las mismas letras
    #    de columna que usó el usuario al procesar (igual que procesador.py)
    try:
        libro_orig = openpyxl.load_workbook(ruta_original, read_only=True, data_only=True)
        hoja_orig = libro_orig.active

        idx_codigo = openpyxl.utils.column_index_from_string(col_codigo) if col_codigo else None
        idx_producto = openpyxl.utils.column_index_from_string(col_producto) if col_producto else None
        idx_precio = openpyxl.utils.column_index_from_string(col_precio) if col_precio else None

        todas_las_filas = [fila for fila in hoja_orig.iter_rows(values_only=True)]
        libro_orig.close()
    except Exception as e:
        print(f"❌ Error al leer el archivo original con openpyxl: {e}")
        return False

    # 2) Leer el Stock Fácil con pandas (la fila 0 siempre trae encabezados limpios)
    try:
        df_stock = pd.read_excel(ruta_stock_facil, sheet_name=0)
    except Exception as e:
        print(f"❌ Error al leer el archivo de Stock Fácil: {e}")
        return False

    # 3) Reconstruir la lista de productos que debería haber en el Stock Fácil
    productos_esperados = _extraer_productos_esperados(todas_las_filas, idx_codigo, idx_producto, idx_precio)

    print(f"✓ Archivo Original: {len(productos_esperados)} productos detectados basándose en columna {col_precio}.")
    print(f"✓ Archivo Stock Fácil: {len(df_stock)} filas encontradas.")

    # Parsear descuentos/aumentos tal como los ingresó el usuario en la interfaz
    descuentos = parsear_porcentajes_encadenados(texto_descuento)
    aumentos = parsear_porcentajes_encadenados(texto_aumento)
    vendedor_factor = float(porcentaje_vendedor or 0)
    dolar_valor = float(valor_dolar) if (esta_en_dolares and valor_dolar) else None

    # Indexar el Stock Fácil por código limpio, para cruzar cada producto al instante
    stock_dict = {}
    for _, fila_stock in df_stock.iterrows():
        codigo_stock = limpiar_texto(fila_stock.get("codigo"))
        if codigo_stock:
            stock_dict[codigo_stock] = fila_stock

    errores = []
    correctos = 0

    # 4) Cruce y validación matemática, producto por producto
    for producto in productos_esperados:
        codigo_orig = producto["codigo"]
        descripcion_orig = producto["producto"]
        precio_orig = producto["precio"]

        if codigo_orig not in stock_dict:
            errores.append(f"Fila {producto['fila']}: El código [{codigo_orig}] se omitió en la exportación.")
            continue

        fila_stock = stock_dict[codigo_orig]

        # Validar que la descripción se haya limpiado igual
        descripcion_stock = str(fila_stock.get("descripcion", "")).strip()
        if descripcion_orig != descripcion_stock:
            errores.append(
                f"Código [{codigo_orig}]: Texto no coincide.\n"
                f"  Esperado: '{descripcion_orig}'\n"
                f"  Encontrado: '{descripcion_stock}'"
            )

        # Validar precio1 (precio de origen, sin ninguna operación aplicada)
        try:
            precio1_stock = float(fila_stock.get("precio1", 0))
            if abs(precio1_stock - round(precio_orig, 2)) > TOLERANCIA_COMPARACION_PRECIOS:
                errores.append(
                    f"Código [{codigo_orig}]: El 'precio1' ({precio1_stock}) "
                    f"no coincide con origen ({round(precio_orig, 2)})"
                )
        except (ValueError, TypeError):
            errores.append(f"Código [{codigo_orig}]: Error de formato en 'precio1'.")

        # Validar precio2 (resultado final, con descuento/aumento/vendedor/dólar aplicados)
        try:
            precio2_stock = float(fila_stock.get("precio2", 0))
            precio2_esperado = calcular_precio_total(
                precio_orig, descuentos, aumentos, vendedor_factor, dolar_valor,
            )
            if abs(precio2_stock - precio2_esperado) > TOLERANCIA_COMPARACION_PRECIOS:
                errores.append(
                    f"Código [{codigo_orig}]: Error matemático en 'precio2'.\n"
                    f"  En Stock Fácil: {precio2_stock}\n"
                    f"  Esperado Real: {precio2_esperado}"
                )
            else:
                correctos += 1
        except (ValueError, TypeError):
            errores.append(f"Código [{codigo_orig}]: Error de formato en 'precio2'.")

    # 5) Presentación de resultados por consola
    print("\n" + "=" * 60)
    print("                      📊 RESUMEN                           ")
    print("=" * 60)
    print(f"• Productos validados con éxito: {correctos}")
    print(f"• Errores / Inconsistencias: {len(errores)}")

    if errores:
        print(f"\n❌ ANOMALÍAS DETECTADAS (Mostrando los primeros {MAXIMO_ERRORES_A_MOSTRAR}):")
        for error in errores[:MAXIMO_ERRORES_A_MOSTRAR]:
            print(f"  - {error}")
        return False

    print("\n¡EXCELENTE! El archivo de Stock Fácil es 100% consistente con tu original.")
    return True