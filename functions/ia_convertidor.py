"""
ia_convertidor.py

Envía productos de un Excel a Gemini en lotes para que los clasifique
por categoría y agrupe variantes de medida (mismo producto, distinta
medida/tamaño). Devuelve el resultado final en JSON, ya parseado.

Incluye en un solo lugar: la lista de categorías, el prompt (editable
acá abajo sin tocar el resto del código) y la lógica de envío a la IA.
"""

import json
from google import genai
from google.genai import types

from functions.ia_config import obtener_api_key_gemini
from functions.procesador import procesar_excel

# ─────────────────────────────────────────────────────────────────────────
# Configuración editable
# ─────────────────────────────────────────────────────────────────────────

TAMANO_LOTE = 40
MODELO = "gemini-2.0-flash"

CATEGORIAS = [
    "Termicas",
    "Cables",
    "Tomas y puntos",
    "Capacitores",
    "Cajas electricas",
    "Caños para instalacion",
    "Prolongadores/zapatillas",
    "Herramientas",
    "Pilas",
    "Aislacion electrica",
    "Precintos",
    "Cable canal",
    "Electronica",
    "Jabalinas",
    "Iluminacion",
    "Lamparas led",
    "Adhesivos y siliconas",
    "Contactores y guardamotor",
    "Candados",
    "Discos y mechas",
]

PROMPT_CLASIFICACION = """Sos un asistente que clasifica productos de una ferretería eléctrica argentina.

Vas a recibir una lista de productos con código, descripción, familia (categoría interna del proveedor) y precio.

IMPORTANTE sobre la "familia": es solo un dato de referencia del proveedor y puede ser imprecisa. NO agrupes productos únicamente porque comparten la misma familia. El criterio principal para agrupar variantes es que el NOMBRE/DESCRIPCIÓN sea el mismo producto, cambiando solo un número de medida/tamaño.

Tu tarea es:

1. Agrupar productos que sean EL MISMO PRODUCTO BASE pero con distinta medida (por ejemplo, mismo nombre pero cambia solo un número, como "8mm", "12mm", "16mm"). Cada uno de esos productos agrupados debe convertirse en una VARIANTE dentro de un mismo producto general.

2. Si un producto NO tiene otras medidas relacionadas en la lista, va como un producto con UNA SOLA variante.

3. Asignar a cada producto general UNA categoría de esta lista exacta (nunca inventes una categoría nueva ni cambies el texto):

{categorias}

4. El nombre del producto general debe ser un nombre limpio y genérico (sin el número de medida específico).

Reglas importantes:
- Nunca inventes códigos, precios o medidas que no estén en los datos que te paso.
- Si no estás seguro de la categoría, elegí la más cercana posible de la lista, nunca dejes la categoría vacía.
- Cada variante debe tener: código, medida (o "Único" si no aplica) y precio.

Productos a clasificar:

{productos}
"""

ESQUEMA_RESPUESTA = {
    "type": "object",
    "properties": {
        "productos": {
            "type": "array",
            "items": {
                "type": "object",
                "properties": {
                    "nombre_general": {"type": "string"},
                    "categoria": {"type": "string"},
                    "variantes": {
                        "type": "array",
                        "items": {
                            "type": "object",
                            "properties": {
                                "codigo": {"type": "string"},
                                "medida": {"type": "string"},
                                "precio": {"type": "number"},
                            },
                            "required": ["codigo", "medida", "precio"],
                        },
                    },
                },
                "required": ["nombre_general", "categoria", "variantes"],
            },
        }
    },
    "required": ["productos"],
}


# ─────────────────────────────────────────────────────────────────────────
# Helpers internos
# ─────────────────────────────────────────────────────────────────────────

def _preparar_productos_desde_excel(ruta_archivo, columna_codigo, columna_producto, columna_precio, columna_familia):
    resultado = procesar_excel(
        ruta_archivo,
        columna_codigo=columna_codigo,
        columna_producto=columna_producto,
        columna_precio=columna_precio,
        columna_familia=columna_familia,
    )

    productos = resultado["productos"]

    for producto in productos:
        producto["precio"] = round(producto["precio"], 2)

    productos.sort(key=lambda p: p.get("familia", ""))

    return productos, resultado["errores"]


def _armar_lotes(productos, tamano_lote=TAMANO_LOTE):
    for i in range(0, len(productos), tamano_lote):
        yield productos[i:i + tamano_lote]


def _formatear_productos_para_prompt(lote):
    lineas = []
    for producto in lote:
        lineas.append(
            f"- Código: {producto['codigo']} | "
            f"Descripción: {producto['producto']} | "
            f"Familia: {producto.get('familia', '')} | "
            f"Precio: {producto['precio']}"
        )
    return "\n".join(lineas)


# ─────────────────────────────────────────────────────────────────────────
# API pública
# ─────────────────────────────────────────────────────────────────────────

def clasificar_excel(ruta_archivo, columna_codigo, columna_producto, columna_precio, columna_familia=None, callback_progreso=None):
    """
    Punto de entrada principal: recibe la ruta de un Excel (tal cual lo
    sube el usuario) y las columnas ya elegidas en la UI. Devuelve el
    JSON final ya clasificado y agrupado.

    callback_progreso: función opcional que se llama después de cada
    lote con (lote_actual, total_lotes), para mostrar progreso en la UI.
    """
    api_key = obtener_api_key_gemini()
    if not api_key:
        raise ValueError("No hay una API Key de Gemini configurada. Configurala en la pantalla de Configuración.")

    productos, errores_lectura = _preparar_productos_desde_excel(
        ruta_archivo, columna_codigo, columna_producto, columna_precio, columna_familia,
    )

    if not productos:
        return {"productos": [], "errores_lectura_excel": errores_lectura}

    cliente = genai.Client(api_key=api_key)
    categorias_texto = "\n".join(f"- {categoria}" for categoria in CATEGORIAS)

    lotes = list(_armar_lotes(productos))
    total_lotes = len(lotes)
    resultado_final = []

    for i, lote in enumerate(lotes, start=1):
        productos_texto = _formatear_productos_para_prompt(lote)

        prompt_final = PROMPT_CLASIFICACION.format(
            categorias=categorias_texto,
            productos=productos_texto,
        )

        respuesta = cliente.models.generate_content(
            model=MODELO,
            contents=prompt_final,
            config=types.GenerateContentConfig(
                response_mime_type="application/json",
                response_schema=ESQUEMA_RESPUESTA,
            ),
        )

        datos_lote = json.loads(respuesta.text)
        resultado_final.extend(datos_lote["productos"])

        if callback_progreso:
            callback_progreso(i, total_lotes)

    return {
        "productos": resultado_final,
        "errores_lectura_excel": errores_lectura,
    }