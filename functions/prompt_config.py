"""
prompt_config.py

Acceso centralizado al prompt de clasificación de catálogo para Tienda
Web. El prompt es editable desde Configuración y se guarda persistente
en config_manager, igual que el resto de la configuración de Lumen.
"""

from functions.config_manager import obtener_valor

CLAVE_PROMPT_CATALOGO = "prompt_catalogo_tienda_web"

PROMPT_CATALOGO_POR_DEFECTO = """Te voy a pasar un Excel de un proveedor de ferretería eléctrica argentina \
que tiene columnas con fórmulas (descuento, aumento, % vendedor, precio \
total). Necesito que uses código (Python con openpyxl, no que calcules \
"a ojo") para leer el archivo y hacer lo siguiente:

1. LECTURA DE DATOS
   - Los encabezados están en la fila 2: Código (A), Descripción (B), \
Familia (C), Precio (D), Descuento % (E), Aumento % (F), \
% Vendedor (G), Precio Total (H).
   - Los productos arrancan en la fila 3.
   - Los porcentajes de Descuento, Aumento y % Vendedor están en la \
fila 2 de las columnas E, F y G (mismo valor para todas las filas).

2. CÁLCULO DEL PRECIO FINAL (hacer con código, no a mano)
   - La columna H ("Precio Total") es una fórmula. Si al leer el archivo \
con openpyxl (data_only=True) el valor cacheado está vacío/None, \
calculalo vos mismo con código, en este orden exacto:
       a) Precio con descuento = Precio (columna D) × (1 - Descuento%)
       b) Precio con aumento = resultado anterior × (1 + Aumento%)
       c) Precio Total = resultado anterior × % Vendedor
     Redondeá cada paso a 2 decimales.
   - Si la columna H ya tiene un valor cargado, usá ese directamente.
   - El precio final SIEMPRE debe ser este "Precio Total", nunca el \
precio base de la columna D solo.

3. AGRUPAR VARIANTES
   - Agrupá productos que sean el mismo producto base pero con distinta \
medida/tamaño (por ejemplo, mismo nombre pero cambia solo un número \
como "8mm", "12mm", "16mm"). Estos se convierten en VARIANTES de un \
mismo producto general.
   - Si un producto no tiene otras medidas relacionadas, va como producto \
con una sola variante.
   - El nombre del producto general NUNCA debe incluir el número de \
medida que varía entre variantes (ese número va solo en el campo \
"medida" de cada variante). Números que son parte fija del producto \
(por ejemplo el ancho de una banda, si no cambia entre variantes) sí \
pueden quedarse en el nombre.
   - La "Familia" (columna C) es solo una pista de referencia del \
proveedor, puede ser imprecisa — no agrupes solo por tener la misma \
familia, el criterio principal es que el nombre/descripción sea \
el mismo producto.

4. CATEGORIZACIÓN — usar SIEMPRE una de estas categorías disponibles:
   Termicas, Cables, Tomas y puntos, Capacitores, Cajas electricas, \
Caños para instalacion, Prolongadores/zapatillas, Herramientas, \
Pilas, Aislacion electrica, Precintos, Cable canal, Electronica, \
Jabalinas, Iluminacion, Lamparas led, Adhesivos y siliconas, \
Contactores y guardamotor, Candados, Discos y mechas, Otros.

   Reglas de categorización:
   - Analizá el nombre y la familia del producto y elegí la categoría \
que MEJOR describa su función real (por ejemplo, una abrazadera o \
herramienta de sujeción va en "Herramientas", un cable va en \
"Cables", una lámpara LED va en "Lamparas led", etc.).
   - "Otros" es SOLO para productos que genuinamente no son de \
electricidad ni ferretería (por ejemplo alimentos, artículos de \
cocina, etc.). NO uses "Otros" como respuesta por defecto — antes \
de asignarla, verificá que ninguna de las otras 19 categorías \
describa mejor al producto.
   - Todo producto DEBE tener una categoría asignada, sin excepción, \
nunca la dejes vacía.

5. GENERAR EL EXCEL DE SALIDA (con código, usando openpyxl)
   Columnas fijas: Código | Nombre | Medida | Categoría | Precio

   - Fila del PRODUCTO GENERAL: nombre en "Nombre" (sin medidas), \
categoría en "Categoría". Código, Medida y Precio quedan vacíos en \
esta fila. TODA la fila (las 5 columnas, arrancando desde la \
columna A) debe estar pintada de amarillo (color FFF59D).
   - Filas de VARIANTES (van justo debajo de su producto general, sin \
pintar):
       • Código: el código real de esa variante.
       • Nombre: repetí el MISMO nombre del producto general (para que \
cada fila sea identificable sin tener que buscar la fila \
amarilla de arriba).
       • Medida: la medida de esa variante (o "Único" si el producto no \
tiene variantes de medida).
       • Categoría: vacío en las filas de variante (la categoría solo \
va en la fila del producto general).
       • Precio: el Precio Total calculado en el paso 2.

Antes de entregar el archivo final, verificá con código (no de memoria):
- Que ningún nombre_general tenga números de medida.
- Que ningún producto haya quedado sin categoría.
- Que la categoría "Otros" se use solo en casos realmente justificados, \
no como respuesta por defecto (contá cuántos productos quedaron en \
"Otros" y revisá si tiene sentido).
- Que los precios usados sean el "Precio Total" con el % vendedor \
aplicado, no el precio base de la columna D.

Generá el archivo Excel final y compartímelo para descargar.
"""


def obtener_prompt_catalogo():
    """Devuelve el prompt guardado por el usuario, o el que viene por defecto si nunca lo editó."""
    return obtener_valor(CLAVE_PROMPT_CATALOGO, PROMPT_CATALOGO_POR_DEFECTO)