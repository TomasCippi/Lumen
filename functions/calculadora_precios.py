"""
calculadora_precios.py

Única fuente de verdad para la cadena de cálculo: descuento -> aumento
-> % vendedor -> dólar. La usan exportador.py, validador_stock.py y
convertidor.py, para no tener la misma fórmula copiada 3 veces.
"""


def parsear_porcentajes_encadenados(texto: str) -> list:
    if not texto:
        return []
    partes = [p for p in texto.split("-") if p.strip() != ""]
    try:
        return [float(p) for p in partes]
    except ValueError:
        return []


def calcular_precio_total(precio_base, descuentos, aumentos, porcentaje_vendedor, valor_dolar=None):
    if porcentaje_vendedor == 0:
        import warnings
        warnings.warn("El porcentaje del vendedor es 0%, el precio final va a dar 0.")

    precio = precio_base
    for descuento in descuentos:
        precio = round(precio * (1 - descuento / 100), 2)
    for aumento in aumentos:
        precio = round(precio * (1 + aumento / 100), 2)
    precio = round(precio * (porcentaje_vendedor / 100), 2)
    if valor_dolar:
        precio = round(precio * valor_dolar, 2)
    return precio