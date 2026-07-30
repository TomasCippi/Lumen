"""
preparar_excel.py

Preparación del Excel antes de leerlo.
--------------------------------------
Dos problemas reales que puede traer un Excel del usuario:

1. Es un .xls viejo (formato Excel 97-2003): openpyxl no puede leerlo.
   -> Se convierte a .xlsx con xlrd + openpyxl.

2. Tiene columnas con FÓRMULAS (ej. =SI(B7="";"";BUSCARV(B7;precios2;7))).
   openpyxl con data_only=True solo lee el ÚLTIMO VALOR que quedó
   guardado en el archivo la última vez que se abrió/guardó en Excel.
   Si ese valor cacheado no existe (o el archivo nunca se recalculó),
   lo que se lee es None o texto vacío.
   -> La única forma confiable de tener SIEMPRE el valor numérico real,
      sin importar la complejidad de la fórmula (BUSCARV, rangos con
      nombre, etc.), es abrir el archivo con Excel de verdad (vía COM,
      con pywin32), forzar el recálculo, y guardarlo. Recién ahí se
      lee con openpyxl.

Requiere que la PC tenga Microsoft Excel instalado (es el escenario
normal en una PC de oficina/comercio). Si no lo tiene, se avisa con un
error claro en vez de fallar silenciosamente.
"""

import os
import tempfile

import openpyxl


# ─────────────────────────────────────────────────────────────────────────
# Paso 1: conversión de .xls viejo a .xlsx
# ─────────────────────────────────────────────────────────────────────────

def convertir_xls_a_xlsx(ruta_archivo: str) -> str:
    """
    Si el archivo es un .xls viejo, lo convierte a .xlsx en una carpeta
    temporal y devuelve la ruta del archivo nuevo. Si ya es .xlsx (o
    cualquier otra extensión), devuelve la misma ruta sin tocar nada.
    """
    if not ruta_archivo.lower().endswith(".xls"):
        return ruta_archivo

    import xlrd  # solo se importa si realmente hace falta

    libro_viejo = xlrd.open_workbook(ruta_archivo)
    hoja_vieja = libro_viejo.sheet_by_index(0)

    libro_nuevo = openpyxl.Workbook()
    hoja_nueva = libro_nuevo.active

    for num_fila in range(hoja_vieja.nrows):
        for num_col in range(hoja_vieja.ncols):
            valor = hoja_vieja.cell_value(num_fila, num_col)
            hoja_nueva.cell(row=num_fila + 1, column=num_col + 1, value=valor)

    carpeta_temp = tempfile.gettempdir()
    nombre_base = os.path.splitext(os.path.basename(ruta_archivo))[0]
    ruta_nueva = os.path.join(carpeta_temp, f"{nombre_base}_convertido.xlsx")

    libro_nuevo.save(ruta_nueva)
    return ruta_nueva


# ─────────────────────────────────────────────────────────────────────────
# Paso 2: recálculo de fórmulas vía Excel (COM / pywin32)
# ─────────────────────────────────────────────────────────────────────────

def recalcular_formulas(ruta_archivo: str) -> str:
    """
    Abre el archivo con Excel real (vía COM), fuerza un recálculo
    completo y lo guarda en una copia temporal como .xlsx. Devuelve la
    ruta de esa copia, que ya tiene los valores de las fórmulas
    actualizados y cacheados.

    Requiere Microsoft Excel instalado + el paquete pywin32.
    """
    try:
        import win32com.client as win32
    except ImportError as e:
        raise RuntimeError(
            "Para recalcular fórmulas hace falta tener Microsoft Excel "
            "instalado y el paquete 'pywin32' (pip install pywin32)."
        ) from e

    carpeta_temp = tempfile.gettempdir()
    nombre_base = os.path.splitext(os.path.basename(ruta_archivo))[0]
    ruta_salida = os.path.join(carpeta_temp, f"{nombre_base}_recalculado.xlsx")

    try:
        excel = win32.Dispatch("Excel.Application")
    except AttributeError:
        # La caché de pywin32 (gen_py) está corrupta. La borramos y reintentamos.
        import shutil
        carpeta_gen_py = os.path.join(os.environ.get("LOCALAPPDATA", ""), "Temp", "gen_py")
        if os.path.isdir(carpeta_gen_py):
            shutil.rmtree(carpeta_gen_py, ignore_errors=True)
        excel = win32.Dispatch("Excel.Application")

    excel.Visible = False
    excel.DisplayAlerts = False

    libro = None
    try:
        libro = excel.Workbooks.Open(os.path.abspath(ruta_archivo))
        excel.CalculateFullRebuild()
        excel.CalculateUntilAsyncQueriesDone()
        libro.SaveAs(ruta_salida, FileFormat=51)  # 51 = xlOpenXMLWorkbook (.xlsx)
    finally:
        if libro is not None:
            try:
                libro.Close(SaveChanges=False)
            except Exception:
                pass
        excel.Quit()

    return ruta_salida


# ─────────────────────────────────────────────────────────────────────────
# Punto de entrada
# ─────────────────────────────────────────────────────────────────────────

def preparar_excel(ruta_archivo: str) -> str:
    """
    Punto de entrada único: aplica todos los pasos necesarios (convertir
    .xls si hace falta, y recalcular fórmulas) y devuelve la ruta final,
    lista para pasarle a detectar_columnas / procesar_excel.
    """
    ruta = convertir_xls_a_xlsx(ruta_archivo)
    ruta = recalcular_formulas(ruta)
    return ruta