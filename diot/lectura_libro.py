"""
Lectura del libro "Registro de Facturación" en SOLO LECTURA.

Formatos aceptados:
  * .xlsb               -> pyxlsb
  * .xlsm/.xlsx/.xltm/.xltx -> openpyxl con data_only=True, es decir, se leen
    los VALORES EN CACHÉ de las fórmulas (lo que Excel escribió al guardar).

Una "rejilla" (grid) es una lista de filas; cada fila es un dict
{índice_de_columna: valor} sin las celdas vacías. El índice 0 es la columna A.

Este módulo nunca escribe en el libro.
"""

import os
import re

from diot.valores_celda import texto

EXTENSIONES_XLSB = (".xlsb",)
EXTENSIONES_OPENPYXL = (".xlsm", ".xlsx", ".xltm", ".xltx")
EXTENSIONES_LIBRO = EXTENSIONES_XLSB + EXTENSIONES_OPENPYXL


# ---------------------------------------------------------------- .xlsb
def _rejilla_xlsb(ruta, hoja):
    import pyxlsb
    with pyxlsb.open_workbook(ruta) as libro:
        with libro.get_sheet(hoja) as hoja_xlsb:
            return [{celda.c: celda.v for celda in fila} for fila in hoja_xlsb.rows()]


# ---------------------------------------------------------------- .xlsm / .xlsx
def _abrir_openpyxl(ruta, solo_valores):
    try:
        import openpyxl
    except ImportError:
        raise ImportError("Para leer .xlsx/.xlsm hace falta openpyxl:  pip install openpyxl")
    return openpyxl.load_workbook(ruta, read_only=True, data_only=solo_valores,
                                  keep_links=False)


def _hoja_tiene_formulas(ruta, hoja):
    """Segunda pasada (sin data_only) para saber si la hoja tiene fórmulas."""
    libro = _abrir_openpyxl(ruta, solo_valores=False)
    try:
        if hoja not in libro.sheetnames:
            return False
        for numero_fila, fila in enumerate(libro[hoja].iter_rows(values_only=True)):
            if numero_fila == 0:
                continue                                   # encabezados
            if any(isinstance(v, str) and v.startswith("=") for v in fila):
                return True
    finally:
        libro.close()
    return False


def _rejilla_openpyxl(ruta, hoja):
    libro = _abrir_openpyxl(ruta, solo_valores=True)
    try:
        if hoja not in libro.sheetnames:
            raise KeyError("la hoja '%s' no existe en %s" % (hoja, os.path.basename(ruta)))
        rejilla, hay_datos = [], False
        for fila in libro[hoja].iter_rows(values_only=True):
            celdas = {i: v for i, v in enumerate(fila) if v is not None and v != ""}
            rejilla.append(celdas)
            if len(rejilla) > 1 and celdas:
                hay_datos = True
    finally:
        libro.close()
    if len(rejilla) > 1 and not hay_datos and _hoja_tiene_formulas(ruta, hoja):
        # El archivo lo escribió un programa, no Excel: las fórmulas no traen caché.
        raise ValueError(
            "La hoja '%s' de %s se lee vacía: el archivo no trae los valores\n"
            "calculados de sus fórmulas. Ábrelo en Excel y guárdalo para que se\n"
            "escriba la caché, o usa la versión .xlsb."
            % (hoja, os.path.basename(ruta)))
    while rejilla and not rejilla[-1]:                     # filas vacías al final
        rejilla.pop()
    return rejilla


# ---------------------------------------------------------------- API
def read_grid(ruta, hoja):
    """Rejilla cruda de una hoja (para Control, Resumen y Declaracion)."""
    extension = os.path.splitext(str(ruta))[1].lower()
    if extension in EXTENSIONES_OPENPYXL:
        return _rejilla_openpyxl(ruta, hoja)
    if extension in EXTENSIONES_XLSB:
        return _rejilla_xlsb(ruta, hoja)
    raise ValueError("Formato no soportado: '%s'. Usa %s"
                     % (extension or ruta, ", ".join(EXTENSIONES_LIBRO)))


def read_sheet(ruta, hoja):
    """Hoja con encabezados en la fila 1 -> (lista_de_encabezados, lista_de_filas_dict)."""
    rejilla = read_grid(ruta, hoja)
    if not rejilla or not rejilla[0]:
        return [], []
    encabezados = [str(rejilla[0].get(i, "")).strip() for i in range(max(rejilla[0]) + 1)]
    indice_por_nombre = {nombre: i for i, nombre in enumerate(encabezados) if nombre}
    filas = [{nombre: celdas.get(i) for nombre, i in indice_por_nombre.items()}
             for celdas in rejilla[1:]]
    return encabezados, filas


def read_control(ruta):
    """Hoja Control: (año fiscal, límite de efectivo, usos CFDI deducibles).

    Las etiquetas están en la columna H y los valores en la I. Si la hoja no
    existe, regresa (None, None, set()) y se usan los valores por omisión.
    """
    anio, limite, usos = None, None, set()
    try:
        rejilla = read_grid(ruta, "Control")
    except Exception:
        return anio, limite, usos
    leyendo_usos = False
    for fila in rejilla:
        etiqueta, valor = texto(fila.get(7)).lower(), fila.get(8)
        if etiqueta.startswith("año fiscal") and isinstance(valor, (int, float)):
            anio = int(valor)
        elif etiqueta.startswith("límite pago en efectivo") and isinstance(valor, (int, float)):
            limite = float(valor)
        elif etiqueta.startswith("usos cfdi deducibles"):
            leyendo_usos = True
            continue
        if leyendo_usos and re.fullmatch(r"[A-Z]\d{2}", texto(fila.get(7))):
            usos.add(texto(fila.get(7)))
    return anio, limite, usos
