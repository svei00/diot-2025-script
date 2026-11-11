"""
Conversión de valores de celda.

El libro llega por tres caminos distintos (pyxlsb para .xlsb, openpyxl para
.xlsm/.xlsx, o filas sintéticas en las pruebas) y cada uno entrega los datos
con tipos diferentes: una fecha puede venir como texto "21/06/2026", como
número serial de Excel o como datetime. Estas funciones normalizan todo eso
para que el resto del código no tenga que preocuparse.
"""

import datetime as _dt
import re

# Excel cuenta los días desde el 30/12/1899 (incluye el famoso bug del 29/02/1900).
INICIO_FECHAS_EXCEL = _dt.datetime(1899, 12, 30)


def numero(valor):
    """Convierte la celda a float. Vacío, texto o error -> 0.0."""
    try:
        return float(valor)
    except (TypeError, ValueError):
        return 0.0


def texto(valor):
    """Convierte la celda a texto sin espacios a los lados. Vacío -> ''."""
    return str(valor).strip() if valor is not None else ""


def es_si(valor):
    """True si la celda dice "Sí"/"Si" (con o sin acento) o trae un TRUE cacheado."""
    if isinstance(valor, bool):
        return valor
    return texto(valor).lower() in ("si", "sí")


def primeros_dos(valor):
    """'01 - Efectivo' -> '01'. Sirve para FormaDePago."""
    return texto(valor)[:2]


def codigo_uso_cfdi(valor):
    """'G03 - Gastos en general' -> 'G03'."""
    return texto(valor)[:3]


def anio_y_mes(valor):
    """Devuelve (año, mes) de una fecha venga como venga; (0, 0) si no se reconoce.

    Ojo: las fechas del libro son TEXTO dd/mm/aaaa. Nunca se interpretan con la
    configuración regional (eso confunde día y mes sin dar error).
    """
    if valor is None or valor == "":
        return 0, 0
    if isinstance(valor, (_dt.datetime, _dt.date)):
        return valor.year, valor.month
    if isinstance(valor, bool):
        return 0, 0
    if isinstance(valor, (int, float)):                       # serial de Excel
        try:
            fecha = INICIO_FECHAS_EXCEL + _dt.timedelta(days=float(valor))
        except (OverflowError, ValueError):
            return 0, 0
        return fecha.year, fecha.month
    cadena = str(valor).strip()
    dia_mes_anio = re.match(r"(\d{1,2})[/-](\d{1,2})[/-](\d{4})", cadena)
    if dia_mes_anio:
        return int(dia_mes_anio.group(3)), int(dia_mes_anio.group(2))
    anio_mes_dia = re.match(r"(\d{4})-(\d{1,2})-(\d{1,2})", cadena)   # ISO
    if anio_mes_dia:
        return int(anio_mes_dia.group(1)), int(anio_mes_dia.group(2))
    return 0, 0


def es_del_periodo(valor_fecha, anio, mes):
    """True si la fecha cae en el año y mes que se declaran."""
    return anio_y_mes(valor_fecha) == (anio, mes)
