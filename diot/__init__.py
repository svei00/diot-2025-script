"""
Paquete `diot`: genera el TXT de carga masiva de la DIOT 2025 a partir del
libro "Registro de Facturación 2026vN".

Mapa del paquete (cada archivo hace UNA cosa):

    valores_celda.py      convierte lo que trae una celda (texto, número, fecha) a algo usable
    redondeo.py           redondeo del Art. 20 CFF y la base mínima que exige el validador del SAT
    lectura_libro.py      abre el libro (.xlsb/.xlsm/.xlsx) en solo lectura y devuelve tablas
    columnas.py           ubica, por encabezado, las columnas de RecibidasXML y PagosRecibidasXML
    facturas.py           universo de facturas (buckets A y B) e importes por tasa de cada factura
    rep.py                universo de REP (bucket C), conversión de moneda e importes por tasa
    acumulados.py         sumas por proveedor y por bucket
    layout_txt.py         posiciones oficiales de los 54 campos y armado de cada renglón
    construir.py          build_diot(): la función pura que junta todo (sin I/O)
    resumen_libro.py      lee la hoja Resumen por encabezado (para conciliar)
    declaracion_libro.py  lee la hoja Declaracion por etiqueta (para conciliar)
    conciliacion.py       arma el texto de conciliación que ve el contador
    archivo_salida.py     nombre del TXT, configuración guardada y escritura del archivo
"""

from diot.construir import build_diot          # noqa: F401  (API pública)
from diot.lectura_libro import read_sheet      # noqa: F401  (API pública)
