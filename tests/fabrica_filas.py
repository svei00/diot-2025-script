"""
Filas SINTÉTICAS para las pruebas. Nunca se lee el libro real (trae datos
de clientes). Cada función regresa una fila "normal" del universo y acepta
cambios por nombre de columna, p. ej. factura(**{"Total": 0}).
"""

ANIO, MES = 2026, 10
FECHA_DEL_MES = "15/10/2026"

ENCABEZADOS_FACTURAS = [
    "Tipo", "Fecha Emision", "RFC Emisor", "LugarDeExpedicion", "UsoCFDI", "SubTotal",
    "Descuento", "IVA 16%", "Retenido IVA", "Total", "Total Trasladados",
    "FormaDePago", "Metodo de Pago", "Conceptos", "Combustible", "IVA 8%",
    "Nivel Educativo", "Estado", "Bancarizado",
]

ENCABEZADOS_REP = [
    "Fecha Emision", "RFC Emisor CFDI", "Lugar de Expedicion CFDI", "FechaPago", "MonedaP",
    "IdDocumento Relacionado", "MonedaDR", "EquivalenciaDR", "MetodoDePagoDR",
    "ImpPagado", "IVA Excento Base", "IVA Cero Base", "IVA 8 Base", "IVA 8 Importe",
    "IVA 16 Base", "IVA 16 Importe", "Ret IVA 10.667 Base", "Ret IVA 10.667 Importe",
    "Ret IVA 4 Base", "Ret IVA 4 Importe", "BancarizadoP", "Estado", "EsDuplicado", "RelUUID",
]


def factura(**cambios):
    fila = {
        "Tipo": "Factura", "Fecha Emision": FECHA_DEL_MES, "RFC Emisor": "AAA010101AAA",
        "LugarDeExpedicion": "45000", "UsoCFDI": "G03 - Gastos en general",
        "SubTotal": 1000.0, "Descuento": 0.0, "IVA 16%": 160.0, "Retenido IVA": 0.0,
        "Total": 1160.0, "Total Trasladados": 160.0, "FormaDePago": "03 - Transferencia",
        "Metodo de Pago": "PUE - Pago en una sola exhibición", "Conceptos": "Material",
        "Combustible": "No", "IVA 8%": 0.0, "Nivel Educativo": None, "Estado": "",
        "Bancarizado": "Sí",
    }
    fila.update(cambios)
    return fila


def factura_efectivo(**cambios):
    datos = {"Bancarizado": "No", "FormaDePago": "01 - Efectivo"}
    datos.update(cambios)
    return factura(**datos)


def rep(**cambios):
    fila = {
        "Fecha Emision": FECHA_DEL_MES, "RFC Emisor CFDI": "BBB010101BBB",
        "Lugar de Expedicion CFDI": "45000", "FechaPago": FECHA_DEL_MES, "MonedaP": "MXN",
        "IdDocumento Relacionado": "uuid-1", "MonedaDR": "MXN", "EquivalenciaDR": 1,
        "MetodoDePagoDR": "PPD", "ImpPagado": 1160.0, "IVA Excento Base": 0.0,
        "IVA Cero Base": 0.0, "IVA 8 Base": 0.0, "IVA 8 Importe": 0.0,
        "IVA 16 Base": 1000.0, "IVA 16 Importe": 160.0, "Ret IVA 10.667 Base": 0.0,
        "Ret IVA 10.667 Importe": 0.0, "Ret IVA 4 Base": 0.0, "Ret IVA 4 Importe": 0.0,
        "BancarizadoP": "Sí", "Estado": "", "EsDuplicado": 0, "RelUUID": "OK",
    }
    fila.update(cambios)
    return fila


def construir(facturas=(), reps=(), **opciones):
    """Atajo: corre build_diot con las filas sintéticas del mes de prueba."""
    from diot.construir import build_diot
    return build_diot(ENCABEZADOS_FACTURAS, list(facturas), ENCABEZADOS_REP, list(reps),
                      ANIO, MES, opciones.pop("op_map", {}), **opciones)
