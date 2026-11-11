"""
Facturas recibidas (hoja RecibidasXML): buckets A y B.

Reglas de flujo de efectivo (LIVA 1-B y 5): una factura PUE se considera
pagada en el mes de su emisión. Se toma si:

  * es del año y mes que se declara (Fecha Emision, texto dd/mm/aaaa);
  * Tipo = "Factura";
  * Estado <> "Cancelado" (Estado es manual: vacío = vigente);
  * Método de pago PUE;
  * UsoCFDI deducible (Control: G01, G03);
  * Total > 0. El prellenado del SAT excluye las facturas de $0
    (remanentes de anticipo, facturas de $0.01 con 100% de descuento).

Luego se reparte:
  * Bucket A: bancarizada (Bancarizado = "Sí").
  * Bucket B: efectivo (FormaDePago 01) con Total <= límite (IVA incluido,
    Art. 27 fr. III LISR), excepto combustible, que siempre debe pagarse
    con medio bancarizado.
"""

from diot.acumulados import ImportesPorTasa
from diot.valores_celda import (codigo_uso_cfdi, es_del_periodo, es_si, numero,
                                primeros_dos, texto)

BUCKET_BANCARIZADAS = "A"
BUCKET_EFECTIVO = "B"
TASA_16 = 0.16


def es_factura_deducible_del_mes(fila, col, anio, mes, usos_deducibles):
    """True si la factura entra al universo de facturas PUE deducibles del mes."""
    if not fila.get(col.rfc):
        return False
    if not es_del_periodo(fila.get(col.fecha), anio, mes):
        return False
    if texto(fila.get(col.tipo)) != "Factura":
        return False
    if texto(fila.get(col.estado)) == "Cancelado":
        return False
    if not texto(fila.get(col.metodo_pago)).startswith("PUE"):
        return False
    if codigo_uso_cfdi(fila.get(col.uso_cfdi)) not in usos_deducibles:
        return False
    return numero(fila.get(col.total)) > 0


def bucket_de_la_factura(fila, col, limite_efectivo):
    """'A' bancarizada, 'B' efectivo deducible, o None si no se acredita."""
    if es_si(fila.get(col.bancarizado)):
        return BUCKET_BANCARIZADAS
    pagada_en_efectivo = primeros_dos(fila.get(col.forma_pago)) == "01"
    dentro_del_limite = numero(fila.get(col.total)) <= limite_efectivo
    es_combustible = es_si(fila.get(col.combustible))
    if pagada_en_efectivo and dentro_del_limite and not es_combustible:
        return BUCKET_EFECTIVO
    return None


def importes_tasa_16(fila, col):
    """Base e IVA 16%. Base = IVA / 0.16 (así absorbe IEPS y deja fuera
    conceptos exentos o a tasa 0% de facturas mixtas)."""
    iva = numero(fila.get(col.iva_16))
    base = iva / TASA_16 if iva else 0.0
    return base, iva


def importes_de_factura(fila, col, bucket):
    """Bolsa de importes de una factura ya clasificada en su bucket."""
    importes = ImportesPorTasa()
    importes.base_16, importes.iva_16 = importes_tasa_16(fila, col)
    importes.iva_retenido = numero(fila.get(col.retenido_iva))
    return importes
