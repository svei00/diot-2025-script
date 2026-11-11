"""
Recibos electrónicos de pago recibidos (hoja PagosRecibidasXML): bucket C.

Un REP ampara el pago de una factura PPD. En flujo de efectivo el IVA se
acredita en el mes del PAGO (FechaPago), no en el de la factura.

Universo (idéntico a Resumen!X:Z y a Declaracion):
  * FechaPago en el año y mes que se declara;
  * Estado <> "Cancelado";
  * BancarizadoP = "Sí";
  * EsDuplicado <> 1;
  * RelUUID = "OK" (la factura que paga sí está en RecibidasXML).

Moneda extranjera: cuando la factura pagada (MonedaDR) no es MXN, todos los
importes del documento relacionado vienen en esa moneda. EquivalenciaDR dice
cuántas unidades de la moneda del documento equivalen a una de la moneda del
pago, así que, con el pago en MXN:

        importe en MXN = importe del documento / EquivalenciaDR

Ejemplo real (octubre 2026, BUCA RECUBRIMIENTOS): equivalencia 0.0544401406,
base 2,045.50 USD -> 37,573.38 MXN.

Si el pago mismo (MonedaP) no es MXN, la conversión necesitaría también el
TipoCambioP. No se adivina: el REP se excluye y se reporta para revisarlo.
"""

from diot.acumulados import ImportesPorTasa
from diot.columnas import valor
from diot.valores_celda import es_del_periodo, es_si, numero, texto

MONEDA_NACIONAL = "MXN"


def motivo_para_excluir_rep(fila, col, anio, mes):
    """Regresa (se_incluye, motivo_a_reportar).

    `motivo_a_reportar` solo se llena cuando la exclusión merece que el
    contador la investigue (factura relacionada faltante, pago en otra moneda).
    Las exclusiones normales (otro mes, cancelado, duplicado) no se reportan.
    """
    if not fila.get(col.rfc):
        return False, None
    if not es_del_periodo(fila.get(col.fecha_pago), anio, mes):
        return False, None
    if texto(fila.get(col.estado)) == "Cancelado":
        return False, None
    if not es_si(fila.get(col.bancarizado)):
        return False, None
    if col.es_duplicado and numero(fila.get(col.es_duplicado)) == 1:
        return False, None
    if col.rel_uuid and texto(fila.get(col.rel_uuid)) != "OK":
        return False, "RelUUID=%s" % texto(fila.get(col.rel_uuid))
    moneda_del_pago = texto(valor(fila, col.moneda_pago)).upper()
    if moneda_del_pago and moneda_del_pago != MONEDA_NACIONAL:
        return False, "MonedaP=%s (revisar a mano)" % moneda_del_pago
    return True, None


def equivalencia_del_documento(fila, col):
    """EquivalenciaDR como divisor. Vacía o 0 -> 1 (documento en MXN)."""
    equivalencia = numero(valor(fila, col.equivalencia))
    return equivalencia if equivalencia > 0 else 1.0


def en_pesos(fila, columna, equivalencia):
    """Importe de una columna del documento relacionado, convertido a MXN."""
    return numero(valor(fila, columna)) / equivalencia


def importes_de_rep(fila, col):
    """Bolsa de importes de un REP del universo, ya convertidos a MXN."""
    equivalencia = equivalencia_del_documento(fila, col)
    importes = ImportesPorTasa()
    importes.base_16 = en_pesos(fila, col.base_16, equivalencia)
    importes.iva_16 = en_pesos(fila, col.iva_16, equivalencia)
    importes.base_8 = en_pesos(fila, col.base_8, equivalencia)
    importes.iva_8 = en_pesos(fila, col.iva_8, equivalencia)
    importes.exento = en_pesos(fila, col.base_exento, equivalencia)
    importes.tasa_cero = en_pesos(fila, col.base_cero, equivalencia)
    importes.iva_retenido = sum(en_pesos(fila, columna, equivalencia)
                                for columna in col.retenciones_iva)
    return importes


def paga_una_factura_pue(fila, col, metodo_por_uuid):
    """True si la factura que paga el REP es PUE.

    No debería existir: una factura PUE ya se consideró pagada al emitirse.
    El SAT la cuenta dos veces (factura y REP) y el script hace lo mismo
    para cuadrar con el prellenado (decisión P5), pero lo avisa: es una
    deducción duplicada y hay que pedir al proveedor que cancele el REP.

    Pagos 2.0 ya no trae MetodoDePagoDR, así que el método se busca en la
    factura relacionada (IdDocumento Relacionado -> UUID de RecibidasXML).
    MetodoDePagoDR se usa solo si viene lleno (REP de Pagos 1.0).
    """
    uuid = texto(valor(fila, col.documento_relacionado)).upper()
    metodo = metodo_por_uuid.get(uuid) or texto(valor(fila, col.metodo_documento))
    return metodo.upper().startswith("PUE")
