"""
build_diot(): de las tablas del libro a los renglones del TXT.

Es una función PURA: recibe filas ya leídas y regresa datos. No abre
archivos, no usa tkinter, no imprime. Por eso se puede llamar desde
Boardflare (Python en Excel) o desde las pruebas con filas sintéticas.

Pasos:
  1. Recorre las facturas: universo del mes -> bucket A o B -> importes.
  2. Recorre los REP: universo del mes -> importes convertidos a MXN.
  3. Suma por proveedor (RFC) y por bucket.
  4. Redondea cada proveedor a pesos y aplica la invariante del SAT.
  5. Arma un renglón de 54 campos por proveedor, ordenados por RFC.
"""

from collections import defaultdict
from dataclasses import dataclass, field
from typing import Dict, List, Tuple

from diot import facturas, rep
from diot.acumulados import ImportesPorTasa, TotalesDeBucket
from diot.columnas import columnas_de_facturas, columnas_de_rep
from diot.layout_txt import ImportesEnteros, armar_renglon
from diot.redondeo import base_minima_para_iva, redondear_cff
from diot.valores_celda import numero

USOS_DEDUCIBLES_POR_OMISION = {"G01", "G03"}
LIMITE_EFECTIVO_POR_OMISION = 2000
TIPO_OPERACION_POR_OMISION = "85"            # 85 = Otros


@dataclass
class ResultadoDiot:
    """Todo lo que build_diot() averiguó, además de los renglones."""
    buckets: Dict[str, TotalesDeBucket] = field(default_factory=lambda: {
        "A": TotalesDeBucket(), "B": TotalesDeBucket(), "C": TotalesDeBucket()})
    ajuste_base_16_pesos: int = 0
    ajuste_base_8_pesos: int = 0
    rep_descartados: List[Tuple[str, str, float]] = field(default_factory=list)
    avisos: List[str] = field(default_factory=list)


# ---------------------------------------------------------------- recorrer el libro
def acumular_facturas(filas, col, anio, mes, usos, limite, por_proveedor, resultado):
    for fila in filas:
        if not facturas.es_factura_deducible_del_mes(fila, col, anio, mes, usos):
            continue
        bucket = facturas.bucket_de_la_factura(fila, col, limite)
        if bucket is None:
            continue
        importes = facturas.importes_de_factura(fila, col, bucket)
        por_proveedor[fila.get(col.rfc)].sumar(importes)
        resultado.buckets[bucket].agregar(importes)


def acumular_rep(filas, col, anio, mes, por_proveedor, resultado):
    for fila in filas:
        se_incluye, motivo = rep.motivo_para_excluir_rep(fila, col, anio, mes)
        if not se_incluye:
            if motivo:
                resultado.rep_descartados.append(
                    (fila.get(col.rfc), motivo, numero(fila.get(col.iva_16))))
            continue
        importes = rep.importes_de_rep(fila, col)
        por_proveedor[fila.get(col.rfc)].sumar(importes)
        resultado.buckets["C"].agregar(importes)


# ---------------------------------------------------------------- redondeo por proveedor
def redondear_proveedor(importes, resultado):
    """Pesos enteros (Art. 20 CFF) + base mínima que exige el validador del SAT."""
    enteros = ImportesEnteros(
        base_16=redondear_cff(importes.base_16), iva_16=redondear_cff(importes.iva_16),
        base_8=redondear_cff(importes.base_8), iva_8=redondear_cff(importes.iva_8),
        exento=redondear_cff(importes.exento), tasa_cero=redondear_cff(importes.tasa_cero),
        iva_retenido=redondear_cff(importes.iva_retenido))
    base_16_ajustada = base_minima_para_iva(enteros.base_16, enteros.iva_16, 16)
    base_8_ajustada = base_minima_para_iva(enteros.base_8, enteros.iva_8, 8)
    resultado.ajuste_base_16_pesos += base_16_ajustada - enteros.base_16
    resultado.ajuste_base_8_pesos += base_8_ajustada - enteros.base_8
    enteros.base_16, enteros.base_8 = base_16_ajustada, base_8_ajustada
    return enteros


def armar_renglones(por_proveedor, tipos_operacion, resultado):
    renglones = []
    for rfc in sorted(por_proveedor):
        enteros = redondear_proveedor(por_proveedor[rfc], resultado)
        if enteros.esta_vacio():
            continue
        tipo_operacion = tipos_operacion.get(rfc, TIPO_OPERACION_POR_OMISION)
        renglones.append(armar_renglon(rfc, tipo_operacion, enteros))
    return renglones


# ---------------------------------------------------------------- API pública
def build_diot(rec_headers, rec_rows, pag_headers, pag_rows, year, month, op_map,
               limite_efe=LIMITE_EFECTIVO_POR_OMISION, usos_ded=None):
    """Regresa (renglones, ResultadoDiot).

    rec_headers/rec_rows: RecibidasXML (read_sheet).
    pag_headers/pag_rows: PagosRecibidasXML (read_sheet).
    op_map: {RFC: tipo de operación} (diot_rfc_op.csv); por omisión "85".
    """
    usos = usos_ded or USOS_DEDUCIBLES_POR_OMISION
    resultado = ResultadoDiot()
    por_proveedor = defaultdict(ImportesPorTasa)

    acumular_facturas(rec_rows, columnas_de_facturas(rec_headers), year, month,
                      usos, limite_efe, por_proveedor, resultado)
    acumular_rep(pag_rows, columnas_de_rep(pag_headers), year, month,
                 por_proveedor, resultado)

    renglones = armar_renglones(por_proveedor, op_map or {}, resultado)
    return renglones, resultado
