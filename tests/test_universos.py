"""Fase 1: universos de facturas y REP, conversión de moneda."""

import pytest

from diot.layout_txt import CAMPO_IVA_ACREDITABLE_16, CAMPO_VALOR_ACTOS_16, leer_campo
from tests.fabrica_filas import construir, factura, factura_efectivo, rep


def test_factura_con_total_cero_no_cuenta():
    _, resultado = construir([factura(), factura(**{"Total": 0, "IVA 16%": 0})])
    assert resultado.buckets["A"].numero_documentos == 1


def test_factura_cancelada_y_ppd_no_entran():
    _, resultado = construir([factura(Estado="Cancelado"),
                              factura(**{"Metodo de Pago": "PPD - Pago en parcialidades"})])
    assert resultado.buckets["A"].numero_documentos == 0


def test_efectivo_dentro_del_limite_va_a_bucket_b():
    _, resultado = construir([factura_efectivo()])
    assert resultado.buckets["B"].numero_documentos == 1
    assert resultado.buckets["B"].importes.iva_16 == 160


def test_efectivo_sobre_el_limite_y_combustible_quedan_fuera():
    _, resultado = construir([factura_efectivo(Total=2500.0),
                              factura_efectivo(Combustible="Sí")])
    assert resultado.buckets["B"].numero_documentos == 0


def test_base_16_es_iva_entre_016():
    renglones, _ = construir([factura(**{"IVA 16%": 201.0})])
    assert leer_campo(renglones[0], CAMPO_IVA_ACREDITABLE_16) == 201
    # 201 / 0.16 = 1256.25 -> 1256, y la invariante del SAT lo sube a 1257
    assert leer_campo(renglones[0], CAMPO_VALOR_ACTOS_16) == 1257


def test_rep_en_dolares_se_divide_entre_equivalencia():
    # Caso real de octubre 2026 (BUCA): 2,045.50 USD a 0.0544401406 -> 37,573.38 MXN
    equivalencia = 0.0544401406
    _, resultado = construir(reps=[rep(MonedaDR="USD", EquivalenciaDR=equivalencia,
                                       **{"IVA 16 Base": 2045.50, "IVA 16 Importe": 327.28})])
    assert resultado.buckets["C"].importes.base_16 == pytest.approx(37573.38, abs=0.01)
    assert resultado.buckets["C"].importes.iva_16 == pytest.approx(327.28 / equivalencia)


def test_rep_equivalencia_vacia_o_cero_vale_uno():
    _, resultado = construir(reps=[rep(EquivalenciaDR=None), rep(EquivalenciaDR=0)])
    assert resultado.buckets["C"].importes.base_16 == 2000


def test_rep_con_pago_en_moneda_extranjera_se_reporta_y_no_suma():
    _, resultado = construir(reps=[rep(MonedaP="USD")])
    assert resultado.buckets["C"].numero_documentos == 0
    assert "MonedaP=USD" in resultado.rep_descartados[0][1]


def test_rep_sin_factura_relacionada_se_reporta():
    _, resultado = construir(reps=[rep(RelUUID="NO EXISTE")])
    assert resultado.buckets["C"].numero_documentos == 0
    assert resultado.rep_descartados[0][1] == "RelUUID=NO EXISTE"


def test_rep_duplicado_no_bancarizado_y_de_otro_mes_se_excluyen_en_silencio():
    _, resultado = construir(reps=[rep(EsDuplicado=1), rep(BancarizadoP="No"),
                                   rep(FechaPago="30/09/2026")])
    assert resultado.buckets["C"].numero_documentos == 0
    assert resultado.rep_descartados == []


def test_un_renglon_por_proveedor_ordenado_por_rfc():
    renglones, _ = construir([factura(**{"RFC Emisor": "ZZZ010101ZZZ"}), factura(), factura()])
    assert [r[2] for r in renglones] == ["AAA010101AAA", "ZZZ010101ZZZ"]
    assert leer_campo(renglones[0], CAMPO_IVA_ACREDITABLE_16) == 320
