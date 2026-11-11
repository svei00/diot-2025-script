"""Fase 2: IVA 8% (estímulo región fronteriza norte)."""

import pytest

from diot.layout_txt import (CAMPO_IVA_ACREDITABLE_8_NORTE, CAMPO_IVA_ACREDITABLE_16,
                             CAMPO_VALOR_ACTOS_8_NORTE, CAMPO_VALOR_ACTOS_16, leer_campo)
from tests.fabrica_filas import construir, factura, factura_efectivo, rep


def factura_8(**cambios):
    datos = {"IVA 16%": 0.0, "IVA 8%": 80.0, "SubTotal": 1000.0, "Total": 1080.0,
             "Total Trasladados": 80.0, "LugarDeExpedicion": "21397"}
    datos.update(cambios)
    return factura(**datos)


def test_bancarizada_8_toma_subtotal_menos_descuento():
    _, resultado = construir([factura_8(Descuento=10.0, **{"IVA 8%": 79.2})])
    assert resultado.buckets["A"].importes.base_8 == pytest.approx(990.0)
    assert resultado.buckets["A"].importes.iva_8 == pytest.approx(79.2)


def test_efectivo_8_con_ieps_usa_iva_entre_008_no_el_neto():
    # Neto 1,050 trae IEPS; el IVA 8% es 80 -> base 1,000.
    fila = factura_efectivo(**{"IVA 16%": 0.0, "IVA 8%": 80.0, "SubTotal": 1050.0,
                               "Total": 1130.0})
    _, resultado = construir([fila])
    assert resultado.buckets["B"].importes.base_8 == pytest.approx(1000.0)


def test_rep_8_en_dolares_se_convierte():
    _, resultado = construir(reps=[rep(MonedaDR="USD", EquivalenciaDR=0.05,
                                       **{"IVA 16 Base": 0, "IVA 16 Importe": 0,
                                          "IVA 8 Base": 100.0, "IVA 8 Importe": 8.0})])
    assert resultado.buckets["C"].importes.base_8 == pytest.approx(2000.0)
    assert resultado.buckets["C"].importes.iva_8 == pytest.approx(160.0)


def test_campos_8_y_18_y_16_vacios_si_solo_hay_8():
    renglones, _ = construir([factura_8()])
    renglon = renglones[0]
    assert leer_campo(renglon, CAMPO_VALOR_ACTOS_8_NORTE) == 1000
    assert leer_campo(renglon, CAMPO_IVA_ACREDITABLE_8_NORTE) == 80
    assert renglon[CAMPO_VALOR_ACTOS_16 - 1] == ""
    assert renglon[CAMPO_IVA_ACREDITABLE_16 - 1] == ""


def test_invariante_8_sube_la_base():
    # SubTotal 187 con IVA 15: trunc(187 * 0.08) = 14 < 15 -> base 188
    renglones, resultado = construir([factura_8(SubTotal=187.0, **{"IVA 8%": 15.0})])
    assert leer_campo(renglones[0], CAMPO_VALOR_ACTOS_8_NORTE) == 188
    assert resultado.ajuste_base_8_pesos == 1


def test_factura_mixta_16_y_8_avisa():
    _, resultado = construir([factura(**{"IVA 8%": 8.0})])
    assert any("IVA 16% e IVA 8%" in aviso for aviso in resultado.avisos)
    assert resultado.buckets["A"].importes.base_8 == pytest.approx(100.0)   # 8 / 0.08


def test_cp_de_frontera_sur_avisa():
    _, resultado = construir([factura_8(LugarDeExpedicion="30700")])
    assert any("frontera SUR" in aviso for aviso in resultado.avisos)


def test_sin_iva_8_el_renglon_16_no_cambia():
    renglones, resultado = construir([factura()])
    assert renglones[0][CAMPO_VALOR_ACTOS_8_NORTE - 1] == ""
    assert resultado.avisos == []
