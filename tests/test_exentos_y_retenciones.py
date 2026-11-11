"""Fase 3: exentos, tasa 0%, retenciones de REP y REP sobre factura PUE."""

import pytest

from diot.layout_txt import (CAMPO_ACTOS_EXENTOS, CAMPO_ACTOS_TASA_CERO,
                             CAMPO_IVA_RETENIDO, leer_campo)
from tests.fabrica_filas import ENCABEZADOS_FACTURAS, construir, factura, factura_efectivo, rep


def colegiatura(**cambios):
    datos = {"IVA 16%": 0.0, "Total Trasladados": 0.0, "SubTotal": 5000.0, "Total": 5000.0,
             "Nivel Educativo": "Primaria", "RFC Emisor": "ESC010101AAA"}
    datos.update(cambios)
    return factura(**datos)


def test_colegiatura_bancarizada_va_al_campo_50():
    renglones, resultado = construir([colegiatura()])
    assert resultado.buckets["A"].importes.exento == 5000
    assert leer_campo(renglones[0], CAMPO_ACTOS_EXENTOS) == 5000


def test_concepto_alumno_tambien_es_exento():
    _, resultado = construir([colegiatura(**{"Nivel Educativo": None,
                                             "Conceptos": "Inscripción Alumno 2026"})])
    assert resultado.buckets["A"].importes.exento == 5000


def test_factura_sin_impuestos_que_no_es_colegiatura_no_es_exenta():
    # Cuotas IMSS / "no objeto": sin traslados, pero no son exentas ni tasa 0%.
    _, resultado = construir([colegiatura(**{"Nivel Educativo": None, "Conceptos": "Cuotas IMSS"})])
    assert resultado.buckets["A"].importes.exento == 0


def test_colegiatura_en_efectivo_no_suma_exento():
    _, resultado = construir([factura_efectivo(**{"IVA 16%": 0.0, "Total Trasladados": 0.0,
                                                  "SubTotal": 900.0, "Total": 900.0,
                                                  "Nivel Educativo": "Primaria"})])
    assert resultado.buckets["B"].importes.exento == 0


def test_rep_exento_y_tasa_cero_se_convierten():
    _, resultado = construir(reps=[rep(MonedaDR="USD", EquivalenciaDR=0.05,
                                       **{"IVA 16 Base": 0, "IVA 16 Importe": 0,
                                          "IVA Excento Base": 10.0, "IVA Cero Base": 5.0})])
    renglones, _ = construir(reps=[rep(**{"IVA Excento Base": 10.0, "IVA Cero Base": 5.0})])
    assert resultado.buckets["C"].importes.exento == pytest.approx(200.0)
    assert resultado.buckets["C"].importes.tasa_cero == pytest.approx(100.0)
    assert leer_campo(renglones[0], CAMPO_ACTOS_TASA_CERO) == 5


def test_retencion_iva_del_rep_va_al_campo_48_y_avisa():
    renglones, resultado = construir(reps=[rep(**{"Ret IVA 10.667 Importe": 106.67})])
    assert leer_campo(renglones[0], CAMPO_IVA_RETENIDO) == 107
    assert any("IVA retenido" in aviso for aviso in resultado.avisos)


def test_rep_sobre_factura_pue_cuenta_dos_veces_y_avisa():
    encabezados = ENCABEZADOS_FACTURAS + ["UUID"]
    factura_pue = colegiatura(UUID="ABC-1", SubTotal=520.0, Total=520.0)
    rep_de_la_pue = rep(**{"RFC Emisor CFDI": "ESC010101AAA", "IdDocumento Relacionado": "abc-1",
                           "MetodoDePagoDR": None, "IVA 16 Base": 0, "IVA 16 Importe": 0,
                           "IVA Excento Base": 520.0, "ImpPagado": 520.0})
    from diot.construir import build_diot
    from tests.fabrica_filas import ANIO, ENCABEZADOS_REP, MES
    renglones, resultado = build_diot(encabezados, [factura_pue], ENCABEZADOS_REP,
                                      [rep_de_la_pue], ANIO, MES, {})
    assert leer_campo(renglones[0], CAMPO_ACTOS_EXENTOS) == 1040     # igual que el SAT
    assert any("factura PUE" in aviso for aviso in resultado.avisos)
