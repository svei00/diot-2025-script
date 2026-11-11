"""Fase 4: texto de conciliación por tasa contra Resumen y Declaracion."""

from diot.conciliacion import summary
from tests.fabrica_filas import construir, factura, rep


def declaracion(mes="Octubre", anio=2026, **valores):
    return {"mes": mes, "anio": anio, "valores": valores}


def texto_de(resultado_y_renglones, **opciones):
    renglones, resultado = resultado_y_renglones
    return summary("out.txt", renglones, resultado, 2026, "RFC", "Octubre", "N", **opciones)


def test_totales_por_tasa_y_ok_contra_declaracion():
    datos = construir([factura()], [rep()])
    texto = texto_de(datos, declaracion=declaracion(
        actos_16_con_efectivo=(2000.0, None), iva_acreditable_total=(320.0, None)))
    assert "Tasa 16%" in texto and "2000.00" in texto
    assert "Actos tasa 16% (con efectivo)" in texto
    assert "DIF" not in texto


def test_diferencia_contra_declaracion_se_muestra():
    texto = texto_de(construir([factura()]),
                     declaracion=declaracion(actos_16_con_efectivo=(5000.0, None)))
    assert "DIF -4000.00" in texto


def test_conteo_con_tolerancia_cero():
    texto = texto_de(construir([factura()]), declaracion=declaracion(facturas_pue_n=(2, 2)))
    assert "DIF -1.00" in texto


def test_declaracion_de_otro_mes_no_se_concilia():
    texto = texto_de(construir([factura()]), declaracion=declaracion(mes="Julio"))
    assert "no se concilia contra este mes" in texto


def test_tasa_cero_de_facturas_se_marca_como_captura_manual():
    texto = texto_de(construir([factura()]),
                     declaracion=declaracion(actos_tasa_cero=(8991.0, 8991.0)))
    assert "capturar a mano (8991.00)" in texto
    assert "NOTA: La tasa 0% de FACTURAS" in texto
