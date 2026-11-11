"""Resumen por encabezado y Declaracion por etiqueta (rejillas sintéticas)."""

import pytest

from diot.declaracion_libro import leer_declaracion
from diot.resumen_libro import ColumnaDeResumenNoEncontrada, leer_resumen


def rejilla_resumen(con_bloques_nuevos=True):
    titulos = {14: "Compras PUE bancarizadas (G01+G03, vigentes)",
               18: "Compras en efectivo ≤ límite (deducibles)",
               23: "REP recibidos (PPD pagados en el mes)"}
    subtitulos = {14: "Facturas", 17: "IVA 16%", 18: "Facturas", 22: "IVA 16%",
                  23: "REPs", 24: "Pagado s/IVA", 25: "IVA 16%"}
    octubre = {0: 10, 1: "Octubre", 14: 405, 17: 685197.03, 18: 74, 22: 6256.12,
               23: 17, 24: 358456.33, 25: 29678.97}
    if con_bloques_nuevos:
        titulos[39] = "IVA 8% — estímulo región fronteriza"
        subtitulos.update({39: "Base 8% facturas", 40: "IVA 8% facturas"})
        octubre.update({39: 192641.81, 40: 15411.31})
    return [titulos, subtitulos, octubre]


def test_resumen_se_lee_por_bloque_y_subtitulo():
    valores = leer_resumen(rejilla_resumen(), 10)
    assert valores["A_n"] == 405
    assert valores["B_iva_16"] == 6256.12
    assert valores["C_iva_16"] == 29678.97        # no confunde con "Pagado s/IVA"
    assert valores["A_base_8"] == 192641.81


def test_resumen_sin_bloques_nuevos_deja_opcionales_en_none():
    valores = leer_resumen(rejilla_resumen(con_bloques_nuevos=False), 10)
    assert valores["A_base_8"] is None


def test_resumen_sin_columna_obligatoria_falla_con_mensaje_claro():
    rejilla = rejilla_resumen()
    del rejilla[1][25]                              # quitamos "IVA 16%" de REP
    with pytest.raises(ColumnaDeResumenNoEncontrada, match="REP recibidos"):
        leer_resumen(rejilla, 10)


def test_declaracion_por_etiqueta_aunque_se_muevan_las_filas():
    rejilla = [
        {1: "Mes a declarar:", 2: "Octubre", 3: 2026},
        {},
        {1: "      Monto a capturar — actos tasa 16% (con efectivo)", 2: 4507075.69},
        {1: "TOTAL IVA ACREDITABLE (después de proporción)", 2: 750343.12, 3: 750343},
    ]
    declaracion = leer_declaracion(rejilla)
    assert (declaracion["mes"], declaracion["anio"]) == ("Octubre", 2026)
    assert declaracion["valores"]["actos_16_con_efectivo"] == (4507075.69, None)
    assert declaracion["valores"]["iva_acreditable_total"] == (750343.12, 750343)
    assert declaracion["valores"]["actos_8_con_efectivo"] is None
