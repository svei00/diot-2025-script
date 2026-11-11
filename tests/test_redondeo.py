from diot.redondeo import base_minima_para_iva, redondear_cff


def test_redondeo_cff_cincuenta_centavos_baja():
    assert redondear_cff(100.50) == 100


def test_redondeo_cff_cincuenta_y_un_centavos_sube():
    assert redondear_cff(100.51) == 101


def test_redondeo_cff_negativos_simetricos():
    assert redondear_cff(-100.51) == -101


def test_base_minima_16_caso_real_de_mayo():
    # SAT rechazó base 1254 con IVA 201: trunc(1254 * 0.16) = 200.
    assert base_minima_para_iva(1254, 201, 16) == 1257


def test_base_minima_no_baja_una_base_suficiente():
    assert base_minima_para_iva(5000, 201, 16) == 5000


def test_base_minima_8():
    # trunc(base * 0.08) >= 15  ->  base >= 187.5  ->  188
    assert base_minima_para_iva(180, 15, 8) == 188
    assert (188 * 8) // 100 >= 15
