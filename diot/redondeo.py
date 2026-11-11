"""
Redondeo de importes para el TXT.

1. Art. 20, párrafo 7, CFF: los importes se expresan en pesos enteros;
   0.01 a 0.50 baja, 0.51 a 0.99 sube.

2. Invariante del validador del SAT. Al cargar el TXT, el SAT recalcula el
   "IVA pagado" de cada renglón a partir de la base y TRUNCA (no redondea).
   Si el IVA acreditable que declaramos es mayor, rechaza el renglón. Pasó en
   mayo 2026: base 1254, IVA 201 -> trunc(1254 x 0.16) = 200 < 201 -> error.

   Por eso subimos la base al mínimo entero que cumple

        (base x tasa_por_ciento) // 100  >=  iva

   Todo en aritmética entera: 1257 * 0.16 en float da 201.12000000000003.
   La misma regla vale para la tasa 16% y para la tasa 8% (región fronteriza).
"""


def redondear_cff(importe):
    """Redondeo a pesos enteros según el Art. 20 CFF."""
    if importe is None:
        return 0
    if importe < 0:
        return -redondear_cff(-importe)
    pesos = int(importe)
    centavos = int(round((importe - pesos) * 100))
    return pesos + 1 if centavos >= 51 else pesos


def base_minima_para_iva(base_entera, iva_entero, tasa_por_ciento):
    """Menor base entera >= base_entera tal que trunc(base x tasa) >= iva.

    trunc(base * t / 100) >= iva   <=>   base >= ceil(iva * 100 / t)

    Ejemplo tasa 16: iva 201 -> ceil(20100 / 16) = 1257.
    """
    if iva_entero <= 0:
        return base_entera
    base_que_exige_el_sat = -(-iva_entero * 100 // tasa_por_ciento)   # ceil entero
    return max(base_entera, base_que_exige_el_sat)
