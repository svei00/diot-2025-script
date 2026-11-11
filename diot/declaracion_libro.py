"""
Lectura de la hoja Declaracion POR ETIQUETA (para conciliar).

Declaracion se reorganiza seguido (se insertan filas), así que nunca se lee
por número de fila. Se busca el renglón cuya columna B contiene la etiqueta
y se toman:
    C = "Según registros" (lo que calcula el libro)
    D = "SAT (captura)"   (lo que muestra el prellenado, cuando se capturó)

La hoja es de UN solo mes: el de "Mes a declarar:" (C2 = nombre del mes,
D2 = año). Solo tiene sentido conciliar contra ella si es el mismo mes que
se está generando.
"""

import re

from diot.valores_celda import numero, texto

ETIQUETA_MES = "mes a declarar"

# clave -> etiqueta (se compara sin mayúsculas y con espacios normalizados)
ETIQUETAS = {
    "facturas_pue_n":        "Núm. de facturas consideradas — PUE bancarizadas",
    "rep_n":                 "Núm. de REP considerados",
    "base_16_rep":           "Base IVA 16% — REP recibidos",
    "actos_16_con_efectivo": "Monto a capturar — actos tasa 16% (con efectivo)",
    "actos_8_con_efectivo":  "Monto a capturar — región fronteriza 8% (con efectivo)",
    "actos_tasa_cero":       "Valor de los demás actos pagados a la tasa 0%",
    "actos_exentos":         "Valor de los actos pagados por los que no se pagará IVA",
    "iva_acreditable_total": "TOTAL IVA ACREDITABLE (después de proporción)",
}


def normalizar(etiqueta):
    return re.sub(r"\s+", " ", texto(etiqueta)).lower()


def renglon_con_etiqueta(rejilla, etiqueta):
    buscada = normalizar(etiqueta)
    for fila in rejilla:
        if buscada in normalizar(fila.get(1)):
            return fila
    return None


def mes_de_la_declaracion(rejilla):
    """(nombre del mes, año) de 'Mes a declarar:'; (None, None) si no está."""
    fila = renglon_con_etiqueta(rejilla, ETIQUETA_MES)
    if fila is None:
        return None, None
    anio = fila.get(3)
    return texto(fila.get(2)), (int(anio) if isinstance(anio, (int, float)) else None)


def leer_declaracion(rejilla):
    """{'mes': str, 'anio': int, 'valores': {clave: (C, D) o None}}."""
    mes, anio = mes_de_la_declaracion(rejilla)
    valores = {}
    for clave, etiqueta in ETIQUETAS.items():
        fila = renglon_con_etiqueta(rejilla, etiqueta)
        if fila is None:
            valores[clave] = None
            continue
        captura_sat = fila.get(3)
        valores[clave] = (numero(fila.get(2)),
                          numero(captura_sat) if captura_sat not in (None, "") else None)
    return {"mes": mes, "anio": anio, "valores": valores}
