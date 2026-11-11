"""
Texto de conciliación que ve el contador al terminar.

Tres bloques:
  1. Por bucket y por tasa: lo que sumó el script contra Resumen (fila del mes).
  2. Totales del TXT contra Declaracion (por etiqueta), si Declaracion es del
     mismo mes. Son las cifras que se capturan en el portal.
  3. Avisos: REP descartados, ajustes de base, facturas raras.

Tolerancias: importes $0.01 contra Resumen (mismas fórmulas, deben ser
idénticos); $1 contra Declaracion (redondeos de la propia hoja); conteos 0.
"""

from diot.layout_txt import (CAMPO_ACTOS_EXENTOS, CAMPO_ACTOS_TASA_CERO,
                             CAMPO_IVA_ACREDITABLE_8_NORTE, CAMPO_IVA_ACREDITABLE_16,
                             CAMPO_IVA_RETENIDO, CAMPO_VALOR_ACTOS_8_NORTE,
                             CAMPO_VALOR_ACTOS_16, leer_campo)

CAMPOS_A_TOTALIZAR = [
    (CAMPO_VALOR_ACTOS_16, "valor actos 16%"),
    (CAMPO_IVA_ACREDITABLE_16, "IVA acreditable 16%"),
    (CAMPO_VALOR_ACTOS_8_NORTE, "valor actos 8% frontera norte"),
    (CAMPO_IVA_ACREDITABLE_8_NORTE, "IVA acreditable 8% frontera norte"),
    (CAMPO_IVA_RETENIDO, "IVA retenido"),
    (CAMPO_ACTOS_EXENTOS, "actos exentos"),
    (CAMPO_ACTOS_TASA_CERO, "actos tasa 0%"),
]

TOLERANCIA_RESUMEN = 0.01
TOLERANCIA_DECLARACION = 1.0
ANCHO = 78

NOMBRES_BUCKET = {
    "A": "A PUE bancarizadas",
    "B": "B efectivo <= %d",
    "C": "C REP (PPD pagados)",
}

# (bucket, concepto que se muestra, atributo de ImportesPorTasa, clave en Resumen)
RENGLONES_POR_TASA = [
    ("A", "base 16%", "base_16", None),
    ("A", "IVA 16%", "iva_16", "A_iva_16"),
    ("B", "base 16%", "base_16", None),
    ("B", "IVA 16%", "iva_16", "B_iva_16"),
    ("C", "base 16%", "base_16", None),
    ("C", "IVA 16%", "iva_16", "C_iva_16"),
    ("A", "base 8%", "base_8", "A_base_8"),
    ("A", "IVA 8%", "iva_8", "A_iva_8"),
    ("B", "base 8%", "base_8", "B_base_8"),
    ("B", "IVA 8%", "iva_8", "B_iva_8"),
    ("C", "base 8%", "base_8", "C_base_8"),
    ("C", "IVA 8%", "iva_8", "C_iva_8"),
    ("A", "exento", "exento", "A_exento"),
    ("C", "exento", "exento", "C_exento"),
    ("C", "tasa 0%", "tasa_cero", "C_tasa_cero"),
    ("A", "IVA ret.", "iva_retenido", None),
    ("B", "IVA ret.", "iva_retenido", None),
    ("C", "IVA ret.", "iva_retenido", None),
]

# (concepto, clave en Declaracion, función que calcula nuestra cifra)
CRUCES_DECLARACION = [
    ("Núm. facturas PUE bancarizadas", "facturas_pue_n",
     lambda r: r.buckets["A"].numero_documentos),
    ("Núm. REP pagados en el mes", "rep_n",
     lambda r: r.buckets["C"].numero_documentos),
    ("Base 16% REP (convertida)", "base_16_rep",
     lambda r: r.buckets["C"].importes.base_16),
    ("Actos tasa 16% (con efectivo)", "actos_16_con_efectivo",
     lambda r: sum(b.importes.base_16 for b in r.buckets.values())),
    ("Actos frontera 8% (con efectivo)", "actos_8_con_efectivo",
     lambda r: sum(b.importes.base_8 for b in r.buckets.values())),
    ("Actos exentos (facturas + REP)", "actos_exentos",
     lambda r: sum(b.importes.exento for b in r.buckets.values())),
    ("Actos tasa 0% (solo REP)", "actos_tasa_cero",
     lambda r: sum(b.importes.tasa_cero for b in r.buckets.values())),
    ("TOTAL IVA acreditable", "iva_acreditable_total",
     lambda r: sum(b.importes.iva_16 + b.importes.iva_8 for b in r.buckets.values())),
]

CONTEOS = {"facturas_pue_n", "rep_n"}

# Renglones que el script no puede derivar completos: la diferencia se captura a mano.
CAPTURA_MANUAL = {"actos_tasa_cero": "facturas 0%: capturar a mano"}

NOTA_TASA_CERO = ("La tasa 0% de FACTURAS no se deriva del export (conceptos a 0% dentro de "
                  "facturas mixtas). El campo 51 solo trae la de REP; la de facturas se "
                  "captura a mano en el portal.")


def marca(diferencia, tolerancia):
    return "OK" if abs(diferencia) < tolerancia + 1e-9 else "DIF %.2f" % diferencia


def encabezado(rfc, mes_largo, anio, tipo, renglones, ruta):
    return ["=" * ANCHO,
            " DIOT %s  -  %s %d   tipo: %s" % (rfc, mes_largo, anio, tipo),
            "=" * ANCHO,
            " Proveedores en TXT:     %d" % len(renglones),
            " Archivo:                %s" % ruta]


def bloque_por_tasa(resultado, resumen, limite):
    lineas = ["-" * ANCHO,
              " %-22s %-10s %16s %16s   %s" % ("Bucket", "Concepto", "Script", "Resumen", "")]
    for bucket, concepto, atributo, clave in RENGLONES_POR_TASA:
        nuestro = getattr(resultado.buckets[bucket].importes, atributo)
        esperado = resumen.get(clave) if (resumen and clave) else None
        estado = marca(nuestro - esperado, TOLERANCIA_RESUMEN) if esperado is not None else ""
        lineas.append(" %-22s %-10s %16.2f %16s   %s" % (
            NOMBRES_BUCKET[bucket].replace("%d", str(int(limite))), concepto, nuestro,
            "%.2f" % esperado if esperado is not None else "-", estado))
    for bucket in "ABC":
        nuestro = resultado.buckets[bucket].numero_documentos
        esperado = resumen.get("%s_n" % bucket) if resumen else None
        estado = marca(nuestro - esperado, 0) if esperado is not None else ""
        lineas.append(" %-22s %-10s %16d %16s   %s" % (
            NOMBRES_BUCKET[bucket].replace("%d", str(int(limite))), "documentos", nuestro,
            "%d" % esperado if esperado is not None else "-", estado))
    return lineas


def bloque_totales_txt(renglones):
    """Suma de cada campo del TXT (lo que verá el portal al cargar el archivo)."""
    lineas = ["-" * ANCHO, " TXT en pesos (Art. 20 CFF), suma por campo:"]
    for numero_campo, nombre in CAMPOS_A_TOTALIZAR:
        total = sum(leer_campo(r, numero_campo) for r in renglones)
        if total or numero_campo in (CAMPO_VALOR_ACTOS_16, CAMPO_IVA_ACREDITABLE_16):
            lineas.append("   campo %-3d %-40s %14d" % (numero_campo, nombre, total))
    return lineas


def bloque_declaracion(resultado, declaracion, mes_largo, anio):
    if not declaracion:
        return []
    if (declaracion["mes"] or "").lower() != mes_largo.lower() or declaracion["anio"] != anio:
        return ["-" * ANCHO,
                " Declaracion es de %s %s: no se concilia contra este mes."
                % (declaracion["mes"], declaracion["anio"])]
    lineas = ["-" * ANCHO,
              " %-34s %16s %16s   %s" % ("vs Declaracion (por etiqueta)", "Script", "Declaracion", "")]
    for concepto, clave, calcular in CRUCES_DECLARACION:
        nuestro = calcular(resultado)
        celdas = declaracion["valores"].get(clave)
        if celdas is None:
            lineas.append(" %-34s %16.2f %16s   etiqueta no encontrada" % (concepto, nuestro, "-"))
            continue
        tolerancia = 0 if clave in CONTEOS else TOLERANCIA_DECLARACION
        estado = marca(nuestro - celdas[0], tolerancia)
        if estado != "OK" and clave in CAPTURA_MANUAL:
            estado = "%s (%.2f)" % (CAPTURA_MANUAL[clave], celdas[0] - nuestro)
        lineas.append(" %-34s %16.2f %16.2f   %s" % (concepto, nuestro, celdas[0], estado))
    return lineas


def bloque_avisos(resultado):
    lineas = []
    if resultado.ajuste_base_16_pesos:
        lineas.append(" Ajuste base 16%% invariante SAT  +%d peso(s)  (trunc(base*0.16) >= IVA)"
                      % resultado.ajuste_base_16_pesos)
    if resultado.ajuste_base_8_pesos:
        lineas.append(" Ajuste base 8%% invariante SAT   +%d peso(s)  (trunc(base*0.08) >= IVA)"
                      % resultado.ajuste_base_8_pesos)
    for rfc, motivo, iva in resultado.rep_descartados:
        lineas.append(" REP descartado: %-14s %-28s IVA %.2f" % (rfc, motivo, iva))
    for aviso in resultado.avisos:
        lineas.append(" AVISO: %s" % aviso)
    lineas.append(" NOTA: %s" % NOTA_TASA_CERO)
    return ["-" * ANCHO] + lineas


def summary(ruta, renglones, resultado, anio, rfc, mes_largo, tipo,
            resumen=None, declaracion=None, limite=2000):
    lineas = encabezado(rfc, mes_largo, anio, tipo, renglones, ruta)
    lineas += bloque_por_tasa(resultado, resumen, limite)
    lineas += bloque_totales_txt(renglones)
    lineas += bloque_declaracion(resultado, declaracion, mes_largo, anio)
    lineas += bloque_avisos(resultado)
    lineas.append("=" * ANCHO)
    return "\n".join(lineas)
