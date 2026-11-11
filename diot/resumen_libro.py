"""
Lectura de la hoja Resumen POR ENCABEZADO (para conciliar).

Resumen tiene dos filas de encabezado:
  * fila 1: título del bloque, en la primera columna del bloque (celdas
    combinadas), p. ej. "Compras PUE bancarizadas (G01+G03, vigentes)";
  * fila 2: subtítulo de cada columna, p. ej. "Facturas", "IVA 16%".

"IVA 16%" aparece en varios bloques, así que una columna se identifica con
el par (bloque, subtítulo). Las columnas de los bloques nuevos (8%, exentos,
tasa 0%) tienen subtítulos únicos y se buscan solo por subtítulo.

Desde noviembre 2026 Resumen!Y ya NO es la base de IVA de los REP (ahora es
"Pagado s/IVA" por proporción). Por eso este módulo no lee la base de REP de
Resumen: la base 16% de REP se concilia contra Declaracion.

Si falta una columna obligatoria, se detiene con un mensaje claro. Las
opcionales (bloques nuevos, ausentes en libros viejos) quedan en None.
"""

from diot.valores_celda import numero, texto

# clave -> (texto contenido en el título del bloque, subtítulo exacto)
COLUMNAS_OBLIGATORIAS = {
    "A_n":      ("Compras PUE bancarizadas", "Facturas"),
    "A_iva_16": ("Compras PUE bancarizadas", "IVA 16%"),
    "B_n":      ("Compras en efectivo", "Facturas"),
    "B_iva_16": ("Compras en efectivo", "IVA 16%"),
    "C_n":      ("REP recibidos (PPD", "REPs"),
    "C_iva_16": ("REP recibidos (PPD", "IVA 16%"),
}

# clave -> subtítulo exacto (únicos en la fila 2)
COLUMNAS_OPCIONALES = {
    "A_base_8": "Base 8% facturas",
    "A_iva_8":  "IVA 8% facturas",
    "C_base_8": "Base 8% REP",
    "C_iva_8":  "IVA 8% REP",
    "B_base_8": "Base 8% efectivo",
    "B_iva_8":  "IVA 8% efectivo",
    "A_exento": "Exento facturas",
    "C_exento": "Exento REP",
    "C_tasa_cero": "Base 0% REP",
}


class ColumnaDeResumenNoEncontrada(KeyError):
    pass


def titulos_de_bloque_por_columna(fila_titulos, ultima_columna):
    """Extiende cada título de bloque hacia la derecha hasta el siguiente título."""
    titulos, titulo_vigente = {}, ""
    for indice in range(ultima_columna + 1):
        if texto(fila_titulos.get(indice)):
            titulo_vigente = texto(fila_titulos.get(indice))
        titulos[indice] = titulo_vigente
    return titulos


def ubicar_columna_obligatoria(clave, fila_subtitulos, titulos):
    texto_bloque, subtitulo = COLUMNAS_OBLIGATORIAS[clave]
    for indice, valor in fila_subtitulos.items():
        if (texto(valor) == subtitulo
                and texto_bloque.lower() in titulos.get(indice, "").lower()):
            return indice
    raise ColumnaDeResumenNoEncontrada(
        "Resumen: no encontré la columna '%s' dentro del bloque '%s…'. "
        "Revisa los encabezados de las filas 1 y 2 de Resumen."
        % (subtitulo, texto_bloque))


def ubicar_columna_opcional(subtitulo, fila_subtitulos):
    for indice, valor in fila_subtitulos.items():
        if texto(valor) == subtitulo:
            return indice
    return None


def ubicar_columnas(fila_titulos, fila_subtitulos):
    """{clave: índice de columna o None}."""
    ultima = max(list(fila_titulos) + list(fila_subtitulos) + [0])
    titulos = titulos_de_bloque_por_columna(fila_titulos, ultima)
    indices = {clave: ubicar_columna_obligatoria(clave, fila_subtitulos, titulos)
               for clave in COLUMNAS_OBLIGATORIAS}
    for clave, subtitulo in COLUMNAS_OPCIONALES.items():
        indices[clave] = ubicar_columna_opcional(subtitulo, fila_subtitulos)
    return indices


def fila_del_mes(rejilla, mes):
    """La fila de datos cuya columna A es el número de mes."""
    for fila in rejilla[2:]:
        if isinstance(fila.get(0), (int, float)) and int(fila[0]) == mes:
            return fila
    return None


def leer_resumen(rejilla, mes):
    """{clave: valor esperado} de la fila del mes; None si no hay fila del mes."""
    if len(rejilla) < 3:
        return None
    indices = ubicar_columnas(rejilla[0], rejilla[1])
    fila = fila_del_mes(rejilla, mes)
    if fila is None:
        return None
    return {clave: (numero(fila.get(indice)) if indice is not None else None)
            for clave, indice in indices.items()}
