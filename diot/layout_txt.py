"""
Layout del TXT de carga masiva de la DIOT (ejercicios 2025 en adelante).

Fuente: "Instructivo para el armado del archivo de carga masiva - DIOT",
SAT, enero 2025, sección 3. Los números son los del SAT (empiezan en 1).
54 campos separados por "|", un renglón por proveedor, UTF-8, fin de línea
CRLF. Los campos sin dato van vacíos pero conservan su "|".

Solo se listan los campos que este generador llena:

    1   Tipo de tercero (04 = proveedor nacional)
    2   Tipo de operación (85 otros, 03 servicios profesionales, 06 arrendamiento)
    3   RFC del proveedor
    8   Valor de actos pagados — Región Fronteriza NORTE 8%
    12  Valor de actos pagados — tasa general 16%
    18  IVA acreditable exclusivamente de actividades gravadas — RF NORTE 8%
    22  IVA acreditable exclusivamente de actividades gravadas — 16%
    48  IVA retenido por el contribuyente
    50  Actos pagados por los que no se pagará el IVA (exentos)
    51  Demás actos pagados a la tasa del 0%
"""

from dataclasses import dataclass

TOTAL_DE_CAMPOS = 54

CAMPO_TIPO_TERCERO = 1
CAMPO_TIPO_OPERACION = 2
CAMPO_RFC = 3
CAMPO_VALOR_ACTOS_8_NORTE = 8
CAMPO_VALOR_ACTOS_16 = 12
CAMPO_IVA_ACREDITABLE_8_NORTE = 18
CAMPO_IVA_ACREDITABLE_16 = 22
CAMPO_IVA_RETENIDO = 48
CAMPO_ACTOS_EXENTOS = 50
CAMPO_ACTOS_TASA_CERO = 51

TIPO_TERCERO_NACIONAL = "04"


@dataclass
class ImportesEnteros:
    """Importes de un proveedor ya redondeados a pesos (Art. 20 CFF)."""
    base_16: int = 0
    iva_16: int = 0
    base_8: int = 0
    iva_8: int = 0
    exento: int = 0
    tasa_cero: int = 0
    iva_retenido: int = 0

    def esta_vacio(self):
        return not any((self.base_16, self.iva_16, self.base_8, self.iva_8,
                        self.exento, self.tasa_cero, self.iva_retenido))


def poner_campo(renglon, numero_de_campo, valor):
    """Escribe `valor` en el campo con número del SAT (1..54)."""
    renglon[numero_de_campo - 1] = str(valor)


def poner_si_no_es_cero(renglon, numero_de_campo, valor):
    if valor:
        poner_campo(renglon, numero_de_campo, valor)


def armar_renglon(rfc, tipo_operacion, enteros):
    """Lista de 54 textos para un proveedor nacional."""
    renglon = [""] * TOTAL_DE_CAMPOS
    poner_campo(renglon, CAMPO_TIPO_TERCERO, TIPO_TERCERO_NACIONAL)
    poner_campo(renglon, CAMPO_TIPO_OPERACION, tipo_operacion)
    poner_campo(renglon, CAMPO_RFC, rfc)
    if enteros.base_16 or enteros.iva_16:
        poner_campo(renglon, CAMPO_VALOR_ACTOS_16, enteros.base_16)
        poner_campo(renglon, CAMPO_IVA_ACREDITABLE_16, enteros.iva_16)
    if enteros.base_8 or enteros.iva_8:
        poner_campo(renglon, CAMPO_VALOR_ACTOS_8_NORTE, enteros.base_8)
        poner_campo(renglon, CAMPO_IVA_ACREDITABLE_8_NORTE, enteros.iva_8)
    poner_si_no_es_cero(renglon, CAMPO_IVA_RETENIDO, enteros.iva_retenido)
    poner_si_no_es_cero(renglon, CAMPO_ACTOS_EXENTOS, enteros.exento)
    poner_si_no_es_cero(renglon, CAMPO_ACTOS_TASA_CERO, enteros.tasa_cero)
    return renglon


def leer_campo(renglon, numero_de_campo):
    """Valor entero de un campo del renglón (vacío -> 0)."""
    contenido = renglon[numero_de_campo - 1]
    return int(contenido) if contenido else 0
