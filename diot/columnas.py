"""
Ubicación de columnas por ENCABEZADO.

El libro cambia de versión en versión (v5 quitó "Importe Neto", agregó
"IVA 8%", etc.). Nunca leemos por letra de columna: buscamos el encabezado.

Dos niveles de búsqueda:
  * buscar_columna_obligatoria(): para las columnas de siempre. Tolera
    variaciones (mayúsculas, nombre antiguo, nombre contenido en otro). Si no
    la encuentra, se detiene con un error claro.
  * buscar_columna_opcional(): para columnas nuevas. SOLO coincidencia exacta
    (sin importar mayúsculas). Si no existe regresa None y el importe vale 0.
    No usa coincidencia parcial a propósito: es preferible "no hay columna" a
    leer en silencio una columna equivocada.
"""

from dataclasses import dataclass, field
from typing import List, Optional


def buscar_columna_obligatoria(encabezados, *candidatos):
    """Primer encabezado que coincida con algún candidato; error si ninguno."""
    for candidato in candidatos:                                   # exacto
        if candidato in encabezados:
            return candidato
    por_minusculas = {e.lower(): e for e in encabezados}
    for candidato in candidatos:                                   # sin mayúsculas
        if candidato.lower() in por_minusculas:
            return por_minusculas[candidato.lower()]
    for candidato in candidatos:                                   # contenido
        for encabezado in encabezados:
            if candidato.lower() in encabezado.lower():
                return encabezado
    raise KeyError("columna no encontrada: probé %s" % (candidatos,))


def buscar_columna_opcional(encabezados, nombre):
    """Encabezado idéntico a `nombre` (sin importar mayúsculas) o None."""
    for encabezado in encabezados:
        if encabezado.lower() == nombre.lower():
            return encabezado
    return None


def columnas_de_retencion_iva(encabezados):
    """Todas las columnas 'Ret IVA <tasa> Importe' del REP (4, 10.667, 2, 16...)."""
    return [e for e in encabezados
            if e.lower().startswith("ret iva ") and e.lower().endswith(" importe")]


# ---------------------------------------------------------------- RecibidasXML
@dataclass
class ColumnasFacturas:
    fecha: str
    tipo: str
    estado: str
    metodo_pago: str
    bancarizado: str
    uso_cfdi: str
    iva_16: str
    forma_pago: str
    combustible: str
    rfc: str
    retenido_iva: str
    total: str
    subtotal: Optional[str] = None
    descuento: Optional[str] = None
    iva_8: Optional[str] = None
    total_trasladados: Optional[str] = None
    nivel_educativo: Optional[str] = None
    conceptos: Optional[str] = None
    lugar_expedicion: Optional[str] = None


def columnas_de_facturas(encabezados):
    obligatoria = lambda *c: buscar_columna_obligatoria(encabezados, *c)
    opcional = lambda nombre: buscar_columna_opcional(encabezados, nombre)
    return ColumnasFacturas(
        fecha=obligatoria("Fecha Emision", "Fecha"),
        tipo=obligatoria("Tipo"),
        estado=obligatoria("Estado"),
        metodo_pago=obligatoria("Metodo de Pago", "Metodo"),
        bancarizado=obligatoria("Bancarizado"),
        uso_cfdi=obligatoria("UsoCFDI"),
        iva_16=obligatoria("IVA 16%"),
        forma_pago=obligatoria("FormaDePago"),
        combustible=obligatoria("Combustible"),
        rfc=obligatoria("RFC Emisor", "RFC"),
        retenido_iva=obligatoria("Retenido IVA"),
        # v5 quitó "Importe Neto"; el límite de efectivo es IVA incluido => Total.
        total=obligatoria("Total", "Importe Neto"),
        subtotal=opcional("SubTotal"),
        descuento=opcional("Descuento"),
        iva_8=opcional("IVA 8%"),
        total_trasladados=opcional("Total Trasladados"),
        nivel_educativo=opcional("Nivel Educativo"),
        conceptos=opcional("Conceptos"),
        lugar_expedicion=opcional("LugarDeExpedicion"),
    )


# ---------------------------------------------------------------- PagosRecibidasXML
@dataclass
class ColumnasRep:
    fecha_pago: str
    estado: str
    bancarizado: str
    base_16: str
    iva_16: str
    rfc: str
    rel_uuid: Optional[str] = None
    es_duplicado: Optional[str] = None
    moneda_documento: Optional[str] = None       # MonedaDR
    equivalencia: Optional[str] = None           # EquivalenciaDR
    moneda_pago: Optional[str] = None            # MonedaP
    metodo_documento: Optional[str] = None       # MetodoDePagoDR
    base_8: Optional[str] = None
    iva_8: Optional[str] = None
    base_exento: Optional[str] = None
    base_cero: Optional[str] = None
    lugar_expedicion: Optional[str] = None
    retenciones_iva: List[str] = field(default_factory=list)


def columnas_de_rep(encabezados):
    obligatoria = lambda *c: buscar_columna_obligatoria(encabezados, *c)
    opcional = lambda nombre: buscar_columna_opcional(encabezados, nombre)
    return ColumnasRep(
        fecha_pago=obligatoria("FechaPago"),
        estado=obligatoria("Estado"),
        bancarizado=obligatoria("BancarizadoP", "Bancarizado"),
        base_16=obligatoria("IVA 16 Base", "TotalTrasladosBaseIVA16"),
        iva_16=obligatoria("IVA 16 Importe", "TotalTrasladosImpuestoIVA16"),
        rfc=obligatoria("RFC Emisor CFDI", "RFC Emisor"),
        rel_uuid=opcional("RelUUID"),
        es_duplicado=opcional("EsDuplicado"),
        moneda_documento=opcional("MonedaDR"),
        equivalencia=opcional("EquivalenciaDR"),
        moneda_pago=opcional("MonedaP"),
        metodo_documento=opcional("MetodoDePagoDR"),
        base_8=opcional("IVA 8 Base"),
        iva_8=opcional("IVA 8 Importe"),
        base_exento=opcional("IVA Excento Base"),
        base_cero=opcional("IVA Cero Base"),
        lugar_expedicion=opcional("Lugar de Expedicion CFDI"),
        retenciones_iva=columnas_de_retencion_iva(encabezados),
    )


def valor(fila, columna):
    """Lee la celda de una columna opcional; si la columna no existe -> None."""
    return fila.get(columna) if columna else None
