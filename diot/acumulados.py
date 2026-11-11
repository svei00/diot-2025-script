"""
Estructuras para sumar importes.

`ImportesPorTasa` es la "bolsa" de importes de una operación (una factura o
un REP), de un proveedor (suma de sus operaciones) o de un bucket completo.
Todas usan las mismas claves, así que se suman igual en los tres niveles.

    base_16 / iva_16     tasa general 16%
    base_8  / iva_8      estímulo región fronteriza (8%)
    exento               actos por los que no se paga IVA (base)
    tasa_cero            actos a tasa 0% (base)
    iva_retenido         IVA que retuvimos al proveedor
"""

from dataclasses import dataclass, fields


@dataclass
class ImportesPorTasa:
    base_16: float = 0.0
    iva_16: float = 0.0
    base_8: float = 0.0
    iva_8: float = 0.0
    exento: float = 0.0
    tasa_cero: float = 0.0
    iva_retenido: float = 0.0

    def sumar(self, otro):
        """Suma, campo por campo, los importes de `otro` a esta bolsa."""
        for campo in fields(self):
            setattr(self, campo.name, getattr(self, campo.name) + getattr(otro, campo.name))


@dataclass
class TotalesDeBucket:
    """Totales de un bucket (A, B o C): cuántos documentos y cuánto suman."""
    numero_documentos: int = 0
    importes: ImportesPorTasa = None

    def __post_init__(self):
        if self.importes is None:
            self.importes = ImportesPorTasa()

    def agregar(self, importes):
        self.numero_documentos += 1
        self.importes.sumar(importes)
