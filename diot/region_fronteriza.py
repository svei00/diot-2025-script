"""
Región del estímulo fiscal de IVA 8%.

Decisión del despacho (P2, 2026-11-11): todo el IVA 8% se reporta como
Región Fronteriza NORTE (campos 8 y 18 del TXT). En octubre 2026 todos los
proveedores con 8% expiden desde Mexicali, B.C. (CP 21xxx) o San Luis Río
Colorado, Son. (CP 84xxx).

Para no equivocarnos en silencio si un día aparece un proveedor de la
frontera SUR, se revisa el código postal de expedición y se avisa. Los
prefijos de CP (dos primeros dígitos) de los estados con municipios en la
región fronteriza sur son:
    24 Campeche, 29-30 Chiapas, 77 Quintana Roo, 86 Tabasco.
"""

from diot.valores_celda import texto

PREFIJOS_CP_FRONTERA_SUR = {"24", "29", "30", "77", "86"}


def es_cp_de_la_frontera_sur(codigo_postal):
    return texto(codigo_postal)[:2] in PREFIJOS_CP_FRONTERA_SUR


def aviso_si_parece_frontera_sur(rfc, codigo_postal):
    """Texto de aviso o None."""
    if not es_cp_de_la_frontera_sur(codigo_postal):
        return None
    return ("%s trae IVA 8%% y expide desde CP %s (¿frontera SUR?). El TXT lo "
            "reporta como frontera NORTE: revisar." % (rfc, texto(codigo_postal)))
