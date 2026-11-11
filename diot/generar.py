"""
Orquestación con I/O (la usan la línea de comandos y la ventana):

    leer libro -> build_diot() -> pedir dónde guardar -> escribir TXT
               -> leer Resumen y Declaracion -> texto de conciliación

build_diot() se queda puro; aquí vive todo lo que abre archivos.
"""

import os
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, List, Optional

from diot import archivo_salida as salida
from diot.conciliacion import summary
from diot.construir import (LIMITE_EFECTIVO_POR_OMISION, USOS_DEDUCIBLES_POR_OMISION,
                            ResultadoDiot, build_diot)
from diot.declaracion_libro import leer_declaracion
from diot.lectura_libro import EXTENSIONES_LIBRO, read_control, read_grid, read_sheet
from diot.resumen_libro import leer_resumen


@dataclass
class Generacion:
    ruta_txt: Optional[str]
    renglones: List[list]
    resultado: ResultadoDiot
    anio: int
    rfc: str
    mes: int
    tipo: str
    limite_efectivo: float
    resumen: Optional[dict] = None
    declaracion: Optional[dict] = None
    errores_de_lectura: List[str] = field(default_factory=list)

    def texto_de_conciliacion(self):
        texto = summary(self.ruta_txt, self.renglones, self.resultado, self.anio, self.rfc,
                        salida.MESES_LARGOS[self.mes], self.tipo, self.resumen,
                        self.declaracion, self.limite_efectivo)
        if self.errores_de_lectura:
            texto += "\n" + "\n".join(" OJO: %s" % e for e in self.errores_de_lectura)
        return texto


def validar_extension(libro):
    extension = os.path.splitext(str(libro))[1].lower()
    if extension not in EXTENSIONES_LIBRO:
        raise ValueError("Formato no soportado: '%s'.\nUsa un libro %s"
                         % (extension or libro, ", ".join(EXTENSIONES_LIBRO)))


def leer_hoja_para_conciliar(libro, hoja, interpretar, errores):
    """Resumen/Declaracion son solo para conciliar: si fallan, se avisa y se sigue.
    Una columna faltante en Resumen sí se muestra (no se lee nada en silencio)."""
    try:
        return interpretar(read_grid(libro, hoja))
    except Exception as error:                              # noqa: BLE001
        errores.append("%s no se pudo conciliar: %s" % (hoja, error))
        return None


def elegir_ruta_de_salida(rfc, libro, mes, anio, tipo, preguntar_ruta):
    carpeta_libro = os.path.dirname(os.path.abspath(libro))
    configuracion = salida.cargar_configuracion()
    ultima_carpeta = configuracion.get(rfc, {}).get("last_save_dir") or carpeta_libro
    if not os.path.isdir(ultima_carpeta):
        ultima_carpeta = carpeta_libro
    ruta = preguntar_ruta(salida.diot_filename(mes, anio, tipo), ultima_carpeta) \
        if preguntar_ruta else None
    if ruta:
        configuracion.setdefault(rfc, {})["last_save_dir"] = str(Path(ruta).parent)
        configuracion.setdefault(rfc, {})["last_libro"] = str(libro)
        salida.guardar_configuracion(configuracion)
    return ruta


def generate(libro, mes, tipo, out_path=None, ask_path=None):
    """Genera el TXT. Si el usuario cancela el "Guardar como", ruta_txt = None."""
    validar_extension(libro)
    carpeta = os.path.dirname(os.path.abspath(libro))
    rfc = salida.detectar_rfc(libro, os.getcwd())

    anio_control, limite_control, usos_control = read_control(libro)
    anio = anio_control or salida.detectar_anio(libro, os.getcwd()) or 2026
    limite = limite_control if limite_control is not None else LIMITE_EFECTIVO_POR_OMISION
    usos = usos_control or USOS_DEDUCIBLES_POR_OMISION

    encabezados_rec, filas_rec = read_sheet(libro, "RecibidasXML")
    encabezados_pag, filas_pag = read_sheet(libro, "PagosRecibidasXML")
    renglones, resultado = build_diot(encabezados_rec, filas_rec, encabezados_pag, filas_pag,
                                      anio, mes, salida.cargar_tipos_operacion(carpeta),
                                      limite_efe=limite, usos_ded=usos)

    errores: List[Any] = []
    resumen = leer_hoja_para_conciliar(libro, "Resumen", lambda g: leer_resumen(g, mes), errores)
    declaracion = leer_hoja_para_conciliar(libro, "Declaracion", leer_declaracion, errores)

    ruta = out_path or elegir_ruta_de_salida(rfc, libro, mes, anio, tipo, ask_path)
    if ruta:
        salida.escribir_txt(ruta, renglones)
    return Generacion(ruta, renglones, resultado, anio, rfc, mes, tipo, limite,
                      resumen, declaracion, errores)
