#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Generador del TXT de carga masiva de la DIOT 2025+ desde el
"Registro de Facturación 2026vN" (.xlsb / .xlsm / .xlsx).

Lee el libro en SOLO LECTURA y escribe únicamente el .txt. La lógica vive
en el paquete `diot/` (ver diot/__init__.py para el mapa de archivos).

Uso:
    python diot_generator.py                      # abre la ventana (recomendado)
    python diot_generator.py --gui
    python diot_generator.py 6                    # Junio normal, "Guardar como"
    python diot_generator.py 6 C1                 # Junio complementaria 1
    python diot_generator.py 6 N "ruta\\out.txt"  # sin diálogos
    python diot_generator.py 6 N "out.txt" --libro "ruta\\Registro.xlsm"

Boardflare / Excel-Python (función pura, sin I/O):
    from diot_generator import build_diot, read_sheet
    rh, rr = read_sheet(path, "RecibidasXML")
    ph, pr = read_sheet(path, "PagosRecibidasXML")
    renglones, resultado = build_diot(rh, rr, ph, pr, year, month, op_map={})
"""

import os
import re
import sys

from diot.archivo_salida import buscar_libro, diot_filename        # noqa: F401
from diot.construir import build_diot                              # noqa: F401
from diot.generar import generate
from diot.lectura_libro import read_sheet                          # noqa: F401
from diot.redondeo import base_minima_para_iva, redondear_cff      # noqa: F401


def separar_opcion_libro(argumentos):
    """Quita '--libro ruta' de la lista y regresa (argumentos, ruta o None)."""
    if "--libro" not in argumentos:
        return argumentos, None
    posicion = argumentos.index("--libro")
    libro = argumentos[posicion + 1]
    return argumentos[:posicion] + argumentos[posicion + 2:], libro


def preguntar_ruta_en_consola(nombre_sugerido, carpeta_inicial):
    """'Guardar como' de tkinter; sin pantalla, guarda junto al libro."""
    try:
        import tkinter as tk
        from tkinter import filedialog
        raiz = tk.Tk()
        raiz.withdraw()
        raiz.attributes("-topmost", True)
        ruta = filedialog.asksaveasfilename(
            initialdir=carpeta_inicial, initialfile=nombre_sugerido, defaultextension=".txt",
            filetypes=[("Archivo DIOT (*.txt)", "*.txt")], title="Guardar DIOT como…")
        raiz.destroy()
        return ruta or None
    except Exception:                                      # noqa: BLE001
        return os.path.join(carpeta_inicial, nombre_sugerido)


def main():
    argumentos, libro = separar_opcion_libro(sys.argv[1:])
    if not argumentos or argumentos[0] in ("--gui", "-g"):
        from diot.ventana import run_gui
        run_gui()
        return

    mes = int(argumentos[0])
    tipo = argumentos[1].upper() if len(argumentos) > 1 else "N"
    ruta_fija = argumentos[2] if len(argumentos) > 2 else None
    if not 1 <= mes <= 12:
        print("Mes inválido (1-12).")
        sys.exit(1)
    if tipo != "N" and not re.fullmatch(r"C\d+", tipo):
        print("Tipo inválido. Usa N o C1, C2, …")
        sys.exit(1)

    libro = libro or buscar_libro(".")
    if not libro or not os.path.isfile(libro):
        print("No encontré el libro (.xlsb/.xlsm/.xlsx). "
              "Usa --libro \"ruta\\Registro.xlsm\" o corre sin argumentos para la ventana.")
        sys.exit(1)

    print("Leyendo %s ..." % libro)
    generacion = generate(libro, mes, tipo, out_path=ruta_fija, ask_path=preguntar_ruta_en_consola)
    if not generacion.ruta_txt:
        print("Cancelado por usuario.")
        sys.exit(0)
    print(generacion.texto_de_conciliacion())


if __name__ == "__main__":
    main()
