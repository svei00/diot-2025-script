"""
Todo lo que toca disco fuera del libro: nombre del TXT, configuración
recordada por RFC, búsqueda del libro, catálogo de tipos de operación y
escritura del archivo.

Nombre del archivo (convención de la macro original):
    Normal:          "01. Ene 2026  N DIOT Declaración.txt"   (2 espacios antes de N)
    Complementaria:  "01. Ene 2026 C1 DIOT Declaración.txt"   (1 espacio antes de C#)
Los dos espacios son a propósito: al ordenar, la N queda arriba de C1, C2...
"""

import csv
import glob
import json
import os
import re
from pathlib import Path

from diot.lectura_libro import EXTENSIONES_LIBRO

ARCHIVO_TIPOS_OPERACION = "diot_rfc_op.csv"      # columnas RFC,Op (03 servicios / 06 arrend.)
RUTA_CONFIGURACION = Path.home() / ".diot_config.json"

MESES_ABREVIADOS = ["", "Ene", "Feb", "Mar", "Abr", "May", "Jun",
                    "Jul", "Ago", "Sep", "Oct", "Nov", "Dic"]
MESES_LARGOS = ["", "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
                "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]


def diot_filename(mes, anio, tipo):
    abreviatura = MESES_ABREVIADOS[mes]
    if tipo.upper() == "N":
        return "%02d. %s %d  N DIOT Declaración.txt" % (mes, abreviatura, anio)
    return "%02d. %s %d %s DIOT Declaración.txt" % (mes, abreviatura, anio, tipo.upper())


def detectar_rfc(*rutas):
    for ruta in rutas:
        encontrado = re.search(r'\b([A-ZÑ&]{3,4}\d{6}[A-Z\d]{3})\b', str(ruta or ""))
        if encontrado:
            return encontrado.group(1)
    return "RFC"


def detectar_anio(*rutas):
    for ruta in rutas:
        encontrado = re.search(r'[\\/](20\d{2})[\\/]', str(ruta or ""))
        if encontrado:
            return int(encontrado.group(1))
    return None


def cargar_configuracion():
    if RUTA_CONFIGURACION.exists():
        try:
            return json.loads(RUTA_CONFIGURACION.read_text(encoding="utf-8"))
        except Exception:
            return {}
    return {}


def guardar_configuracion(configuracion):
    try:
        RUTA_CONFIGURACION.write_text(
            json.dumps(configuracion, indent=2, ensure_ascii=False), encoding="utf-8")
    except Exception:
        pass


def buscar_libro(carpeta="."):
    """Primero por nombre, luego cualquiera. Prioridad .xlsb > .xlsm > .xlsx."""
    for patron in ("Registro de Facturaci*v*%s", "*%s"):
        for extension in EXTENSIONES_LIBRO:
            encontrados = [ruta for ruta in glob.glob(os.path.join(carpeta, patron % extension))
                           if not os.path.basename(ruta).startswith("~$")]
            if encontrados:
                return sorted(encontrados, reverse=True)[0]
    return None


def cargar_tipos_operacion(carpeta):
    """diot_rfc_op.csv -> {RFC: tipo de operación}. Sin archivo -> {}."""
    tipos, ruta = {}, os.path.join(carpeta, ARCHIVO_TIPOS_OPERACION)
    if os.path.exists(ruta):
        with open(ruta, newline="", encoding="utf-8-sig") as archivo:
            for fila in csv.DictReader(archivo):
                if fila.get("RFC"):
                    tipos[fila["RFC"].strip()] = str(fila.get("Op", "85")).strip()
    return tipos


def escribir_txt(ruta, renglones):
    """UTF-8, campos separados por '|', fin de línea CRLF."""
    with open(ruta, "w", encoding="utf-8", newline="") as archivo:
        for renglon in renglones:
            archivo.write("|".join(renglon) + "\r\n")
