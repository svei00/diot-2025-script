"""
Ventana (tkinter) para quien no quiere teclear rutas: elegir libro, mes y
tipo; "Generar DIOT" abre el "Guardar como" y muestra la conciliación.
"""

import os
import re

from diot import archivo_salida as salida
from diot.generar import generate


def preguntar_donde_guardar(nombre_sugerido, carpeta_inicial):
    from tkinter import filedialog
    return filedialog.asksaveasfilename(
        initialdir=carpeta_inicial, initialfile=nombre_sugerido, defaultextension=".txt",
        filetypes=[("Archivo DIOT (*.txt)", "*.txt")],
        title="Guardar DIOT como…") or None


def run_gui():
    import tkinter as tk
    from tkinter import filedialog, messagebox, scrolledtext, ttk

    configuracion = salida.cargar_configuracion()
    ultima = next(iter(configuracion.values()), {}) if configuracion else {}
    libro_sugerido = ultima.get("last_libro") or salida.buscar_libro(".") or ""

    raiz = tk.Tk()
    raiz.title("Generador DIOT — SAT")
    raiz.resizable(False, False)
    marco = ttk.Frame(raiz, padding=14)
    marco.grid()

    ttk.Label(marco, text="Libro de registro (.xlsb / .xlsm / .xlsx):").grid(row=0, column=0, sticky="w")
    v_libro = tk.StringVar(value=libro_sugerido)
    ttk.Entry(marco, textvariable=v_libro, width=64).grid(
        row=1, column=0, columnspan=2, sticky="we", pady=(0, 8))

    def examinar():
        carpeta = os.path.dirname(v_libro.get()) or "."
        ruta = filedialog.askopenfilename(
            title="Selecciona el Registro de Facturación",
            initialdir=carpeta if os.path.isdir(carpeta) else ".",
            filetypes=[("Libros de Excel (*.xlsb;*.xlsm;*.xlsx)", "*.xlsb *.xlsm *.xlsx"),
                       ("Todos", "*.*")])
        if ruta:
            v_libro.set(ruta)

    ttk.Button(marco, text="Examinar…", command=examinar).grid(row=1, column=2, padx=(8, 0), pady=(0, 8))

    ttk.Label(marco, text="Mes:").grid(row=2, column=0, sticky="w")
    v_mes = tk.StringVar(value=salida.MESES_LARGOS[1])
    ttk.Combobox(marco, textvariable=v_mes, values=salida.MESES_LARGOS[1:], state="readonly",
                 width=16).grid(row=3, column=0, sticky="w", pady=(0, 8))

    ttk.Label(marco, text="Tipo:").grid(row=2, column=1, sticky="w")
    v_tipo = tk.StringVar(value="N")
    ttk.Combobox(marco, textvariable=v_tipo, values=["N", "C1", "C2", "C3", "C4", "C5"],
                 width=8).grid(row=3, column=1, sticky="w", pady=(0, 8))

    salida_texto = scrolledtext.ScrolledText(marco, width=96, height=26, font=("Consolas", 9))
    salida_texto.grid(row=5, column=0, columnspan=3, pady=(8, 0))

    def generar():
        libro = v_libro.get().strip()
        if not os.path.isfile(libro):
            messagebox.showerror("DIOT", "No encuentro el libro:\n%s" % libro)
            return
        mes = salida.MESES_LARGOS.index(v_mes.get())
        tipo = v_tipo.get().strip().upper() or "N"
        if tipo != "N" and not re.fullmatch(r"C\d+", tipo):
            messagebox.showerror("DIOT", "Tipo inválido. Usa N o C1, C2, …")
            return
        salida_texto.delete("1.0", "end")
        salida_texto.insert("end", "Leyendo %s …\n" % libro)
        raiz.update()
        try:
            generacion = generate(libro, mes, tipo, ask_path=preguntar_donde_guardar)
        except Exception as error:                          # noqa: BLE001
            salida_texto.insert("end", "\nERROR: %s\n" % error)
            messagebox.showerror("DIOT", str(error))
            return
        if not generacion.ruta_txt:
            salida_texto.insert("end", "Cancelado.\n")
            return
        salida_texto.delete("1.0", "end")
        salida_texto.insert("end", generacion.texto_de_conciliacion())

    ttk.Button(marco, text="Generar DIOT", command=generar).grid(row=3, column=2, sticky="e", pady=(0, 8))
    raiz.mainloop()
