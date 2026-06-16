# DIOT Generator — Registro de Facturación 2026v5

Generador de archivos **DIOT** (Declaración Informativa de Operaciones con Terceros) para el SAT, basado en `Registro de Facturación 2026v5.xlsb`. Lee la hoja en **modo lectura** (nunca modifica); escribe solo el archivo `.txt` con el formato de 54 campos que el SAT requiere.

## Características

- ✅ **Lógica de flujo de caja** (paid-basis): PUE bancarizadas, efectivo ≤ $2,000, y REP (PPD pagados en el mes)  
- ✅ **Redondeo Art. 20 CFF**: 0.01–0.50 → abajo, 0.51–0.99 → arriba  
- ✅ **Invariante SAT**: Garantiza `round(base×0.16) ≥ iva_acreditable` por renglón (evita errores del SAT)  
- ✅ **RFC/año auto-detectados** desde la ruta del archivo  
- ✅ **Múltiples RFC**: cada RFC puede tener su carpeta; la config se persiste por RFC  
- ✅ **Nombre de archivo**: `MM. Abbr YYYY  N DIOT Declaración.txt` (normal) o `… C1/C2/… DIOT Declaración.txt` (complementarias)  
- ✅ **Tkinter Save-As**: elige dónde guardar; recuerda la última carpeta usada  
- ✅ **100% Python, sin dependencias raras**: pyxlsb, pandas  

## Requisitos

- **Python 3.7+**  
- `pip install pyxlsb pandas`  
- El archivo **`Registro de Facturación 2026v5.xlsb`** en la misma carpeta  

## Instalación rápida

```bash
git clone https://github.com/<tu-usuario>/diot-generator.git
cd diot-generator
pip install -r requirements.txt
```

## Uso

### CLI directo (sin diálogo Save-As)
```bash
python diot_generator.py 5 N "C:\output\mayo.txt"
# o con rutas relativas
python diot_generator.py 5 C1
```

### Con diálogo gráfico (recomendado)
```bash
# En Windows:
Generar_DIOT.bat

# En Linux/Mac:
python diot_generator.py 5
```
El script abrirá un Save-As dialog; elige dónde guardar y confirma.

### Argumentos
```
python diot_generator.py <mes 1-12> [N|C1|C2|...] [ruta_salida.txt]

mes:        1-12 (enero-diciembre)
tipo:       N=normal (default), C1/C2/.../Cn=complementaria
ruta:       ruta.txt (opcional; si no se da, abre Save-As)
```

## Salida

Archivo `.txt` con 54 campos separados por `|` (UTF-8, CRLF):
```
04|85|RFC|...|Base 16%|...|IVA acreditable|...|
```

Cada línea es un proveedor. Los campos se llenan solo si hay datos:
- Campo 1: `04` (tipo tercero nacional)  
- Campo 2: `85` Otros (default; ver `diot_rfc_op.csv` para `03`/`06`)  
- Campo 3: RFC del proveedor  
- Campo 12: Base 16%  
- Campo 22: IVA acreditable  
- Campo 48: IVA retenido (si aplica)  

## Reconciliación

El script imprime los totales por bucket y los compara con tu hoja `Declaracion`:

```
Bucket A (PUE banc)     base   3,451,475.08   IVA    552,236.01
Bucket B (efectivo)     base       33,307.16   IVA      5,329.15
Bucket C (REP/PPD)      base        3,139.17   IVA        502.26
──────────────────────────────────────────────────────────
TXT TOTAL (CFF redon.)  base      3,487,919    IVA       558,073
```

Compara estos números contra `Declaracion` C34 (VALOR DE ACTOS 16%) y C35 (TOTAL IVA ACREDITABLE). Diferencias < 0.01% caen bajo **importancia relativa** y el SAT las acepta.

## Clasificación de operaciones (tipo)

Por defecto, todos los proveedores se clasifican como `85` (Otros). Para cambiar:

1. Crea un archivo `diot_rfc_op.csv` en la misma carpeta:
   ```csv
   RFC,Op
   TME840315KT6,03
   RDI841003QJ4,06
   ```
2. Ejecuta el script; leerá el CSV y aplicará las clasificaciones.

Ver `diot_rfc_op.csv.example` para una plantilla.

## Persistencia (RFC/carpeta)

El script recuerda la última carpeta donde guardaste la DIOT para cada RFC, en:
```
%USERPROFILE%\.diot_config.json  (Windows)
~/.diot_config.json              (Linux/Mac)
```

Cada vez que guardas, se actualiza la ruta. Así, si tienes varios RFC, cada uno mantiene su propia preferencia de carpeta.

## Lógica de agregación (buckets)

Espeja exactamente la hoja `Resumen` de tu registro:

**Bucket A (PUE bancarizadas deducibles)**  
- Fuente: `RecibidasXML`  
- Filtro: Fecha en mes, Tipo=Factura, Estado=Vigente, Metodo comienza con "PUE", Bancarizado="Sí", UsoCFDI ∈ {G01, G03}  
- Columnas: base/IVA → Resumen P/Q/R  

**Bucket B (Efectivo ≤ $2,000)**  
- Fuente: `RecibidasXML`  
- Filtro: igual a A, pero FormaDePago="01", Importe Neto ≤ 2,000, Combustible ≠ "Sí"  
- Columnas: base/IVA → Resumen T/U/V/W  

**Bucket C (REP, PPD pagados en el mes)**  
- Fuente: `PagosRecibidasXML`  
- Filtro: FechaPago en mes, Estado=Vigente, BancarizadoP="Sí"  
- Columnas: base/IVA → Resumen X/Y/Z  

**Base 16% = IVA ÷ 0.16** (no subtotal directo). Esto asegura que:
1. La base sea consistente con el IVA que SAT validará.
2. El IEPS (en Telmex, etc.) quede incluido correctamente en la base de 16%.
3. Líneas exentas/0% no contaminen la base de 16%.

## Errores del SAT y soluciones

### "El importe del 'IVA acreditable'… no puede ser mayor al 'IVA pagado'…"

**Causa**: Antes se usaba redondeo half-up (Python default), que podía causar `round(iva) > round(base×0.16)`.

**Solución (v2+)**: Redondeo Art. 20 CFF + invariante SAT. Si tras redondear, `cff_round(base×0.16) < cff_round(iva)`, se sube la base 1 peso. Verificado: 0 violaciones en 119 proveedores.

### "¿Por qué la base no coincide exactamente con la Declaracion?"

El script usa `base = iva / 0.16` tras redondear. Esto difiere del subtotal directo porque:
- Telmex y otros llevan IEPS: el IVA es sobre subtotal+IEPS, no solo subtotal.
- Líneas exentas/0% en multiparte no cuentan en la base de 16%.

Compara contra C34 (VALOR DE ACTOS = IVA/0.16), no contra el subtotal crudo. Diffs < 0.01% = importancia relativa.

## Requisitos del sistema

| Sistema | Python | Notas |
|---|---|---|
| Windows | 3.7+ | Funciona en cmd, PowerShell. Doble-clic en `Generar_DIOT.bat` |
| Linux | 3.7+ | Instala `python3-tk` para tkinter: `sudo apt-get install python3-tk` |
| macOS | 3.7+ | tkinter incluido; usa `python3 diot_generator.py 5` |

## Desarrollo

El script está estructurado para Boardflare (Excel-Python). La función `build_diot(rec_headers, rec_rows, pag_headers, pag_rows, year, month, op_map)` es pura (sin I/O); puedes llamarla desde Boardflare para una vista previa de totales.

Para editar la lógica de buckets, mira la función `build_diot()` alrededor de la línea 120.

## Licencia

MIT. Úsalo, modifícalo, comparte. Ver [LICENSE](LICENSE).

## Autor & Créditos

Generado para GOOH841231EPA (Hugo Edgar González Orozco).  
Con asistencia de Claude (Anthropic).

---

¿Preguntas? Abre un **issue** en GitHub. ¿Mejoras? Haz un **pull request**.
