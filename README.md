# DIOT Generator — Registro de Facturación 2026v5

Generador de archivos **DIOT** (Declaración Informativa de Operaciones con Terceros) para el SAT, basado en el `Registro de Facturación 2026v5` en cualquiera de sus formatos: **`.xlsb`, `.xlsm` o `.xlsx`**. Lee la hoja en **modo lectura** (nunca modifica); escribe solo el archivo `.txt` con el formato de 54 campos que el SAT requiere.

## Características

- ✅ **Lógica de flujo de caja** (paid-basis): PUE bancarizadas, efectivo ≤ $2,000, y REP (PPD pagados en el mes)  
- ✅ **Redondeo Art. 20 CFF**: 0.01–0.50 → abajo, 0.51–0.99 → arriba  
- ✅ **Invariante SAT**: Garantiza `round(base×0.16) ≥ iva_acreditable` por renglón (evita errores del SAT)  
- ✅ **RFC/año auto-detectados** desde la ruta del archivo  
- ✅ **Múltiples RFC**: cada RFC puede tener su carpeta; la config se persiste por RFC  
- ✅ **Nombre de archivo**: `MM. Abbr YYYY  N DIOT Declaración.txt` (normal) o `… C1/C2/… DIOT Declaración.txt` (complementarias)  
- ✅ **Tkinter Save-As**: elige dónde guardar; recuerda la última carpeta usada  
- ✅ **Formatos aceptados**: `.xlsb`, `.xlsm`, `.xlsx` (también `.xltm`/`.xltx`)  
- ✅ **100% Python, sin dependencias raras**: pyxlsb, openpyxl, pandas  

## Requisitos

- **Python 3.7+**  
- `pip install pyxlsb openpyxl pandas` (`pyxlsb` solo se usa para `.xlsb`, `openpyxl` solo para `.xlsm`/`.xlsx`)  
- El **`Registro de Facturación 2026vN`** en `.xlsb`, `.xlsm` o `.xlsx` — puede vivir en cualquier carpeta; lo eliges con el file picker  

> **Nota sobre `.xlsm` / `.xlsx`:** de estos formatos se leen los **valores en caché** de las fórmulas, que es lo que Excel escribe al guardar. Si el archivo lo generó un programa (nunca pasó por Excel), las fórmulas no traen valor y el script te lo dice: ábrelo en Excel, guárdalo y vuelve a correrlo. El `.xlsb` no tiene ese detalle.

## Instalación rápida

```bash
git clone https://github.com/<tu-usuario>/diot-generator.git
cd diot-generator
pip install -r requirements.txt
```

## Uso

### GUI (recomendado — no tienes que teclear rutas)
```bash
Generar_DIOT.bat        # Windows: doble clic
python diot_generator.py   # cualquier SO
```
Se abre una ventana con:
- **Examinar…** → file picker para el libro `.xlsb` / `.xlsm` / `.xlsx` (recuerda el último que usaste)
- **Mes** → desplegable Enero–Diciembre
- **Tipo** → `N` normal o `C1`, `C2`, … complementaria
- **Generar DIOT** → corre y abre el Save-As para elegir dónde guardar

Al terminar imprime la conciliación contra la hoja `Resumen` en la misma ventana.

### CLI (para automatizar)
```bash
python diot_generator.py 6                       # Junio normal, abre Save-As
python diot_generator.py 6 C1                    # Junio complementaria 1
python diot_generator.py 6 N "C:\out\jun.txt"    # sin ningún diálogo
python diot_generator.py 6 N "jun.txt" --libro "D:\ruta\Registro 2026v5.xlsm"
```

### Argumentos
```
python diot_generator.py [mes 1-12] [N|C1|C2|...] [ruta_salida.txt] [--libro ruta.xlsb|.xlsm|.xlsx]

(sin argumentos)  abre la GUI
mes:        1-12 (enero-diciembre)
tipo:       N=normal (default), C1/C2/.../Cn=complementaria
ruta:       ruta.txt (opcional; si no se da, abre Save-As)
--libro:    ruta al libro .xlsb/.xlsm/.xlsx (opcional; si no se da, busca uno en la
            carpeta actual, prefiriendo .xlsb > .xlsm > .xlsx)
```

## Parámetros leídos de la hoja `Control`

No están hardcodeados — el script los lee del libro:

| Parámetro | Celda `Control` | Valor 2026 |
|---|---|---|
| Año fiscal | `Año fiscal` | 2026 |
| Límite pago en efectivo (**IVA incluido**) | `Límite pago en efectivo` | 2,000 |
| Usos CFDI deducibles | `Usos CFDI deducibles` | G01, G03 |

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

El script compara cada bucket contra la fila del mes en `Resumen` y marca `OK` o la diferencia. Junio 2026:

```
 Bucket                             Base            IVA   vs Resumen
 A (PUE bancarizada)          2488829.60      398212.74   OK
 B (efectivo <= 2000)           16982.13        2717.14   OK
 C (REP / PPD)                 133729.54       21396.72   OK
────────────────────────────────────────────────────────────────────
 TXT TOTAL (CFF redon.)          2639659         422327
 Ajuste base invariante SAT  +121 peso(s)  (trunc(base*0.16) >= IVA)
 REP descartado: AEA041220KM3   RelUUID=NO EXISTE  IVA 462.96
```

Si algún bucket sale `DIF`, el TXT no cuadra con el libro — revisa antes de subirlo. Diferencias < 0.01% caen bajo **importancia relativa** y el SAT las acepta.

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
- Filtro: Fecha Emision en mes, Tipo=Factura, Estado ≠ "Cancelado", Metodo comienza con "PUE", Bancarizado="Sí", UsoCFDI ∈ {G01, G03}  
- Columnas: base/IVA → Resumen O/P/Q/R  

**Bucket B (Efectivo ≤ $2,000)**  
- Fuente: `RecibidasXML`  
- Filtro: igual a A, pero FormaDePago="01", **`Total` ≤ 2,000** (el límite de `Control` es *IVA incluido*), Combustible ≠ "Sí"  
- Columnas: base/IVA → Resumen S/T/U/V/W  

**Bucket C (REP, PPD pagados en el mes)**  
- Fuente: `PagosRecibidasXML`  
- Filtro: FechaPago en mes, Estado ≠ "Cancelado", BancarizadoP="Sí", **`RelUUID` = "OK"**, `EsDuplicado` ≠ 1  
- Columnas: base/IVA → Resumen X/Y/Z  

> ⚠️ `Estado` es una columna **manual** (así lo dice `Resumen`: *"marcar Cancelado en columna Estado…"*).
> Vacío = vigente. Por eso el filtro es `≠ "Cancelado"` y **no** `== "Vigente"`: en `PagosRecibidasXML`
> el `Estado` viene en blanco y exigir `"Vigente"` tiraba **todos** los REP del mes.

> ⚠️ `RelUUID = "NO EXISTE"` significa que el CFDI que ese REP dice pagar **no está** en `RecibidasXML`.
> `Resumen` lo excluye, así que el script también — y te lo reporta al final para que lo investigues.

**Base 16% = IVA ÷ 0.16** (no subtotal directo). Esto asegura que:
1. La base sea consistente con el IVA que SAT validará.
2. El IEPS (en Telmex, etc.) quede incluido correctamente en la base de 16%.
3. Líneas exentas/0% no contaminen la base de 16%.

## Errores del SAT y soluciones

### "El importe del 'IVA acreditable'… no puede ser mayor al 'IVA pagado'…"

**Causa**: el validador del SAT recalcula el *IVA pagado* del renglón desde la base y **trunca**, no redondea. El acuse `ErroresCargaMasiva_…2026.005.txt` (mayo) rechazó 3 renglones donde `round(base×0.16) == iva` pero `trunc(base×0.16) < iva`:

| Línea | RFC | base | IVA acred. | `trunc(base×.16)` |
|---|---|---|---|---|
| 5 | ABU8605024N6 | 1254 | 201 | **200** ← SAT lo rechazó |
| 60 | GMM070201AA8 | 1860 | 298 | **297** |
| 89 | REF970701UK4 | 2979 | 477 | **476** |

**Solución (v3)**: `ensure_invariant()` sube la base al mínimo entero que cumple

```
(base × 16) // 100  >=  iva          # aritmética entera, sin errores de float
```

es decir `base >= ceil(iva × 25 / 4)`. Esa condición es **más estricta** que el redondeo, así que satisface las dos reglas posibles (truncar o redondear) sin tener que adivinar cuál usa el SAT.

**Costo**: +1 a +6 pesos por renglón. En junio 2026: **+121 pesos** sobre una base de 2,639,659 = **0.005%** → irrelevancia relativa. Verificado: **0 violaciones** en los 115 renglones, tanto truncando como redondeando.

### ⚠️ Bug en la hoja `Declaracion`: C29 = 0

`Declaracion!C29` ("Base IVA 16% REP recibidos") está en **0**, pero `C31` ("IVA acreditable prellenado") **sí** incluye el IVA de los REP. Eso deja el `VALOR DE ACTOS` (C30/C34) **subvaluado** por la base de los REP:

```
C34 (Monto a Capturar)   = 2,505,811.73
C35 (Total IVA acred.)   =   422,326.60
C34 × 0.16               =   400,929.88   <   422,326.60   ← el SAT rechazaría esto
```

El TXT sí es consistente (base 2,639,659 × 0.16 = 422,345 ≥ 422,327) porque el script suma la base de los REP. **Arregla C29** para que apunte a `Resumen!Y` (junio: 133,729.54) antes de capturar la declaración.

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
